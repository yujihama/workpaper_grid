import type { LogEntry } from "./log";
import { colNum, type Assets, type Merge, type Pos, type StructOp } from "./types";

// サーバ側の契約。他アプリに載せるときはこのインターフェースを満たすものを渡せばよい
export type AiItem = Pos & { current: string; comment: string };
export type AiOp = Pos & { input: string };

export interface WorkbookApi {
  /** xlsx を IronCalc 形式に変換して受け取る。merges は wasm から取れないのでサーバが付ける */
  open(file: File): Promise<{ bytes: Uint8Array; sessionId: string; merges: Merge[] }>;
  /** 元ファイルにあって IronCalc が持たないもの（メモ・画像・グラフ・図形・リンク・入力規則）を表示用に受け取る */
  assets(sessionId: string): Promise<Assets>;
  /** ブラウザ側の差分をサーバへ送り、元ファイルを土台に図形・メモ等を残した xlsx を受け取る */
  sync(sessionId: string, diffs: Uint8Array, merges: Merge[], origNames: (string | null)[], structOps: StructOp[]): Promise<Blob>;
  /** 編集ログの追記（失敗しても操作は止めない） */
  appendLog(sessionId: string, entries: LogEntry[]): Promise<void>;
  /** コメント付きセルをまとめて AI に渡し、セルごとの更新命令を受け取る */
  ai(items: AiItem[]): Promise<{ ops: AiOp[] }>;
}

// server/workpaper.py に対応する HTTP 実装
export function createHttpApi(base = "/api"): WorkbookApi {
  return {
    async open(file) {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${base}/open`, { method: "POST", body: fd });
      if (!res.ok) throw new Error(`開けませんでした（${res.status}）: ${(await res.text()).slice(0, 200)}`);
      const merges = (JSON.parse(res.headers.get("X-Merges") ?? "[]") as number[][]).map(([sheet, r1, c1, r2, c2]) => ({ sheet, r1, c1, r2, c2 }));
      return { bytes: new Uint8Array(await res.arrayBuffer()), sessionId: res.headers.get("X-Session-Id") ?? "", merges };
    },
    async assets(sessionId) {
      const res = await fetch(`${base}/assets/${sessionId}`);
      if (!res.ok) throw new Error(`付随情報の取得に失敗しました（${res.status}）`);
      const j = await res.json();
      // サーバは refs を A1 形式で返す。表示用に実座標へ正規化する
      const cell = (ref: string) => {
        const m = /^\$?([A-Z]+)\$?(\d+)/.exec(ref)!;
        return { row: Number(m[2]), col: colNum(m[1]) };
      };
      const range = (ref: string) => {
        const [a, b] = ref.split(":");
        const p = cell(a),
          q2 = cell(b ?? a);
        return { r1: p.row, c1: p.col, r2: q2.row, c2: q2.col };
      };
      return {
        notes: j.notes,
        images: j.images,
        shapes: [...j.charts, ...j.shapes],
        hyperlinks: j.hyperlinks.map((h: { sheet: number; ref: string; target: string }) => ({ sheet: h.sheet, ...cell(h.ref), target: h.target })),
        validations: j.validations.flatMap((v: { sheet: number; sqref: string; type: string; formula1: string }) =>
          v.sqref.split(/\s+/).map((r: string) => ({ sheet: v.sheet, ...range(r), type: v.type, formula1: v.formula1 }))
        ),
        defaults: j.defaults ?? [],
      };
    },
    async sync(sessionId, diffs, merges, origNames, structOps) {
      const fd = new FormData();
      fd.append("file", new Blob([diffs as BlobPart]), "diffs.bin");
      fd.append("merges", JSON.stringify(merges.map((g) => [g.sheet, g.r1, g.c1, g.r2, g.c2])));
      fd.append("orig_names", JSON.stringify(origNames));
      fd.append("struct_ops", JSON.stringify(structOps));
      const res = await fetch(`${base}/sync/${sessionId}`, { method: "POST", body: fd });
      if (!res.ok) throw new Error(`保存に失敗しました（${res.status}）`);
      return res.blob();
    },
    async appendLog(sessionId, entries) {
      await fetch(`${base}/log/${sessionId}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ entries }),
      }).catch(() => {});
    },
    async ai(items) {
      const res = await fetch(`${base}/ai`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ items }) });
      if (!res.ok) throw new Error(`AI 呼び出しに失敗しました（${res.status}）`);
      return res.json();
    },
  };
}
