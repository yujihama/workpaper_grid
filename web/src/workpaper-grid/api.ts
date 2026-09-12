import type { LogEntry } from "./log";
import type { Merge, Pos } from "./types";

// サーバ側の契約。他アプリに載せるときはこのインターフェースを満たすものを渡せばよい
export type AiItem = Pos & { current: string; comment: string };
export type AiOp = Pos & { input: string };

export interface WorkbookApi {
  /** xlsx を IronCalc 形式に変換して受け取る。merges は wasm から取れないのでサーバが付ける */
  open(file: File): Promise<{ bytes: Uint8Array; sessionId: string; merges: Merge[] }>;
  /** ブラウザ側の差分をサーバへ送り、変更履歴シートと結合を反映した xlsx を受け取る */
  sync(sessionId: string, diffs: Uint8Array, merges: Merge[]): Promise<Blob>;
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
    async sync(sessionId, diffs, merges) {
      const fd = new FormData();
      fd.append("file", new Blob([diffs as BlobPart]), "diffs.bin");
      fd.append("merges", JSON.stringify(merges.map((g) => [g.sheet, g.r1, g.c1, g.r2, g.c2])));
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
