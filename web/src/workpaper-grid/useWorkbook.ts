import { useEffect, useRef, useState } from "react";
import init, { Model, type Area, type Clipboard } from "@ironcalc/wasm";
import wasmUrl from "@ironcalc/wasm/wasm_bg.wasm?url";
import type { WorkbookApi } from "./api";
import { describe, type LogAction, type LogCell, type LogEntry, type LogInput } from "./log";
import { colDrop, colMove, colShift, identity, remap, remapSheets, rowDrop, rowMove, rowShift, sheetDrop, sheetMove, type Move } from "./remap";
import { colName, colRange, EMPTY_SNAP, key, rowRange, same, type CellEdit, type Merge, type Pos, type Proposal, type SheetInfo, type Snap, type StructOp } from "./types";
import { pxToIcH, pxToIcW } from "./units";

// 履歴1件 = モデル側のundo回数 + UI状態の前後 + ログ
type Entry = { id: number; steps: number; before: Snap; after: Snap; log: LogEntry };
export type FocusReq = Pos & { n: number };
export type InternalClip = { clip: Clipboard; sheet: number; rows: string[][]; cut: boolean };

/**
 * 調書ブック1冊の状態と操作。座標はすべて 1 始まりの実番号で受ける。
 * 画面（グリッド）はこの hook の戻り値だけを使い、モデルを直接触るのは読み取りに限る。
 */
export function useWorkbook(api: WorkbookApi) {
  const [ready, setReady] = useState(false);
  const modelRef = useRef<Model | null>(null);
  const [sid, setSid] = useState<string | null>(null);
  const sidRef = useRef<string | null>(null);
  const [fileName, setFileName] = useState("");
  const [sheet, setSheet] = useState(0);
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [tick, setTick] = useState(0); // モデルが変わったことを画面に伝える
  const [selected, setSelected] = useState<{ row: number; col: number } | null>(null);
  const [focusReq, setFocusReq] = useState<FocusReq | null>(null);
  const [busy, setBusy] = useState(false); // 開く・保存中
  const [regen, setRegen] = useState(false); // AI再生成中
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [ui, setUi] = useState<Snap>(EMPTY_SNAP);
  const uiRef = useRef(ui);
  const undoStack = useRef<Entry[]>([]);
  const redoStack = useRef<Entry[]>([]);
  const entrySeq = useRef(0);
  const clip = useRef<InternalClip | null>(null); // IronCalc の書式付きクリップボード
  const explicitRows = useRef(new Set<string>()); // このセッションで高さを指定した行（既定扱いの補正をしない）

  useEffect(() => {
    init(wasmUrl).then(() => setReady(true));
  }, []);

  const mm = () => modelRef.current!;
  const syncSheets = () => {
    const m = modelRef.current;
    if (m) setSheets(m.getWorksheetsProperties().map((w) => ({ name: w.name, state: w.state })));
  };
  const refresh = () => {
    setTick((t) => t + 1);
    syncSheets();
  };
  const applyUi = (s: Snap) => {
    uiRef.current = s;
    setUi(s);
  };
  const sheetName = (i: number) => sheets[i]?.name ?? "";
  const addr = (p: Pos) => `${sheetName(p.sheet)}!${colName(p.col)}${p.row}`;
  const mergeAt = (s: number, r: number, c: number) => ui.merges.find((g) => g.sheet === s && r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
  const rangeOf = (cells: [number, number][], s = sheet) => {
    const rs = cells.map((x) => x[0]),
      cs = cells.map((x) => x[1]);
    const a = `${colName(Math.min(...cs))}${Math.min(...rs)}`;
    const b = `${colName(Math.max(...cs))}${Math.max(...rs)}`;
    return `${sheetName(s)}!${a === b ? a : `${a}:${b}`}`;
  };
  const areaLabel = (a: Area) =>
    a.width === 1 && a.height === 1
      ? `${colName(a.column)}${a.row}`
      : `${colName(a.column)}${a.row}:${colName(a.column + a.width - 1)}${a.row + a.height - 1}`;
  const guarded = (fn: () => void) => {
    try {
      fn();
    } catch (e) {
      console.error(e);
      alert(String(e));
    }
  };

  // ---- ログ：ブラウザに保持しつつ、都度サーバへ追記（失敗しても操作は止めない） ----
  const send = (entries: LogEntry[]) => (sidRef.current ? api.appendLog(sidRef.current, entries) : Promise.resolve());
  const record = (input: LogInput, seq?: number): LogEntry => {
    const e: LogEntry = { seq: seq ?? ++entrySeq.current, ts: new Date().toISOString(), ...input };
    setLogs((l) => [...l, e]);
    void send([e]);
    return e;
  };

  // ---- 履歴：1操作 = モデルに加えた変更（steps回のundoで戻る）＋UI状態の前後＋ログ ----
  const commit = (steps: number, next: (s: Snap) => Snap, log: LogInput) => {
    const before = uiRef.current;
    const after = next(before);
    const id = ++entrySeq.current;
    undoStack.current.push({ id, steps, before, after, log: record(log, id) });
    redoStack.current = [];
    applyUi(after);
    modelRef.current?.evaluate();
    refresh();
    return id;
  };
  const flip = (cells: LogCell[]) => cells.map((c) => ({ ...c, before: c.after, after: c.before }));
  const undo = () => {
    const m = modelRef.current;
    const e = undoStack.current.pop();
    if (!m || !e) return;
    for (let i = 0; i < e.steps && m.canUndo(); i++) m.undo();
    redoStack.current.push(e);
    applyUi(e.before);
    record({ actor: "user", action: "undo", sheet: e.log.sheet, range: e.log.range, cells: flip(e.log.cells), ref: e.id });
    m.evaluate();
    refresh();
  };
  const redo = () => {
    const m = modelRef.current;
    const e = redoStack.current.pop();
    if (!m || !e) return;
    for (let i = 0; i < e.steps && m.canRedo(); i++) m.redo();
    undoStack.current.push(e);
    applyUi(e.after);
    record({ actor: "user", action: "redo", sheet: e.log.sheet, range: e.log.range, cells: e.log.cells, ref: e.id });
    m.evaluate();
    refresh();
  };
  // 複数ステップの操作を途中失敗も含めて1件にする（失敗時は適用済み分を undo）
  const atomic = (steps: () => number, next: (s: Snap) => Snap, log: LogInput) => {
    let done = 0;
    try {
      done = steps();
      commit(done, next, log);
    } catch (e) {
      for (let i = 0; i < done; i++) mm().undo();
      console.error(e);
      alert(String(e));
    }
  };

  // ---- ファイルを開く ----
  const open = async (f: File) => {
    setBusy(true);
    try {
      const { bytes, sessionId, merges } = await api.open(f);
      modelRef.current?.free();
      const m = Model.from_bytes(bytes, "en");
      m.evaluate();
      modelRef.current = m;
      sidRef.current = sessionId;
      setSid(sessionId);
      setFileName(f.name);
      const ws = m.getWorksheetsProperties().map((w) => ({ name: w.name, state: w.state }));
      setSheets(ws);
      setSheet(Math.max(0, ws.findIndex((w) => w.state === "visible"))); // 先頭が非表示なら最初の表示シートへ
      applyUi({ ...EMPTY_SNAP, merges, origNames: ws.map((w) => w.name) });
      undoStack.current = [];
      redoStack.current = [];
      entrySeq.current = 0;
      explicitRows.current.clear();
      clip.current = null;
      setLogs([]);
      setSelected(null);
      refresh();
      // 元ファイルにあって IronCalc が持たないもの（メモ・画像・グラフ・リンク・入力規則）は別途取り寄せて重ねて表示する
      api.assets(sessionId).then((assets) => {
        if (sidRef.current === sessionId) applyUi({ ...uiRef.current, assets });
      }).catch((e) => console.error(e));
    } catch (e) {
      alert(String(e));
    } finally {
      setBusy(false);
    }
  };

  // ---- 入力の適用（手入力・貼り付け・数式バー・Ctrl+Enter・置換の共通部分）。1回の操作を履歴1件にまとめる ----
  const applyInputs = (edits: CellEdit[], action: LogAction, note?: string) => {
    const m = modelRef.current;
    if (!m) return;
    const touched: [number, number][] = [];
    const cells: LogCell[] = [];
    for (const { row: r, col: c, value } of edits) {
      if (uiRef.current.locks.has(key(sheet, r, c))) continue;
      const before = m.getCellContent(sheet, r, c);
      if (before === value) continue;
      m.setUserInput(sheet, r, c, value);
      touched.push([r, c]);
      cells.push({ addr: addr({ sheet, row: r, col: c }), before, after: value });
    }
    if (!touched.length) return;
    commit(
      touched.length,
      (s) => {
        const marks = new Map(s.marks);
        for (const [r, c] of touched) marks.set(key(sheet, r, c), "changed");
        return { ...s, marks };
      },
      { actor: "user", action, sheet: sheetName(sheet), range: rangeOf(touched), cells, note }
    );
  };
  const clearCells = (cells: [number, number][]) => {
    const m = modelRef.current;
    if (!m) return;
    applyInputs(cells.filter(([r, c]) => m.getCellContent(sheet, r, c) !== "").map(([row, col]) => ({ row, col, value: "" })), "clear");
  };
  const markCells = (cells: [number, number][]) => (s: Snap) => {
    const marks = new Map(s.marks);
    for (const [r, c] of cells) marks.set(key(sheet, r, c), "changed");
    return { ...s, marks };
  };

  // ---- 行・列の挿入/削除/非表示/サイズ/移動 ----
  // 構造操作は Snap にも記録し（undo/redo で戻る）、保存時にサーバへ送って図形・メモ等の位置に再生する
  const structural = (action: LogAction, range: string, note: string, op: () => void, move: Move, rec?: Omit<StructOp, "sheet_orig_name">) =>
    guarded(() => {
      op();
      commit(
        1,
        (s) => {
          const next = remap(s, sheet, move);
          return rec ? { ...next, structOps: [...next.structOps, { ...rec, sheet_orig_name: s.origNames[sheet] ?? null }] } : next;
        },
        { actor: "user", action, sheet: sheetName(sheet), range, cells: [], note }
      );
    });
  const cnt = (n: number) => `${n} 件`;
  const insertRowsAt = (at: number, n: number) =>
    structural("insert_rows", rowRange(at, at + n - 1), cnt(n), () => mm().insertRows(sheet, at, n), rowShift(at, n), { kind: "insert_rows", at, n, delta: 0 });
  const deleteRowsAt = (at: number, n: number) =>
    structural("delete_rows", rowRange(at, at + n - 1), cnt(n), () => mm().deleteRows(sheet, at, n), rowDrop(at, n), { kind: "delete_rows", at, n, delta: 0 });
  const insertColsAt = (at: number, n: number) =>
    structural("insert_cols", colRange(at, at + n - 1), cnt(n), () => mm().insertColumns(sheet, at, n), colShift(at, n), { kind: "insert_cols", at, n, delta: 0 });
  const deleteColsAt = (at: number, n: number) =>
    structural("delete_cols", colRange(at, at + n - 1), cnt(n), () => mm().deleteColumns(sheet, at, n), colDrop(at, n), { kind: "delete_cols", at, n, delta: 0 });
  const hideRows = (a: number, b: number, hidden: boolean) =>
    structural(hidden ? "hide_rows" : "unhide_rows", rowRange(a, b), cnt(b - a + 1), () => mm().setRowsHidden(sheet, a, b, hidden), identity);
  const hideCols = (a: number, b: number, hidden: boolean) =>
    structural(hidden ? "hide_cols" : "unhide_cols", colRange(a, b), cnt(b - a + 1), () => mm().setColumnsHidden(sheet, a, b, hidden), identity);
  // 高さ・幅は Excel のピクセルで受け、IronCalc の単位に換算して入れる（units.ts）
  const setRowHeight = (a: number, b: number, h: number) => {
    for (let r = a; r <= b; r++) explicitRows.current.add(key(sheet, r, 0));
    structural("row_height", rowRange(a, b), `${Math.round(h)}px`, () => mm().setRowsHeight(sheet, a, b, pxToIcH(h)), identity);
  };
  const setColWidth = (a: number, b: number, w: number) =>
    structural("col_width", colRange(a, b), `${Math.round(w)}px`, () => mm().setColumnsWidth(sheet, a, b, pxToIcW(w)), identity);
  // 行 [a,b] を「実行 t の手前」へ移動（t は行番号）
  const moveRowsTo = (a: number, b: number, t: number) => {
    const n = b - a + 1;
    if (t >= a && t <= b + 1) return;
    const delta = t > b ? t - n - a : t - a;
    structural("move_rows", rowRange(a, b), `${t > b ? t - n : t} 行目へ`, () => mm().moveRows(sheet, a, n, delta), rowMove(a, n, delta), { kind: "move_rows", at: a, n, delta });
  };
  const moveColsTo = (a: number, b: number, t: number) => {
    const n = b - a + 1;
    if (t >= a && t <= b + 1) return;
    const delta = t > b ? t - n - a : t - a;
    structural("move_cols", colRange(a, b), `${colName(t > b ? t - n : t)} 列へ`, () => mm().moveColumns(sheet, a, n, delta), colMove(a, n, delta), { kind: "move_cols", at: a, n, delta });
  };

  // ---- 書式 ----
  const applyStyle = (areas: Area[], label: string, fn: (a: Area) => void, action: LogAction = "style") => {
    if (!areas.length) return;
    atomic(
      () => {
        let n = 0;
        for (const a of areas) {
          fn(a);
          n++;
        }
        return n;
      },
      (s) => s,
      { actor: "user", action, sheet: sheetName(sheet), range: areas.map(areaLabel).join(","), cells: [], note: label }
    );
  };
  const setStyle = (areas: Area[], path: string, value: string, label: string) => applyStyle(areas, label, (a) => mm().updateRangeStyle(a, path, value));
  const styleAt = (p: Pos) => mm().getCellStyle(p.sheet, p.row, p.col).style;
  const toggleFont = (areas: Area[], anchor: Pos, k: "b" | "i" | "u" | "strike", label: string) => {
    const on = !!styleAt(anchor).font[k];
    setStyle(areas, `font.${k}`, on ? "false" : "true", `${label}${on ? "解除" : ""}`);
  };
  const setBorder = (areas: Area[], type: "All" | "Outer" | "None", label: string) =>
    applyStyle(areas, label, (a) => mm().setAreaWithBorder(a, { item: { style: "thin", color: "#000000" }, type } as never));
  const resolveColor = (v: string | [number, number] | undefined) => (v === undefined || !modelRef.current ? undefined : modelRef.current.resolveColor(v));
  // 書式のコピー：コピー元の主な書式を貼り先の範囲へ（罫線は対象外）
  const pasteFormat = (src: Pos, areas: Area[]) => {
    const st = styleAt(src);
    const paths: [string, string][] = [
      ["font.b", String(!!st.font.b)],
      ["font.i", String(!!st.font.i)],
      ["font.u", String(!!st.font.u)],
      ["font.strike", String(!!st.font.strike)],
      ["num_fmt", st.num_fmt || "general"],
      ["alignment.horizontal", st.alignment?.horizontal ?? "general"],
      ["alignment.vertical", st.alignment?.vertical ?? "bottom"],
      ["alignment.wrap_text", String(!!st.alignment?.wrap_text)],
    ];
    const fc = resolveColor(st.font.color);
    if (fc) paths.push(["font.color", fc]);
    paths.push(["fill.bg_color", resolveColor(st.fill.color) ?? "#ffffff"]);
    applyStyle(areas, `元: ${addr(src)}`, (a) => paths.forEach(([p, v]) => mm().updateRangeStyle(a, p, v)), "paste_fmt");
  };

  // ---- 結合・解除・ロック（結合は保存時にサーバが xlsx へ反映。ロックはこのセッション内のみ） ----
  const mergeCells = (g: Merge) => {
    if (uiRef.current.merges.some((o) => o.sheet === g.sheet && !(o.r2 < g.r1 || o.r1 > g.r2 || o.c2 < g.c1 || o.c1 > g.c2))) return alert("すでに結合されたセルを含んでいます");
    const m = mm();
    const cleared: LogCell[] = [];
    for (let r = g.r1; r <= g.r2; r++)
      for (let c = g.c1; c <= g.c2; c++) {
        if (r === g.r1 && c === g.c1) continue;
        const v = m.getCellContent(g.sheet, r, c);
        if (v !== "") {
          m.setUserInput(g.sheet, r, c, ""); // Excel 同様、左上以外の値は捨てる
          cleared.push({ addr: addr({ sheet: g.sheet, row: r, col: c }), before: v, after: "" });
        }
      }
    commit(cleared.length, (s) => ({ ...s, merges: [...s.merges, g] }), {
      actor: "user", action: "merge", sheet: sheetName(g.sheet), range: `${colName(g.c1)}${g.r1}:${colName(g.c2)}${g.r2}`, cells: cleared,
    });
  };
  const unmergeAt = (p: Pos) => {
    const g = mergeAt(p.sheet, p.row, p.col);
    if (!g) return;
    commit(0, (s) => ({ ...s, merges: s.merges.filter((o) => o !== g) }), {
      actor: "user", action: "unmerge", sheet: sheetName(p.sheet), range: `${colName(g.c1)}${g.r1}:${colName(g.c2)}${g.r2}`, cells: [],
    });
  };
  const setLock = (cells: [number, number][], lock: boolean) => {
    if (!cells.length) return;
    commit(
      0,
      (s) => {
        const locks = new Set(s.locks);
        for (const [r, c] of cells) lock ? locks.add(key(sheet, r, c)) : locks.delete(key(sheet, r, c));
        return { ...s, locks };
      },
      { actor: "user", action: lock ? "lock" : "unlock", sheet: sheetName(sheet), range: rangeOf(cells), cells: [] }
    );
  };

  // ---- シートの追加・削除・改名・並べ替え ----
  const sheetOps = {
    add: () =>
      guarded(() => {
        const m = mm();
        const before = m.getWorksheetsProperties().length;
        m.newSheet();
        const name = m.getWorksheetsProperties()[before].name;
        commit(1, (s) => ({ ...s, origNames: [...s.origNames, null] }), { actor: "user", action: "sheet_add", sheet: name, range: "-", cells: [], note: name });
        setSheet(before);
      }),
    rename: (idx: number, name: string) => {
      const cur = sheetName(idx);
      name = name.trim();
      if (!name || name === cur) return;
      if (sheets.some((s, i) => i !== idx && s.name === name)) return alert(`「${name}」は既にあります`);
      guarded(() => {
        mm().renameSheet(idx, name);
        commit(1, (s) => s, { actor: "user", action: "sheet_rename", sheet: name, range: "-", cells: [], note: `${cur} → ${name}` });
      });
    },
    canRemove: () => sheets.filter((s) => s.state === "visible").length > 1,
    remove: (idx: number) =>
      guarded(() => {
        const name = sheetName(idx);
        mm().deleteSheet(idx);
        commit(1, (s) => remapSheets(s, sheetDrop(idx)), { actor: "user", action: "sheet_delete", sheet: name, range: "-", cells: [], note: name });
        setSheet(Math.max(0, Math.min(idx, sheets.length - 2)));
        setSelected(null);
      }),
    move: (idx: number, to: number) => {
      if (to < 0 || to >= sheets.length || to === idx) return;
      guarded(() => {
        mm().moveSheet(idx, to);
        commit(1, (s) => remapSheets(s, sheetMove(idx, to)), { actor: "user", action: "sheet_move", sheet: sheetName(idx), range: "-", cells: [], note: `${sheetName(idx)} を ${to + 1} 番目へ` });
        setSheet(to);
      });
    },
  };

  // ---- コメント（AI への修正指示）と再生成 ----
  const setComment = (p: Pos, text: string) =>
    commit(
      0,
      (s) => ({ ...s, comments: [...s.comments.filter((c) => !same(c, p)), ...(text ? [{ ...p, text }] : [])] }),
      { actor: "user", action: text ? "comment" : "comment_remove", sheet: sheetName(p.sheet), range: addr(p), cells: [], note: text }
    );
  const regenerate = async () => {
    const m = modelRef.current;
    const comments = uiRef.current.comments;
    if (!m || !comments.length) return;
    setRegen(true);
    try {
      const { ops } = await api.ai(comments.map((c) => ({ ...c, current: m.getCellContent(c.sheet, c.row, c.col), comment: c.text })));
      const proposals: Proposal[] = [];
      const cells: LogCell[] = [];
      for (const op of ops) {
        const before = m.getCellContent(op.sheet, op.row, op.col);
        m.setUserInput(op.sheet, op.row, op.col, op.input);
        proposals.push({ sheet: op.sheet, row: op.row, col: op.col, before, after: op.input });
        cells.push({ addr: addr(op), before, after: op.input, note: comments.find((c) => same(c, op))?.text ?? "" });
      }
      const id = entrySeq.current + 1;
      commit(
        ops.length,
        (s) => {
          const marks = new Map(s.marks);
          for (const p of proposals) marks.set(key(p.sheet, p.row, p.col), "ai");
          return { ...s, marks, comments: [], proposals, proposalEntry: id };
        },
        { actor: "ai", action: "ai_update", sheet: "-", range: `${ops.length} セル`, cells, note: "コメントから再生成" }
      );
    } catch (e) {
      alert(String(e));
    } finally {
      setRegen(false);
    }
  };
  const approve = () =>
    commit(
      0,
      (s) => {
        const marks = new Map(s.marks);
        for (const [k, v] of marks) if (v === "ai") marks.set(k, "changed");
        return { ...s, marks, proposals: [], proposalEntry: null };
      },
      {
        actor: "user",
        action: "ai_approve",
        sheet: "-",
        range: `${ui.proposals.length} セル`,
        cells: ui.proposals.map((p) => ({ addr: addr(p), before: p.before, after: p.after })),
        ref: ui.proposalEntry ?? undefined,
      }
    );
  const reject = () => {
    const id = ui.proposalEntry;
    if (id == null) return;
    while (undoStack.current.length) {
      const top = undoStack.current[undoStack.current.length - 1];
      undo();
      if (top.id === id) break;
    }
  };

  // ---- 検索・置換（現在シート） ----
  const findMatches = (q: string, maxC: number) => {
    const m = modelRef.current;
    const qs = q.toLowerCase();
    if (!m || !qs) return [] as { row: number; col: number; content: string }[];
    const out: { row: number; col: number; content: string }[] = [];
    for (let c = 1; c <= Math.max(maxC, 1); c++)
      for (const r of Array.from(m.getRowsWithData(sheet, c))) {
        const content = m.getCellContent(sheet, r, c);
        if (content.toLowerCase().includes(qs) || m.getFormattedCellValue(sheet, r, c).toLowerCase().includes(qs)) out.push({ row: r, col: c, content });
      }
    return out.sort((a, b) => a.row - b.row || a.col - b.col);
  };
  const replaceAll = (q: string, rep: string, maxC: number) => {
    const hits = findMatches(q, maxC).filter((h) => h.content.toLowerCase().includes(q.toLowerCase()));
    if (!hits.length) return 0;
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    applyInputs(hits.map((h) => ({ row: h.row, col: h.col, value: h.content.replace(re, rep) })), "replace", `${q} → ${rep}`);
    return hits.length;
  };

  // ---- 書式付きコピー／貼り付け（IronCalc のクリップボード。glide のテキストコピーと併用） ----
  const copyInternal = (r1: number, c1: number, r2: number, c2: number, cut: boolean) => {
    const m = modelRef.current;
    if (!m) return;
    try {
      m.setSelectedSheet(sheet);
      m.setSelectedCell(r1, c1); // IronCalc は「選択セルが範囲の角」であることを要求する
      m.setSelectedRange(r1, c1, r2, c2);
      const cb = m.copyToClipboard();
      clip.current = { clip: cb, sheet, rows: cb.csv.replace(/\r/g, "").split("\n").map((l) => l.split("\t")), cut };
    } catch (err) {
      console.error(err);
      clip.current = null;
    }
  };
  /** 貼り付けようとしている値が直前の内部コピーと一致すれば書式付きで貼り、true を返す */
  const pasteInternal = (row: number, col: number, values: readonly (readonly string[])[]) => {
    const c = clip.current;
    const m = modelRef.current;
    if (!c || !m) return false;
    const sameText = c.rows.length === values.length && c.rows.every((r, i) => r.length === values[i].length && r.every((v, j) => v === values[i][j]));
    if (!sameText) return false;
    guarded(() => {
      m.setSelectedSheet(sheet);
      m.setSelectedCell(row, col);
      m.pasteFromClipboard(c.sheet, c.clip.range, c.clip.data, c.cut);
      const h = c.rows.length,
        w = Math.max(...c.rows.map((r) => r.length));
      const cells: [number, number][] = [];
      for (let r = row; r < row + h; r++) for (let cc = col; cc < col + w; cc++) cells.push([r, cc]);
      commit(1, markCells(cells), {
        actor: "user", action: "paste", sheet: sheetName(sheet), range: rangeOf(cells), cells: [], note: c.cut ? "切り取り→貼り付け（書式付き）" : "貼り付け（書式付き）",
      });
      if (c.cut) clip.current = null;
    });
    return true;
  };
  // オートフィル：IronCalc の autoFill（連番・曜日・数式のコピーは Excel 同様）
  const fill = (area: Area, vertical: boolean, to: number, dstCells: [number, number][]) =>
    guarded(() => {
      if (vertical) mm().autoFillRows(area, to);
      else mm().autoFillColumns(area, to);
      commit(1, markCells(dstCells), { actor: "user", action: "fill", sheet: sheetName(sheet), range: rangeOf(dstCells), cells: [], note: `元: ${areaLabel(area)}` });
    });

  // ---- 保存：差分だけサーバへ送り、サーバが変更履歴シートと結合を反映した xlsx を返す ----
  const save = async () => {
    const m = modelRef.current;
    if (!m || !sidRef.current) return;
    setBusy(true);
    try {
      const outName = fileName.replace(/\.xlsx$/i, "") + "_updated.xlsx";
      // 保存の記録を先に送り、届いてから同期する（サーバは受信済みのログを履歴シートに書く）
      const e: LogEntry = { seq: ++entrySeq.current, ts: new Date().toISOString(), actor: "user", action: "save", sheet: "-", range: "-", cells: [], note: outName };
      setLogs((l) => [...l, e]);
      await send([e]);
      const u = uiRef.current;
      const blob = await api.sync(sidRef.current, m.flushSendQueue(), u.merges, u.origNames, u.structOps);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = outName;
      a.click();
      URL.revokeObjectURL(url);
      applyUi({ ...uiRef.current, marks: new Map() });
    } catch (e) {
      alert(String(e));
    } finally {
      setBusy(false);
    }
  };

  const focusCell = (p: Pos) => {
    if (p.sheet !== sheet) setSheet(p.sheet);
    setSelected({ row: p.row, col: p.col });
    setFocusReq({ ...p, n: Date.now() });
  };

  return {
    // 状態
    ready, model: modelRef.current, sid, fileName, sheet, sheets, tick, selected, ui, logs, busy, regen, focusReq,
    canUndo: undoStack.current.length > 0, canRedo: redoStack.current.length > 0,
    clip, explicitRows,
    // 読み取りの補助
    sheetName, addr, mergeAt, rangeOf, resolveColor, describe,
    // 操作
    open, setSheet, setSelected, focusCell, undo, redo, refresh,
    applyInputs, clearCells,
    insertRowsAt, deleteRowsAt, insertColsAt, deleteColsAt, hideRows, hideCols, setRowHeight, setColWidth, moveRowsTo, moveColsTo,
    setStyle, toggleFont, setBorder, pasteFormat, styleAt,
    mergeCells, unmergeAt, setLock,
    sheetOps, setComment, regenerate, approve, reject,
    findMatches, replaceAll, copyInternal, pasteInternal, fill, save,
  };
}

export type Workbook = ReturnType<typeof useWorkbook>;
