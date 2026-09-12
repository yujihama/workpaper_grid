import { columnNameFromNumber } from "@ironcalc/wasm";

// 座標はすべて 1 始まりの実番号（Excel と同じ）。グリッド添字への変換は WorkpaperGrid の中だけで行う
export type Pos = { sheet: number; row: number; col: number };
export type Mark = "changed" | "ai"; // セルの状態（黄=手動変更、橙=AI更新）
export type Comment = Pos & { text: string }; // 再生成の指示（Excelのメモ相当）
export type Proposal = Pos & { before: string; after: string }; // AIが更新した内容（承認待ち）
export type Merge = { sheet: number; r1: number; c1: number; r2: number; c2: number }; // 結合セル（両端含む）
export type SheetInfo = { name: string; state: string }; // state: visible / hidden / veryHidden
export type CellEdit = { row: number; col: number; value: string };

// 元ファイルにあって IronCalc が持たないもの（表示用。保存時は元ファイル側に残す）
export type Anchor = { col: number; row: number; colOff: number; rowOff: number }; // 1始まり、オフセットは EMU
export type Emu = { cx: number; cy: number };
export type ImageAsset = { sheet: number; from: Anchor; to: Anchor | null; ext: Emu | null; dataUrl: string };
export type ShapeAsset = { sheet: number; from: Anchor; to: Anchor | null; ext: Emu | null; name: string; text?: string; kind: "chart" | "shape" };
export type Note = Pos & { author: string; text: string }; // Excel のメモ（旧コメント）
export type Hyperlink = Pos & { target: string };
export type Validation = { sheet: number; r1: number; c1: number; r2: number; c2: number; type: string; formula1: string };
export type SheetDefaults = { sheet: number; rowPx: number; colPx: number }; // 既定の行高・列幅（Excel のピクセル）
export type Assets = { notes: Note[]; images: ImageAsset[]; shapes: ShapeAsset[]; hyperlinks: Hyperlink[]; validations: Validation[]; defaults: SheetDefaults[] };
export const EMPTY_ASSETS: Assets = { notes: [], images: [], shapes: [], hyperlinks: [], validations: [], defaults: [] };

// 行列の挿入・削除・移動。保存時にサーバへ送り、温存した図形・メモ等の位置に再生する
export type StructOp = { sheet_orig_name: string | null; kind: "insert_rows" | "delete_rows" | "insert_cols" | "delete_cols" | "move_rows" | "move_cols"; at: number; n: number; delta: number };

// セルに紐づく UI 状態。履歴（Ctrl+Z）はこのスナップショットとモデルの undo を組にして戻す
export type Snap = {
  marks: Map<string, Mark>;
  comments: Comment[];
  proposals: Proposal[];
  proposalEntry: number | null;
  merges: Merge[];
  locks: Set<string>; // 編集ロック（セッション内のみ。IronCalc に保護属性が無いため xlsx には残らない）
  origNames: (string | null)[]; // 各シートが元ファイルのどのシートか（新規シートは null）。保存時の土台合わせに使う
  structOps: StructOp[];
  assets: Assets;
};

export const key = (s: number, r: number, c: number) => `${s}:${r}:${c}`;
export const parseKey = (k: string) => k.split(":").map(Number) as [number, number, number];
export const same = (a: Pos, b: Pos) => a.sheet === b.sheet && a.row === b.row && a.col === b.col;
export const EMPTY_SNAP: Snap = {
  marks: new Map(), comments: [], proposals: [], proposalEntry: null, merges: [], locks: new Set(), origNames: [], structOps: [], assets: EMPTY_ASSETS,
};
export const EMU_PER_PX = 9525;

export const colName = columnNameFromNumber;
export const colNum = (s: string) => {
  let col = 0;
  for (const ch of s) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col;
};
export const rowRange = (a: number, b: number) => (b > a ? `${a}:${b}` : `${a}`);
export const colRange = (a: number, b: number) => (b > a ? `${colName(a)}:${colName(b)}` : colName(a));

// ツールバーの表示形式（IronCalc の num_fmt）
export const NUM_FMTS: [string, string][] = [
  ["標準", "general"],
  ["数値 #,##0", "#,##0"],
  ["小数 #,##0.00", "#,##0.00"],
  ["パーセント 0.0%", "0.0%"],
  ["日付 yyyy/mm/dd", "yyyy/mm/dd"],
  ["文字列 @", "@"],
];
