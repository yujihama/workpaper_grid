// 編集ログ：1操作 = 1件（undo の単位と同じ）。追記のみで、undo/redo も「取り消した」という記録として残す
export type LogCell = { addr: string; before: string; after: string; note?: string };
export type LogAction =
  | "edit" | "paste" | "clear" | "fill" | "replace"
  | "insert_rows" | "delete_rows" | "insert_cols" | "delete_cols"
  | "hide_rows" | "unhide_rows" | "hide_cols" | "unhide_cols"
  | "row_height" | "col_width" | "move_rows" | "move_cols"
  | "style" | "paste_fmt" | "merge" | "unmerge" | "lock" | "unlock"
  | "sheet_add" | "sheet_delete" | "sheet_rename" | "sheet_move"
  | "comment" | "comment_remove"
  | "ai_update" | "ai_approve"
  | "undo" | "redo" | "save";
export type LogEntry = {
  seq: number;
  ts: string;
  actor: "user" | "ai";
  action: LogAction;
  sheet: string;
  range: string;
  cells: LogCell[];
  note?: string;
  ref?: number;
};
export type LogInput = Omit<LogEntry, "seq" | "ts">;

// ログ1件を人が読める1行にする
const q = (s: string) => (s.length > 24 ? s.slice(0, 24) + "…" : s || "（空）");
export function describe(e: LogEntry): string {
  const one = e.cells.length === 1 ? e.cells[0] : null;
  const chg = one ? `${one.addr} ${q(one.before)} → ${q(one.after)}` : `${e.range}（${e.cells.length}セル）`;
  const at = `${e.sheet}!${e.range}`;
  switch (e.action) {
    case "edit": return `入力 ${chg}`;
    case "paste": return `貼り付け ${chg}`;
    case "clear": return `クリア ${chg}`;
    case "fill": return `オートフィル ${at}`;
    case "replace": return `置換 ${chg}（${e.note}）`;
    case "insert_rows": return `行を挿入 ${at}`;
    case "delete_rows": return `行を削除 ${at}`;
    case "insert_cols": return `列を挿入 ${at}`;
    case "delete_cols": return `列を削除 ${at}`;
    case "hide_rows": return `行を非表示 ${at}`;
    case "unhide_rows": return `行を再表示 ${at}`;
    case "hide_cols": return `列を非表示 ${at}`;
    case "unhide_cols": return `列を再表示 ${at}`;
    case "row_height": return `行の高さ ${at} → ${e.note}`;
    case "col_width": return `列の幅 ${at} → ${e.note}`;
    case "move_rows": return `行を移動 ${at} → ${e.note}`;
    case "move_cols": return `列を移動 ${at} → ${e.note}`;
    case "style": return `書式 ${at}: ${e.note}`;
    case "paste_fmt": return `書式のコピー ${at}`;
    case "merge": return `セルを結合 ${at}`;
    case "unmerge": return `結合を解除 ${at}`;
    case "lock": return `編集をロック ${at}`;
    case "unlock": return `ロックを解除 ${at}`;
    case "sheet_add": return `シートを追加「${e.note}」`;
    case "sheet_delete": return `シートを削除「${e.note}」`;
    case "sheet_rename": return `シート名を変更 ${e.note}`;
    case "sheet_move": return `シートを移動 ${e.note}`;
    case "comment": return `コメント ${e.range}「${q(e.note ?? "")}」`;
    case "comment_remove": return `コメント削除 ${e.range}`;
    case "ai_update": return `AI更新 ${chg}`;
    case "ai_approve": return `AI更新を承認（#${e.ref}）`;
    case "undo": return `元に戻す（#${e.ref}）`;
    case "redo": return `やり直す（#${e.ref}）`;
    case "save": return `保存 ${e.note}`;
  }
}
