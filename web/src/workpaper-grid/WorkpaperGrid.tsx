import { useEffect, useMemo, useRef, useState } from "react";
import {
  DataEditor,
  GridCellKind,
  CompactSelection,
  type DataEditorRef,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type GridSelection,
  type Item,
  type Theme,
} from "@glideapps/glide-data-grid";
import type { Area } from "@ironcalc/wasm";
import { drawBorders, drawCellText, drawCommentMark, drawLine, drawLockMark, fontOf, hAlignOf, isNumeric, textWidth } from "./render";
import { colName, colRange, key, NUM_FMTS, rowRange, same, type Merge, type Pos } from "./types";
import type { Workbook } from "./useWorkbook";

const MIN_ROWS = 40;
const MIN_COLS = 10;
const ROWNO_W = 44; // 自前の行番号列（Excel 同様に実番号を出すため glide の行マーカーは使わない）
const clampW = (w: number) => Math.min(800, Math.max(30, Math.round(w))); // IronCalcの列幅はpx
const emptySel: GridSelection = { columns: CompactSelection.empty(), rows: CompactSelection.empty() };

type Menu = {
  x: number;
  y: number;
  kind: "row" | "col" | "cell" | "sheet";
  row: number;
  col: number;
  rows: [number, number];
  cols: [number, number];
  sheetIdx?: number;
};
type MenuItem = { label: string; run: () => void } | "-";
type Note = Pos & { x: number; y: number; text: string; initial: string };

/**
 * 調書グリッド本体。ツールバー・数式バー・検索・シートタブ・右クリックメニュー・付箋を含む。
 * グリッド添字（0始まり、列0は行番号列）↔ 実座標（1始まり）の変換はこの中だけで行い、
 * モデル操作はすべて useWorkbook の戻り値（wb）へ実座標で渡す。
 */
export function WorkpaperGrid({ wb, title }: { wb: Workbook; title?: string }) {
  const { sheet, sheets, ui, selected, model } = wb;
  const gridRef = useRef<DataEditorRef | null>(null);
  const [gridSel, setGridSel] = useState<GridSelection>(emptySel);
  const gridSelRef = useRef(gridSel);
  gridSelRef.current = gridSel;
  const [menu, setMenu] = useState<Menu | null>(null);
  const [note, setNote] = useState<Note | null>(null);
  const [colW, setColW] = useState<Record<number, number>>({}); // 列幅ドラッグ中の一時的な幅（確定時にモデルへ）
  const [rowH, setRowH] = useState<Record<number, number>>({}); // 行高ドラッグ中の一時的な高さ
  const [cursor, setCursor] = useState(""); // 行番号の境目にいるときの row-resize カーソル
  const rowEdge = useRef<{ row0: number } | null>(null); // 行番号セルの境界にポインタがある
  const rowDrag = useRef<{ real: number; startY: number; startH: number; h: number } | null>(null);
  const colDrag = useRef<number | null>(null); // 列見出しドラッグ選択の起点
  const moveDrag = useRef<{ kind: "row" | "col"; target: number } | null>(null); // Alt+ドラッグでの行列移動先（グリッド添字）
  const [moveTarget, setMoveTarget] = useState<{ kind: "row" | "col"; idx: number } | null>(null);
  const altDown = useRef(false);
  const [fb, setFb] = useState<{ editing: boolean; text: string; target: Pos | null }>({ editing: false, text: "", target: null }); // 数式バー
  const fbRef = useRef(fb);
  fbRef.current = fb;
  const fbInput = useRef<HTMLInputElement | null>(null);
  const [find, setFind] = useState<{ open: boolean; q: string; rep: string; hits: number | null }>({ open: false, q: "", rep: "", hits: null });
  const fillAll = useRef(false); // Ctrl+Enter で確定したとき、選択範囲すべてに入れる
  const lastRange = useRef<{ r1: number; c1: number; r2: number; c2: number } | null>(null); // 直前に選ばれていた複数セル範囲（実座標）
  const [fmtCopy, setFmtCopy] = useState<Pos | null>(null); // 書式のコピー元（次に選んだ範囲へ貼る）
  const fmtCopyRef = useRef(fmtCopy);
  fmtCopyRef.current = fmtCopy;
  const [renaming, setRenaming] = useState<{ idx: number; name: string } | null>(null); // シート名をタブ上で編集中
  const [confirmBox, setConfirmBox] = useState<{ text: string; ok: () => void } | null>(null); // アプリ内の確認ダイアログ（window.confirm は環境により出ない）

  // ---- 表示する行・列（非表示の行列は抜く。高さ/幅 0 = 非表示）とグリッド添字↔実座標 ----
  const vis = useMemo(() => {
    if (!model || sheet >= sheets.length) return { rows: [] as number[], cols: [] as number[], frozenCols: 0, maxR: 0, maxC: 0 };
    let maxR = 0,
      maxC = 0;
    for (let c = 1; c <= 30; c++) {
      const rs = model.getRowsWithData(sheet, c);
      if (rs.length) {
        maxR = Math.max(maxR, ...Array.from(rs));
        maxC = Math.max(maxC, c);
      }
    }
    const rows: number[] = [],
      cols: number[] = [];
    for (let r = 1; r <= Math.max(maxR + 8, MIN_ROWS); r++) if (model.getRowHeight(sheet, r) > 0) rows.push(r);
    for (let c = 1; c <= Math.max(maxC + 2, MIN_COLS); c++) if (model.getColumnWidth(sheet, c) > 0) cols.push(c);
    const fc = model.getFrozenColumnsCount(sheet); // 枠固定（列）。行の固定は glide に無い
    return { rows, cols, frozenCols: cols.filter((c) => c <= fc).length, maxR, maxC };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [model, sheet, wb.tick, sheets.length]);
  const R = (row0: number) => vis.rows[row0];
  const C = (col0: number) => vis.cols[col0 - 1];
  const rowIdx = (r: number) => vis.rows.indexOf(r);
  const colIdx = (c: number) => vis.cols.indexOf(c) + 1; // 非表示なら 0

  const columns = useMemo<GridColumn[]>(
    () => [
      { title: "", id: "__row", width: ROWNO_W, themeOverride: { bgCell: "#f3f4f6" } },
      ...vis.cols.map((c) => ({ title: colName(c), id: String(c), width: colW[c] ?? clampW(model ? model.getColumnWidth(sheet, c) : 100) })),
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [vis.cols, sheet, wb.tick, colW]
  );
  const commentKeys = useMemo(() => new Set(ui.comments.map((c) => key(c.sheet, c.row, c.col))), [ui.comments]);
  const extOf = (r: number, c: number) => model!.getCellStyle(sheet, r, c);
  const styleOf = (r: number, c: number) => extOf(r, c).style;
  const mergeAt = (r: number, c: number) => wb.mergeAt(sheet, r, c);

  // ---- Excel の「はみ出し」：折り返し無し・左寄せの文字列が幅に収まらないとき、右隣の空セルへ流し込む ----
  const overflowEnd = (r: number, c: number): number | null => {
    const m = model!;
    const content = m.getCellContent(sheet, r, c);
    if (!content || mergeAt(r, c)) return null;
    const display = m.getFormattedCellValue(sheet, r, c);
    const st = styleOf(r, c);
    if (isNumeric(display) || st.alignment?.wrap_text || hAlignOf(st, false) !== "left") return null;
    const need = textWidth(display, fontOf(st).css) + 12;
    let width = colW[c] ?? clampW(m.getColumnWidth(sheet, c));
    if (need <= width) return null;
    let end = c;
    let i = vis.cols.indexOf(c);
    while (width < need && i >= 0 && i < vis.cols.length - 1 && end - c < 8) {
      const nc = vis.cols[i + 1];
      if (m.getCellContent(sheet, r, nc) !== "" || mergeAt(r, nc)) break;
      end = nc;
      width += colW[nc] ?? clampW(m.getColumnWidth(sheet, nc));
      i++;
    }
    return end > c ? end : null;
  };
  const overflowOwner = (r: number, c: number): number | null => {
    const m = model!;
    if (m.getCellContent(sheet, r, c) !== "" || mergeAt(r, c)) return null;
    const i = vis.cols.indexOf(c);
    for (let k = i - 1; k >= 0 && i - k <= 8; k--) {
      const oc = vis.cols[k];
      if (m.getCellContent(sheet, r, oc) === "") continue;
      const end = overflowEnd(r, oc);
      return end !== null && end >= c ? oc : null;
    }
    return null;
  };

  // ---- セルの内容と見た目 ----
  const getCellContent = ([col, row]: Item): GridCell => {
    if (!model || row >= vis.rows.length) return { kind: GridCellKind.Text, data: "", displayData: "", allowOverlay: false };
    if (col === 0)
      return {
        kind: GridCellKind.Text,
        data: String(R(row)),
        displayData: String(R(row)),
        allowOverlay: false,
        readonly: true,
        contentAlign: "center",
        themeOverride: { bgCell: "#f3f4f6", textDark: "#6b7686", baseFontStyle: "400 12px" },
      };
    const r = R(row),
      c = C(col);
    const owner = overflowOwner(r, c);
    if (owner !== null) return getCellContent([colIdx(owner), row]); // はみ出しの受け側は持ち主と同じセルとして描く（span）
    const mg = mergeAt(r, c);
    const ar = mg ? mg.r1 : r, // 結合セルは左上のセルの内容と書式で描く
      ac = mg ? mg.c1 : c;
    const content = model.getCellContent(sheet, ar, ac);
    const display = model.getFormattedCellValue(sheet, ar, ac);
    const st = styleOf(ar, ac);
    const mark = ui.marks.get(key(sheet, ar, ac));
    // 見た目は xlsx の書式そのまま（数式セルも色を変えない）。変更済み／AI更新のマークだけ上書きする
    const themeOverride: Partial<Theme> = {
      bgCell: mark ? (mark === "ai" ? "#ffd6a5" : "#ffe58f") : (wb.resolveColor(st.fill.color) ?? "#ffffff"),
      textDark: wb.resolveColor(st.font.color) ?? "#1f2a37",
      baseFontStyle: `${st.font.i ? "italic " : ""}${st.font.b ? 600 : 400} ${fontOf(st).px}px`,
    };
    const spanOk = mg && mg.c2 > mg.c1 && colIdx(mg.c1) > 0 && colIdx(mg.c2) > 0;
    const ov = mg ? null : overflowEnd(r, c);
    return {
      kind: GridCellKind.Text,
      data: mg && r !== ar ? "" : content, // 結合の2行目以降は内容を出さない（塗りと罫線だけ揃える）
      displayData: mg && r !== ar ? "" : display,
      allowOverlay: true, // 数式セルも Excel 同様に編集可（編集すると数式文字列がそのまま出る）
      readonly: ui.locks.has(key(sheet, ar, ac)),
      allowWrapping: true,
      contentAlign: hAlignOf(st, isNumeric(display)),
      span: spanOk ? [colIdx(mg.c1), colIdx(mg.c2)] : ov !== null ? [col, colIdx(ov)] : undefined,
      themeOverride,
    };
  };

  // ---- 選択（グリッド添字）→ 実座標 ----
  const rangeCells = (sel: GridSelection) => {
    const cells: [number, number][] = [];
    if (sel.current)
      for (const rg of [sel.current.range, ...sel.current.rangeStack])
        for (let y = rg.y; y < rg.y + rg.height; y++) for (let x = Math.max(rg.x, 1); x < rg.x + rg.width; x++) if (y < vis.rows.length) cells.push([R(y), C(x)]);
    return cells;
  };
  const cellsOf = (sel: GridSelection) => {
    const cells = rangeCells(sel);
    for (const y of sel.rows.toArray()) for (const c of vis.cols) if (y < vis.rows.length) cells.push([R(y), c]);
    for (const x of sel.columns.toArray()) if (x > 0) for (const r of vis.rows) cells.push([r, C(x)]);
    return cells;
  };
  const areasOf = (sel: GridSelection): Area[] => {
    const out: Area[] = [];
    if (sel.current)
      for (const rg of [sel.current.range, ...sel.current.rangeStack]) {
        const x0 = Math.max(rg.x, 1);
        if (x0 >= rg.x + rg.width || rg.y + rg.height > vis.rows.length) continue;
        out.push({ sheet, row: R(rg.y), column: C(x0), height: R(rg.y + rg.height - 1) - R(rg.y) + 1, width: C(rg.x + rg.width - 1) - C(x0) + 1 });
      }
    for (const y of sel.rows.toArray()) if (y < vis.rows.length) out.push({ sheet, row: R(y), column: 1, height: 1, width: Math.max(vis.maxC, 1) });
    for (const x of sel.columns.toArray()) if (x > 0) out.push({ sheet, row: 1, column: C(x), height: Math.max(vis.maxR, 1), width: 1 });
    return out;
  };
  const visibleIn = (list: number[], a: number, b: number) => list.filter((x) => x >= a && x <= b).length;
  // 実番号の範囲 [a,b] の中と、その両隣に接する非表示の行・列。あれば再表示すべき範囲と件数を返す
  const hiddenAround = (list: number[], a: number, b: number) => {
    const ia = list.indexOf(a),
      ib = list.indexOf(b);
    const from = Math.min(a, ia > 0 ? list[ia - 1] + 1 : 1);
    const to = Math.max(b, ib >= 0 && ib < list.length - 1 ? list[ib + 1] - 1 : b);
    const count = to - from + 1 - visibleIn(list, from, to);
    return count > 0 ? { from, to, count } : null;
  };

  // ---- 入力 ----
  const onCellsEdited = (edits: readonly { location: Item; value: EditableGridCell }[]) => {
    const list = edits.flatMap(({ location: [col, row], value }) => (value.kind !== GridCellKind.Text || col === 0 ? [] : [{ row: R(row), col: C(col), value: value.data }]));
    if (fillAll.current && list.length === 1) {
      // Ctrl+Enter：選択範囲すべてに同じ値を入れる（Excel 同様）。glide はエディタを開くと範囲を1セルに戻すので、開く前の範囲を使う
      fillAll.current = false;
      const rg = lastRange.current;
      const cells =
        rg && list[0].row >= rg.r1 && list[0].row <= rg.r2 && list[0].col >= rg.c1 && list[0].col <= rg.c2
          ? vis.rows.filter((r) => r >= rg.r1 && r <= rg.r2).flatMap((r) => vis.cols.filter((c) => c >= rg.c1 && c <= rg.c2).map((c) => [r, c] as [number, number]))
          : rangeCells(gridSelRef.current);
      if (cells.length > 1) {
        wb.applyInputs(cells.map(([row, col]) => ({ row, col, value: list[0].value })), "edit", "Ctrl+Enter");
        return true;
      }
    }
    fillAll.current = false;
    wb.applyInputs(list, edits.length > 1 ? "paste" : "edit");
    return true;
  };
  const onPaste = (target: Item, values: readonly (readonly string[])[]) => {
    if (target[0] === 0 || target[1] >= vis.rows.length) return true;
    return !wb.pasteInternal(R(target[1]), C(target[0]), values); // 内部コピーと一致すれば書式付き。それ以外は glide のテキスト貼り付け
  };
  const onFillPattern = (e: { patternSource: { x: number; y: number; width: number; height: number }; fillDestination: { x: number; y: number; width: number; height: number }; preventDefault: () => void }) => {
    e.preventDefault();
    const s = e.patternSource,
      d = e.fillDestination;
    if (s.x === 0) return;
    const area: Area = { sheet, row: R(s.y), column: C(s.x), width: C(s.x + s.width - 1) - C(s.x) + 1, height: R(s.y + s.height - 1) - R(s.y) + 1 };
    const cells: [number, number][] = [];
    for (let y = d.y; y < d.y + d.height; y++) for (let x = Math.max(d.x, 1); x < d.x + d.width; x++) if (y < vis.rows.length) cells.push([R(y), C(x)]);
    if (!cells.length) return;
    const vertical = d.height !== s.height;
    const to = vertical ? (d.y + d.height > s.y + s.height ? R(d.y + d.height - 1) : R(d.y)) : d.x + d.width > s.x + s.width ? C(d.x + d.width - 1) : C(Math.max(d.x, 1));
    wb.fill(area, vertical, to, cells);
  };

  // ---- 右クリックメニュー ----
  const openMenu = (
    kind: Menu["kind"],
    col0: number,
    row0: number,
    ev: { bounds: { x: number; y: number }; localEventX: number; localEventY: number; preventDefault: () => void }
  ) => {
    ev.preventDefault();
    const rs = gridSel.rows.toArray(),
      cs = gridSel.columns.toArray();
    const rg = gridSel.current?.range;
    const row = R(row0),
      col = C(Math.max(col0, 1));
    let rows: [number, number] = [row, row];
    let cols: [number, number] = [col, col];
    if (rs.includes(row0)) rows = [R(Math.min(...rs)), R(Math.max(...rs))];
    else if (rg && kind === "cell" && row0 >= rg.y && row0 < rg.y + rg.height && col0 >= rg.x && col0 < rg.x + rg.width) {
      rows = [R(rg.y), R(rg.y + rg.height - 1)];
      cols = [C(Math.max(rg.x, 1)), C(rg.x + rg.width - 1)];
    }
    if (cs.includes(col0)) cols = [C(Math.min(...cs.filter((x) => x > 0))), C(Math.max(...cs))];
    if (kind === "row" && !rs.includes(row0)) {
      setGridSel({ ...emptySel, rows: CompactSelection.fromSingleSelection(row0) });
      wb.setSelected(null);
    }
    if (kind === "col" && !cs.includes(col0)) {
      setGridSel({ ...emptySel, columns: CompactSelection.fromSingleSelection(col0) });
      wb.setSelected(null);
    }
    setMenu({ x: ev.bounds.x + ev.localEventX, y: ev.bounds.y + ev.localEventY, kind, row, col, rows, cols });
  };
  const menuItems = (mn: Menu): MenuItem[] => {
    if (mn.kind === "sheet") {
      const i = mn.sheetIdx!;
      return [
        { label: "名前の変更…", run: () => setRenaming({ idx: i, name: wb.sheetName(i) }) },
        { label: "シートを追加", run: wb.sheetOps.add },
        {
          label: "シートを削除",
          run: () => {
            if (!wb.sheetOps.canRemove()) return alert("最後のシートは削除できません");
            setConfirmBox({ text: `シート「${wb.sheetName(i)}」を削除します。よろしいですか？`, ok: () => wb.sheetOps.remove(i) });
          },
        },
        "-",
        { label: "左へ移動", run: () => wb.sheetOps.move(i, i - 1) },
        { label: "右へ移動", run: () => wb.sheetOps.move(i, i + 1) },
      ];
    }
    const nR = mn.rows[1] - mn.rows[0] + 1,
      nC = mn.cols[1] - mn.cols[0] + 1;
    const vR = visibleIn(vis.rows, mn.rows[0], mn.rows[1]),
      vC = visibleIn(vis.cols, mn.cols[0], mn.cols[1]);
    const hr = hiddenAround(vis.rows, mn.rows[0], mn.rows[1]),
      hc = hiddenAround(vis.cols, mn.cols[0], mn.cols[1]);
    const rl = vR > 1 ? `${vR} 行` : "行",
      cl = vC > 1 ? `${vC} 列` : "列";
    const rowOps: MenuItem[] = [
      { label: `上に${rl}を挿入`, run: () => wb.insertRowsAt(mn.rows[0], vR) },
      { label: `下に${rl}を挿入`, run: () => wb.insertRowsAt(mn.rows[1] + 1, vR) },
      "-",
      { label: `${rl}を削除`, run: () => wb.deleteRowsAt(mn.rows[0], nR) },
      "-",
      { label: `${rl}を非表示`, run: () => wb.hideRows(mn.rows[0], mn.rows[1], true) },
      ...(hr ? [{ label: `非表示の ${hr.count} 行を再表示（${rowRange(hr.from, hr.to)}）`, run: () => wb.hideRows(hr.from, hr.to, false) }] : []),
    ];
    const colOps: MenuItem[] = [
      { label: `左に${cl}を挿入`, run: () => wb.insertColsAt(mn.cols[0], vC) },
      { label: `右に${cl}を挿入`, run: () => wb.insertColsAt(mn.cols[1] + 1, vC) },
      "-",
      { label: `${cl}を削除`, run: () => wb.deleteColsAt(mn.cols[0], nC) },
      "-",
      { label: `${cl}を非表示`, run: () => wb.hideCols(mn.cols[0], mn.cols[1], true) },
      ...(hc ? [{ label: `非表示の ${hc.count} 列を再表示（${colRange(hc.from, hc.to)}）`, run: () => wb.hideCols(hc.from, hc.to, false) }] : []),
    ];
    if (mn.kind === "row") return rowOps;
    if (mn.kind === "col") return colOps;
    const p = { sheet, row: mn.row, col: mn.col };
    const has = commentKeys.has(key(sheet, mn.row, mn.col));
    const selCells = cellsOf(gridSel);
    const target = selCells.length ? selCells : ([[mn.row, mn.col]] as [number, number][]);
    const inMerge = mergeAt(mn.row, mn.col);
    const rg = gridSel.current?.range;
    const canMerge = !inMerge && rg && rg.width * rg.height > 1 && rg.y + rg.height <= vis.rows.length;
    const allLocked = target.every(([r, c]) => ui.locks.has(key(sheet, r, c)));
    return [
      { label: has ? "コメントを編集…" : "コメントを付ける…", run: () => openNote(mn.col, mn.row) },
      ...(has ? [{ label: "コメントを削除", run: () => wb.setComment(p, "") }] : []),
      "-",
      { label: "内容をクリア", run: () => wb.clearCells(target) },
      ...(canMerge
        ? [{ label: "セルを結合", run: () => wb.mergeCells({ sheet, r1: R(rg.y), c1: C(Math.max(rg.x, 1)), r2: R(rg.y + rg.height - 1), c2: C(rg.x + rg.width - 1) } as Merge) }]
        : []),
      ...(inMerge ? [{ label: "結合を解除", run: () => wb.unmergeAt(p) }] : []),
      { label: allLocked ? "ロックを解除" : "編集をロック", run: () => wb.setLock(target, !allLocked) },
      "-",
      rowOps[0],
      rowOps[3],
      colOps[0],
      colOps[3],
    ];
  };

  // ---- コメント（Excelのメモ相当の付箋） ----
  const openNote = (col: number, row: number) => {
    const b = gridRef.current?.getBounds(colIdx(col), rowIdx(row));
    if (!b) return;
    const existing = ui.comments.find((c) => same(c, { sheet, row, col }))?.text ?? "";
    const x = b.x + b.width + 4 + 300 > window.innerWidth ? b.x - 304 : b.x + b.width + 4;
    setNote({ sheet, row, col, x, y: Math.min(b.y, window.innerHeight - 200), text: existing, initial: existing });
  };
  const saveNote = () => {
    if (!note) return;
    setNote(null);
    const text = note.text.trim();
    if (text !== note.initial) wb.setComment(note, text);
    gridRef.current?.focus();
  };

  // ---- 一覧などからのジャンプ（wb.focusCell）を受けてスクロール ----
  useEffect(() => {
    const f = wb.focusReq;
    if (!f || f.sheet !== sheet) return;
    const y = rowIdx(f.row),
      x = colIdx(f.col);
    if (y < 0 || x <= 0) return; // 非表示の行列にあるセルへは飛ばない
    setGridSel({ ...emptySel, current: { cell: [x, y], range: { x, y, width: 1, height: 1 }, rangeStack: [] } });
    requestAnimationFrame(() => gridRef.current?.scrollTo(x, y, "both", 0, 0, { vAlign: "center", hAlign: "center" }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wb.focusReq, sheet, vis]);

  // ---- 数式バー：選択セルの内容を見る／編集する。「=」で始まる入力中はセルをクリックすると参照が入る ----
  const cellText = selected && model ? model.getCellContent(sheet, selected.row, selected.col) : "";
  const fbCommit = () => {
    const f = fbRef.current;
    setFb({ editing: false, text: "", target: null });
    if (!f.editing || !f.target || f.target.sheet !== sheet) return;
    wb.applyInputs([{ row: f.target.row, col: f.target.col, value: f.text }], "edit", "数式バー");
    gridRef.current?.focus();
  };
  const fbInsertRef = (p: Pos) => {
    const f = fbRef.current;
    const ref = p.sheet === f.target?.sheet ? `${colName(p.col)}${p.row}` : `${wb.sheetName(p.sheet)}!${colName(p.col)}${p.row}`;
    const el = fbInput.current;
    const at = el?.selectionStart ?? f.text.length;
    const text = f.text.slice(0, at) + ref + f.text.slice(el?.selectionEnd ?? at);
    setFb({ ...f, text });
    requestAnimationFrame(() => {
      el?.focus();
      el?.setSelectionRange(at + ref.length, at + ref.length);
    });
  };

  // ---- 書式ツールバー ----
  const anchor = selected ? { sheet, ...selected } : null;
  const ast = anchor && model ? wb.styleAt(anchor) : null;
  const areas = () => areasOf(gridSel);
  const toggle = (k: "b" | "i" | "u" | "strike", label: string) => anchor && wb.toggleFont(areas(), anchor, k, label);

  // ---- キーボード：Excel同様、どこにフォーカスがあっても効くよう window（capture）で受ける ----
  const keys = useRef<(e: KeyboardEvent) => void>(() => {});
  keys.current = (e) => {
    const t = e.target as HTMLElement | null;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    const inText = !!t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.isContentEditable);
    if (inText) {
      if (mod && e.key === "Enter" && t?.tagName === "TEXTAREA" && !note) fillAll.current = true; // glide のセル編集中の Ctrl+Enter
      return; // セル編集中・付箋入力中は通常動作
    }
    if (!model) return;
    if (mod && k === "z" && !e.shiftKey) {
      e.preventDefault();
      wb.undo();
    } else if ((mod && k === "y") || (mod && k === "z" && e.shiftKey)) {
      e.preventDefault();
      wb.redo();
    } else if (mod && (k === "f" || k === "h")) {
      e.preventDefault();
      setFind((f) => ({ ...f, open: true }));
      setTimeout(() => document.getElementById("wg-find-q")?.focus(), 0);
    } else if (mod && (k === "c" || k === "x")) {
      // glide がテキストを OS クリップボードへ入れるのに合わせ、IronCalc 側でも書式付きでコピーしておく（同一アプリ内の貼り付けで使う）
      const rg = gridSelRef.current.current?.range;
      if (rg && rg.x + rg.width > 1 && rg.y + rg.height <= vis.rows.length) {
        const x0 = Math.max(rg.x, 1);
        wb.copyInternal(R(rg.y), C(x0), R(rg.y + rg.height - 1), C(rg.x + rg.width - 1), k === "x");
      }
    } else if (mod && k === "b") {
      e.preventDefault();
      toggle("b", "太字");
    } else if (mod && k === "i") {
      e.preventDefault();
      toggle("i", "斜体");
    } else if (mod && k === "u") {
      e.preventDefault();
      toggle("u", "下線");
    } else if (e.shiftKey && e.key === "F2" && selected && !note) {
      e.preventDefault();
      openNote(selected.col, selected.row);
    } else if (e.key === "Escape") {
      setFmtCopy(null);
      wb.clip.current = null;
    }
  };
  useEffect(() => {
    // capture で受ける：glide のセル編集エディタが Enter を処理して確定する前に Ctrl+Enter のフラグを立てるため
    const h = (e: KeyboardEvent) => {
      altDown.current = e.altKey;
      keys.current(e);
    };
    const u = (e: KeyboardEvent) => (altDown.current = e.altKey);
    window.addEventListener("keydown", h, true);
    window.addEventListener("keyup", u, true);
    return () => {
      window.removeEventListener("keydown", h, true);
      window.removeEventListener("keyup", u, true);
    };
  }, []);
  // Alt+ドラッグの行列移動と、書式コピーの確定、数式バーへのフォーカス復帰は mouseup で行う
  useEffect(() => {
    const up = () => {
      if (fbRef.current.editing && fbRef.current.text.startsWith("=")) setTimeout(() => fbInput.current?.focus(), 0);
      const d = moveDrag.current;
      moveDrag.current = null;
      setMoveTarget(null);
      if (d) {
        const sel = gridSelRef.current;
        if (d.kind === "row") {
          const rs = sel.rows.toArray();
          if (rs.length) wb.moveRowsTo(R(Math.min(...rs)), R(Math.max(...rs)), d.target >= vis.rows.length ? (vis.maxR || R(vis.rows.length - 1)) + 1 : R(d.target));
        } else {
          const cs = sel.columns.toArray().filter((x) => x > 0);
          if (cs.length) wb.moveColsTo(C(Math.min(...cs)), C(Math.max(...cs)), d.target > vis.cols.length ? C(vis.cols.length) + 1 : C(d.target));
        }
        return;
      }
      const src = fmtCopyRef.current;
      if (src) {
        const sel = gridSelRef.current;
        if (sel.current || sel.rows.length || sel.columns.length) {
          setFmtCopy(null);
          setTimeout(() => wb.pasteFormat(src, areasOf(sel)), 0);
        }
      }
    };
    window.addEventListener("mouseup", up);
    return () => window.removeEventListener("mouseup", up);
  });

  // ---- 行高ドラッグ（glide に無いので自前） ----
  const startRowResize = (e: React.MouseEvent) => {
    if (e.button !== 0 || !gridRef.current || e.altKey) return;
    let row0 = rowEdge.current?.row0 ?? -1;
    if (row0 < 0) {
      // ホバーを経ずに押された場合も、行番号列の各行の下端 ±5px を直接判定する
      for (let i = 0; i < vis.rows.length; i++) {
        const b = gridRef.current.getBounds(0, i);
        if (!b) continue;
        if (b.y > e.clientY + 5) break;
        if (e.clientX >= b.x && e.clientX <= b.x + b.width && Math.abs(e.clientY - (b.y + b.height)) <= 5) {
          row0 = i;
          break;
        }
      }
      if (row0 < 0) return;
    }
    e.preventDefault();
    e.stopPropagation();
    const real = R(row0);
    const startH = gridRef.current.getBounds(0, row0)?.height ?? 34;
    rowDrag.current = { real, startY: e.clientY, startH, h: startH };
    const move = (ev: MouseEvent) => {
      const d = rowDrag.current;
      if (!d) return;
      d.h = Math.max(12, Math.round(d.startH + ev.clientY - d.startY));
      setRowH({ [d.real]: d.h });
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      const d = rowDrag.current;
      rowDrag.current = null;
      setRowH({});
      setCursor("");
      if (!d || d.h === d.startH) return;
      wb.setRowHeight(d.real, d.real, d.h);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  };

  const cur = selected ? wb.addr({ sheet, ...selected }) : "";
  const hiddenSheets = sheets.filter((s) => s.state !== "visible");

  return (
    <div className="wg">
      <div className="wg-head">
        <span className="title">{title ?? wb.fileName ?? "調書"}</span>
        <button className="icon" title="元に戻す (Ctrl+Z)" onClick={wb.undo} disabled={!wb.canUndo}>
          ↶
        </button>
        <button className="icon" title="やり直す (Ctrl+Y)" onClick={wb.redo} disabled={!wb.canRedo}>
          ↷
        </button>
        <button onClick={() => setFind((f) => ({ ...f, open: !f.open }))} disabled={!wb.sid} title="Ctrl+F / Ctrl+H">
          検索・置換
        </button>
        <button className="primary" onClick={wb.save} disabled={!wb.sid || wb.busy}>
          保存（xlsx）
        </button>
      </div>
      {wb.sid && (
        <div className="wg-toolbar">
          <button className={ast?.font.b ? "on" : ""} title="太字 (Ctrl+B)" onClick={() => toggle("b", "太字")} disabled={!selected}><b>B</b></button>
          <button className={ast?.font.i ? "on" : ""} title="斜体 (Ctrl+I)" onClick={() => toggle("i", "斜体")} disabled={!selected}><i>I</i></button>
          <button className={ast?.font.u ? "on" : ""} title="下線 (Ctrl+U)" onClick={() => toggle("u", "下線")} disabled={!selected}><u>U</u></button>
          <button className={ast?.font.strike ? "on" : ""} title="取り消し線" onClick={() => toggle("strike", "取り消し線")} disabled={!selected}><s>S</s></button>
          <span className="sep" />
          <button title="左揃え" onClick={() => wb.setStyle(areas(), "alignment.horizontal", "left", "左揃え")} disabled={!selected}>⇤</button>
          <button title="中央揃え" onClick={() => wb.setStyle(areas(), "alignment.horizontal", "center", "中央揃え")} disabled={!selected}>↔</button>
          <button title="右揃え" onClick={() => wb.setStyle(areas(), "alignment.horizontal", "right", "右揃え")} disabled={!selected}>⇥</button>
          <button className={ast?.alignment?.wrap_text ? "on" : ""} title="折り返し" onClick={() => wb.setStyle(areas(), "alignment.wrap_text", ast?.alignment?.wrap_text ? "false" : "true", "折り返し")} disabled={!selected}>⏎</button>
          <span className="sep" />
          <label className="clr" title="塗りつぶし">
            ▩<input type="color" value={wb.resolveColor(ast?.fill.color) ?? "#ffffff"} onChange={(e) => wb.setStyle(areas(), "fill.bg_color", e.target.value, `塗り ${e.target.value}`)} disabled={!selected} />
          </label>
          <label className="clr" title="文字色">
            A<input type="color" value={wb.resolveColor(ast?.font.color) ?? "#000000"} onChange={(e) => wb.setStyle(areas(), "font.color", e.target.value, `文字色 ${e.target.value}`)} disabled={!selected} />
          </label>
          <span className="sep" />
          <button title="外枠罫線" onClick={() => wb.setBorder(areas(), "Outer", "外枠罫線")} disabled={!selected}>▢</button>
          <button title="格子罫線" onClick={() => wb.setBorder(areas(), "All", "格子罫線")} disabled={!selected}>▦</button>
          <button title="罫線なし" onClick={() => wb.setBorder(areas(), "None", "罫線なし")} disabled={!selected}>▢̸</button>
          <span className="sep" />
          <select title="表示形式" value={NUM_FMTS.some(([, v]) => v === ast?.num_fmt) ? ast!.num_fmt : ""} onChange={(e) => e.target.value && wb.setStyle(areas(), "num_fmt", e.target.value, `表示形式 ${e.target.value}`)} disabled={!selected}>
            <option value="">表示形式…</option>
            {NUM_FMTS.map(([l, v]) => (
              <option key={v} value={v}>{l}</option>
            ))}
          </select>
          <span className="sep" />
          <button className={fmtCopy ? "on" : ""} title="書式のコピー：押してから貼り先の範囲を選ぶ（Esc で中止）" onClick={() => setFmtCopy(fmtCopy ? null : anchor)} disabled={!selected}>
            🖌 書式コピー
          </button>
        </div>
      )}
      {wb.sid && (
        <div className="wg-fbar">
          <span className="addr">{cur || "—"}</span>
          <span className="fx">fx</span>
          <input
            ref={fbInput}
            value={fb.editing ? fb.text : cellText}
            placeholder={selected ? "" : "セルを選択"}
            disabled={!selected && !fb.editing}
            onFocus={() => !fbRef.current.editing && selected && setFb({ editing: true, text: cellText, target: { sheet, ...selected } })}
            onChange={(e) => setFb((f) => ({ ...f, text: e.target.value }))}
            onKeyDown={(e) => {
              if (e.key === "Enter") fbCommit();
              if (e.key === "Escape") {
                setFb({ editing: false, text: "", target: null });
                gridRef.current?.focus();
              }
            }}
            onBlur={() => {
              if (fbRef.current.editing && !fbRef.current.text.startsWith("=")) fbCommit(); // 参照入力中（= 始まり）以外は、外を押したら確定
            }}
          />
          {fb.editing && fb.text.startsWith("=") && <span className="hint">セルをクリックすると参照が入ります。Enter で確定</span>}
        </div>
      )}
      {find.open && (
        <div className="wg-findbar">
          <input id="wg-find-q" placeholder="検索" value={find.q} onChange={(e) => setFind((f) => ({ ...f, q: e.target.value, hits: null }))} onKeyDown={(e) => e.key === "Enter" && findNext()} />
          <input placeholder="置換後" value={find.rep} onChange={(e) => setFind((f) => ({ ...f, rep: e.target.value }))} onKeyDown={(e) => e.key === "Enter" && replaceAll()} />
          <button onClick={findNext} disabled={!find.q}>次を検索</button>
          <button onClick={replaceAll} disabled={!find.q}>すべて置換</button>
          {find.hits !== null && <span className="hint">{find.hits} 件</span>}
          <button className="x" onClick={() => setFind((f) => ({ ...f, open: false }))}>×</button>
        </div>
      )}
      {sheets.length > 0 && (
        <div className="wg-tabs">
          {sheets.map((s, i) =>
            s.state !== "visible" ? null : renaming?.idx === i ? (
              <input
                key={`rename-${i}`}
                className="rename"
                autoFocus
                value={renaming.name}
                onChange={(e) => setRenaming({ idx: i, name: e.target.value })}
                onKeyDown={(e) => {
                  if (e.key === "Enter") applyRename();
                  if (e.key === "Escape") setRenaming(null);
                }}
                onBlur={applyRename}
              />
            ) : (
              <button
                key={`${i}-${s.name}`}
                className={i === sheet ? "on" : ""}
                onClick={() => {
                  wb.setSheet(i);
                  wb.setSelected(null);
                  setGridSel(emptySel);
                }}
                onDoubleClick={() => setRenaming({ idx: i, name: s.name })}
                onContextMenu={(e) => {
                  e.preventDefault();
                  setMenu({ x: e.clientX, y: e.clientY, kind: "sheet", row: 0, col: 0, rows: [0, 0], cols: [0, 0], sheetIdx: i });
                }}
                title="ダブルクリックで名前を変更、右クリックで操作"
              >
                {s.name}
              </button>
            )
          )}
          <button className="add" title="シートを追加" onClick={wb.sheetOps.add}>
            ＋
          </button>
          {hiddenSheets.length > 0 && (
            <span className="hidden-sheets" title={hiddenSheets.map((s) => `${s.name}（${s.state}）`).join("、")}>
              非表示シート {hiddenSheets.length}: {hiddenSheets.map((s) => s.name).join("、")}
            </span>
          )}
        </div>
      )}
      <div
        className={`wg-grid ${cursor ? "rowresize" : fmtCopy ? "fmtcopy" : ""}`}
        onMouseDownCapture={(e) => {
          // 数式バーで「=」入力中は、セルをクリックしても入力欄のフォーカスを保つ（参照挿入のため）
          if (fbRef.current.editing && fbRef.current.text.startsWith("=")) e.preventDefault();
          startRowResize(e);
        }}
      >
        {model && sheet < sheets.length ? (
          <DataEditor
            key={`${sheet}-${sheets.map((s) => s.name).join("|")}`}
            ref={gridRef}
            columns={columns}
            rows={vis.rows.length}
            getCellContent={getCellContent}
            onCellsEdited={onCellsEdited}
            getCellsForSelection
            onPaste={onPaste}
            fillHandle
            onFillPattern={onFillPattern}
            keybindings={{ search: false }}
            onDelete={(sel) => {
              wb.clearCells(cellsOf(sel));
              return false;
            }}
            rowMarkers="none"
            minColumnWidth={30}
            maxColumnWidth={800}
            onColumnResize={(_c, w, idx) => idx >= 1 && setColW((s) => ({ ...s, [C(idx)]: w }))}
            onColumnResizeEnd={(_c, w, idx) => {
              if (idx < 1) return;
              setColW({});
              wb.setColWidth(C(idx), C(idx), w); // 列見出しの境界ドラッグで幅を確定（Excel 同様）
            }}
            rowSelect="multi"
            columnSelect="multi"
            rangeSelect="rect"
            width="100%"
            height="100%"
            smoothScrollX
            smoothScrollY
            gridSelection={gridSel}
            onGridSelectionChange={(sel) => {
              const cur = sel.current;
              // 数式バーで「=」入力中のクリックは参照の挿入にする（選択は動かさない）
              if (fbRef.current.editing && fbRef.current.text.startsWith("=") && cur && cur.cell[0] > 0 && cur.cell[1] < vis.rows.length) {
                fbInsertRef({ sheet, row: R(cur.cell[1]), col: C(cur.cell[0]) });
                return;
              }
              if (cur && cur.range.x === 0) {
                // 行番号列をクリック／ドラッグ／Shift+クリックしたら Excel と同じく行選択にする（current は残して範囲拡張を効かせる）
                setGridSel({ ...sel, rows: CompactSelection.fromSingleSelection([cur.range.y, cur.range.y + cur.range.height]) });
                wb.setSelected(null);
                return;
              }
              setGridSel(sel);
              const c = cur?.cell;
              wb.setSelected(c && c[0] > 0 && c[1] < vis.rows.length ? { col: C(c[0]), row: R(c[1]) } : null);
              if (cur && cur.range.width * cur.range.height > 1 && cur.range.y + cur.range.height <= vis.rows.length) {
                const x0 = Math.max(cur.range.x, 1);
                lastRange.current = { r1: R(cur.range.y), c1: C(x0), r2: R(cur.range.y + cur.range.height - 1), c2: C(cur.range.x + cur.range.width - 1) };
              } else if (cur && (!lastRange.current || R(cur.cell[1]) < lastRange.current.r1 || R(cur.cell[1]) > lastRange.current.r2 || C(cur.cell[0]) < lastRange.current.c1 || C(cur.cell[0]) > lastRange.current.c2)) {
                lastRange.current = null; // 範囲の外を選んだら忘れる（範囲内の1セルに縮んだのはエディタ起動なので保持）
              }
            }}
            onCellContextMenu={([col, row], ev) => openMenu(col === 0 ? "row" : "cell", col, row, ev)}
            onHeaderContextMenu={(col, ev) => col >= 1 && openMenu("col", col, 0, ev)}
            onMouseMove={(a) => {
              // 行番号の境界（上下 4px）にいるときは行高ドラッグの構え（カーソルを row-resize に）
              if (a.buttons === 0) {
                let edge: number | null = null;
                if (a.kind === "cell" && a.location[0] === 0) {
                  if (a.localEventY > a.bounds.height - 4) edge = a.location[1];
                  else if (a.localEventY < 4 && a.location[1] > 0) edge = a.location[1] - 1;
                }
                rowEdge.current = edge === null ? null : { row0: edge };
                setCursor(edge === null ? "" : "row-resize");
              }
              if (a.buttons !== 1) {
                colDrag.current = null;
                return;
              }
              // Alt+ドラッグ：選択中の行（行番号列）／列（見出し）を移動。移動先は境界線で示す（glide のイベントに altKey が無いので window で追う）
              if (altDown.current) {
                if (a.kind === "cell" && a.location[0] === 0 && gridSel.rows.length) {
                  const t = a.location[1] + (a.localEventY > a.bounds.height / 2 ? 1 : 0);
                  moveDrag.current = { kind: "row", target: t };
                  setMoveTarget({ kind: "row", idx: t });
                } else if (a.kind === "header" && a.location[0] >= 1 && gridSel.columns.length) {
                  const t = a.location[0] + (a.localEventX > a.bounds.width / 2 ? 1 : 0);
                  moveDrag.current = { kind: "col", target: t };
                  setMoveTarget({ kind: "col", idx: t });
                }
                return;
              }
              // 列見出しを押したままドラッグで複数列を選択（Excel 同様。glide は Shift/Ctrl+クリックしか持たない）
              if (a.kind !== "header" || a.location[0] < 1) return;
              const col = a.location[0];
              if (colDrag.current == null) colDrag.current = gridSel.columns.length ? (gridSel.columns.first() ?? col) : col;
              const lo = Math.min(colDrag.current, col),
                hi = Math.max(colDrag.current, col);
              setGridSel({ ...emptySel, columns: CompactSelection.fromSingleSelection([Math.max(lo, 1), hi + 1]) });
              wb.setSelected(null);
            }}
            rowHeight={(row) => {
              const real = R(row);
              if (rowH[real] !== undefined) return rowH[real]; // ドラッグ中
              const h = model.getRowHeight(sheet, real);
              if (wb.explicitRows.current.has(key(sheet, real, 0))) return Math.max(12, Math.round(h)); // 自分で指定した行はそのまま
              return h < 30 ? 34 : Math.round(h); // 既定に近い低い行高は折り返しが読める 34px に揃え、明示的に高い行はそれに従う
            }}
            freezeColumns={1 + vis.frozenCols}
            drawCell={(a, draw) => {
              const { ctx, rect } = a;
              if (a.col === 0) {
                draw();
                const prev = a.row > 0 ? R(a.row - 1) : 0;
                if (R(a.row) - prev > 1) drawLine(ctx, rect.x, rect.y + 1, rect.x + rect.width, rect.y + 1, "#1f2a37", 2); // 直前に非表示行がある印
                if (moveTarget?.kind === "row" && moveTarget.idx === a.row) drawLine(ctx, rect.x, rect.y + 1, rect.x + rect.width + 2000, rect.y + 1, "#2563eb", 3);
                return;
              }
              const r = R(a.row),
                c = C(a.col);
              const mg = mergeAt(r, c);
              drawBorders(ctx, rect, styleOf(r, c), { top: !mg || r === mg.r1, bottom: !mg || r === mg.r2, left: !mg || c === mg.c1, right: !mg || c === mg.c2 }, wb.resolveColor);
              if (a.cell.kind === GridCellKind.Text) {
                // 文字は自前で描く（縦位置・折り返し・下線・取り消し線・フォント名・条件付き書式のアイコン/バー）
                const ar = mg ? mg.r1 : r,
                  ac = mg ? mg.c1 : c;
                if (!mg || r === ar) drawCellText(ctx, rect, extOf(ar, ac), a.cell.displayData, a.theme.textDark, wb.resolveColor);
              } else draw();
              if (moveTarget?.kind === "col" && moveTarget.idx === a.col) drawLine(ctx, rect.x + 1, rect.y, rect.x + 1, rect.y + rect.height, "#2563eb", 3);
              if (ui.locks.has(key(sheet, r, c))) drawLockMark(ctx, rect);
              if (commentKeys.has(key(sheet, r, c))) drawCommentMark(ctx, rect);
            }}
          />
        ) : (
          <div className="wg-empty">調書を開いてください</div>
        )}
      </div>
      <div className="wg-legend">
        <span><i style={{ background: "#ffe58f" }} />変更済み</span>
        <span><i style={{ background: "#ffd6a5" }} />AI更新（承認待ち）</span>
        <span><i className="tri" />コメント</span>
        <span className="cur">{cur}</span>
      </div>
      <div className="wg-help">
        右クリック: 行・列の挿入/削除/非表示、結合、ロック、コメント ／ Alt+ドラッグ: 選択した行・列の移動 ／ Delete: クリア ／ Ctrl+Z / Y ／ Ctrl+C / X / V（アプリ内は書式付き）／
        Ctrl+Enter: 選択範囲すべてに入力 ／ Ctrl+F: 検索・置換 ／ Ctrl+B / I / U ／ 右下の■をドラッグ: オートフィル
      </div>

      {menu && (
        <>
          <div
            className="wg-backdrop"
            onMouseDown={() => setMenu(null)}
            onContextMenu={(e) => {
              e.preventDefault();
              setMenu(null);
            }}
          />
          <div className="wg-menu" style={{ left: Math.min(menu.x, window.innerWidth - 280), top: Math.min(menu.y, window.innerHeight - 380) }}>
            {menuItems(menu).map((it, i) =>
              it === "-" ? (
                <hr key={i} />
              ) : (
                <button
                  key={i}
                  onClick={() => {
                    setMenu(null);
                    it.run();
                    gridRef.current?.focus(); // 続けてキー操作できるようグリッドへ戻す
                  }}
                >
                  {it.label}
                </button>
              )
            )}
          </div>
        </>
      )}

      {confirmBox && (
        <>
          <div className="wg-backdrop dim" onMouseDown={() => setConfirmBox(null)} />
          <div className="wg-dialog">
            <p>{confirmBox.text}</p>
            <div className="row">
              <button
                className="primary"
                autoFocus
                onClick={() => {
                  const ok = confirmBox.ok;
                  setConfirmBox(null);
                  ok();
                }}
              >
                削除する
              </button>
              <button onClick={() => setConfirmBox(null)}>キャンセル</button>
            </div>
          </div>
        </>
      )}

      {note && (
        <>
          <div className="wg-backdrop" onMouseDown={saveNote} />
          <div className="wg-note" style={{ left: note.x, top: note.y }}>
            <div className="note-head">{wb.addr(note)} へのコメント（AIへの修正指示）</div>
            <textarea
              autoFocus
              value={note.text}
              placeholder="例: 承認印の有無を確認した旨も記載して"
              onChange={(e) => setNote({ ...note, text: e.target.value })}
              onKeyDown={(e) => {
                if (e.key === "Escape") setNote(null);
                if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) saveNote();
              }}
            />
            <div className="row">
              <button className="primary" onClick={saveNote}>
                {note.initial ? "更新" : "付ける"}
              </button>
              {note.initial && (
                <button
                  onClick={() => {
                    setNote(null);
                    wb.setComment(note, "");
                  }}
                >
                  削除
                </button>
              )}
              <button onClick={() => setNote(null)}>キャンセル</button>
              <span className="hint">Ctrl+Enter で確定</span>
            </div>
          </div>
        </>
      )}
    </div>
  );

  function applyRename() {
    const r = renaming;
    setRenaming(null);
    if (r) wb.sheetOps.rename(r.idx, r.name);
  }
  function findNext() {
    const hits = wb.findMatches(find.q, vis.maxC);
    setFind((f) => ({ ...f, hits: hits.length }));
    if (!hits.length) return;
    const cur = selected ?? { row: 0, col: 0 };
    const next = hits.find((h) => h.row > cur.row || (h.row === cur.row && h.col > cur.col)) ?? hits[0];
    wb.focusCell({ sheet, ...next });
  }
  function replaceAll() {
    setFind((f) => ({ ...f, hits: wb.replaceAll(find.q, find.rep, vis.maxC) }));
  }
}
