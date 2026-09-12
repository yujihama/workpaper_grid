import { key, parseKey, type Mark, type Pos, type Snap } from "./types";

// 行/列の挿入・削除・移動に合わせて、セルに紐づく状態（マーク・コメント・提案・結合・ロック）の座標をずらす
export type Move = (row: number, col: number) => [number, number] | null;

export function remap(s: Snap, sheet: number, f: Move): Snap {
  const marks = new Map<string, Mark>();
  for (const [k, v] of s.marks) {
    const [ks, kr, kc] = parseKey(k);
    if (ks !== sheet) {
      marks.set(k, v);
      continue;
    }
    const p = f(kr, kc);
    if (p) marks.set(key(ks, p[0], p[1]), v);
  }
  const locks = new Set<string>();
  for (const k of s.locks) {
    const [ks, kr, kc] = parseKey(k);
    if (ks !== sheet) {
      locks.add(k);
      continue;
    }
    const p = f(kr, kc);
    if (p) locks.add(key(ks, p[0], p[1]));
  }
  const move = <T extends Pos>(xs: T[]) =>
    xs.flatMap((x) => {
      if (x.sheet !== sheet) return [x];
      const p = f(x.row, x.col);
      return p ? [{ ...x, row: p[0], col: p[1] }] : [];
    });
  const merges = s.merges.flatMap((g) => {
    if (g.sheet !== sheet) return [g];
    const a = f(g.r1, g.c1),
      b = f(g.r2, g.c2);
    return a && b ? [{ ...g, r1: a[0], c1: a[1], r2: b[0], c2: b[1] }] : []; // 結合の一部が消えたら結合は解く
  });
  return { ...s, marks, locks, comments: move(s.comments), proposals: move(s.proposals), merges };
}

// シートの削除・移動に合わせてシート番号を付け替える
export function remapSheets(s: Snap, f: (si: number) => number | null): Snap {
  const mk = <T extends { sheet: number }>(xs: T[]) => xs.flatMap((x) => (f(x.sheet) === null ? [] : [{ ...x, sheet: f(x.sheet)! }]));
  const marks = new Map<string, Mark>();
  for (const [k, v] of s.marks) {
    const [ks, kr, kc] = parseKey(k);
    const ns = f(ks);
    if (ns !== null) marks.set(key(ns, kr, kc), v);
  }
  const locks = new Set<string>();
  for (const k of s.locks) {
    const [ks, kr, kc] = parseKey(k);
    const ns = f(ks);
    if (ns !== null) locks.add(key(ns, kr, kc));
  }
  return { ...s, marks, locks, comments: mk(s.comments), proposals: mk(s.proposals), merges: mk(s.merges) };
}

export const identity: Move = (r, c) => [r, c];
export const rowShift = (at: number, n: number): Move => (r, c) => (r >= at ? [r + n, c] : [r, c]);
export const rowDrop = (at: number, n: number): Move => (r, c) => (r < at ? [r, c] : r < at + n ? null : [r - n, c]);
export const colShift = (at: number, n: number): Move => (r, c) => (c >= at ? [r, c + n] : [r, c]);
export const colDrop = (at: number, n: number): Move => (r, c) => (c < at ? [r, c] : c < at + n ? null : [r, c - n]);
// 行 row から count 行を delta だけ動かす（間の行は逆方向へ count ずれる）。IronCalc の moveRows と同じ意味
export const rowMove = (row: number, count: number, delta: number): Move => (r, c) => {
  if (r >= row && r < row + count) return [r + delta, c];
  if (delta > 0 && r >= row + count && r < row + count + delta) return [r - count, c];
  if (delta < 0 && r >= row + delta && r < row) return [r + count, c];
  return [r, c];
};
export const colMove = (col: number, count: number, delta: number): Move => (r, c) => {
  if (c >= col && c < col + count) return [r, c + delta];
  if (delta > 0 && c >= col + count && c < col + count + delta) return [r, c - count];
  if (delta < 0 && c >= col + delta && c < col) return [r, c + count];
  return [r, c];
};
// シートを idx から to へ動かしたときの番号の付け替え
export const sheetMove = (idx: number, to: number) => (si: number) =>
  si === idx ? to : idx < to ? (si > idx && si <= to ? si - 1 : si) : si >= to && si < idx ? si + 1 : si;
export const sheetDrop = (idx: number) => (si: number) => (si < idx ? si : si === idx ? null : si - 1);
