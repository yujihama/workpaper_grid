import type { CellStyle, ExtendedCellStyle } from "@ironcalc/wasm";

// Canvas 描画の純関数群。モデルには触らず、呼び出し側が取り出したスタイルを受け取る
export type Rect = { x: number; y: number; width: number; height: number };
export type ColorResolver = (v: string | [number, number] | undefined) => string | undefined;
type BorderItem = { style: string; color?: string | [number, number] } | undefined;

export const fontOf = (st: CellStyle) => {
  const px = Math.round(((st.font.sz || 11) * 96) / 72);
  return { px, css: `${st.font.i ? "italic " : ""}${st.font.b ? "600" : "400"} ${px}px "${st.font.name || "Calibri"}", "Noto Sans JP", "Yu Gothic", sans-serif` };
};
export const isNumeric = (display: string) => /^-?[\d,]+(\.\d+)?%?$/.test(display);
export const hAlignOf = (st: CellStyle, numeric: boolean): "left" | "center" | "right" => {
  const h = st.alignment?.horizontal;
  return h === "center" || h === "centerContinuous" ? "center" : h === "right" || ((h === undefined || h === "general") && numeric) ? "right" : "left";
};

// 文字幅の計測（はみ出し判定に使う）
let measureCtx: CanvasRenderingContext2D | null = null;
export const textWidth = (text: string, font: string) => {
  if (!measureCtx) measureCtx = document.createElement("canvas").getContext("2d");
  const ctx = measureCtx!;
  ctx.font = font;
  return Math.max(...text.split("\n").map((l) => ctx.measureText(l).width));
};

// セルが必要とする高さ（px）。Excel の「行の高さの自動調整」を再現するのに使う。
// Excel の実測: Calibri 11pt は 1 行 20px・2 行 40px、14pt は 25px → 1 行あたり フォント px × 1.36
export const neededHeight = (text: string, st: CellStyle, maxW: number) => {
  if (!measureCtx) measureCtx = document.createElement("canvas").getContext("2d");
  const ctx = measureCtx!;
  const { px, css } = fontOf(st);
  ctx.font = css;
  const lines = st.alignment?.wrap_text ? wrapLines(ctx, text, maxW) : text.split("\n");
  return lines.length * Math.round(px * 1.36);
};

// 折り返し：Excel と同じく文字単位（日本語向け）
export const wrapLines = (ctx: CanvasRenderingContext2D, text: string, maxW: number) => {
  const out: string[] = [];
  for (const para of text.split("\n")) {
    let line = "";
    for (const ch of para) {
      if (line && ctx.measureText(line + ch).width > maxW) {
        out.push(line);
        line = ch;
      } else line += ch;
    }
    out.push(line);
  }
  return out;
};

const ICON_GLYPH: Record<string, string> = {
  ArrowUp: "↑", ArrowRight: "→", ArrowDown: "↓", ArrowAngleUp: "↗", ArrowAngleDown: "↘", Circle: "●", TriangleUp: "▲", TriangleDown: "▼",
  FlatRectangle: "▬", Rhombus: "◆", Flag: "⚑", Check: "✓", Cross: "✕", Exclamation: "!",
};

// 文字の描画：縦位置・折り返し・下線・取り消し線・フォント名・条件付き書式のアイコン/バーは glide が扱えないので自前で描く
export function drawCellText(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  ext: ExtendedCellStyle,
  display: string,
  fg: string,
  color: ColorResolver,
  opts: { underline?: boolean } = {}
) {
  const st = ext.style;
  const { px, css } = fontOf(st);
  const pad = 6;
  let x0 = rect.x + pad,
    innerW = rect.width - pad * 2;
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();
  // 条件付き書式：データバー → アイコン → 星 の順に描き、文字は右へ寄せる
  let showValue = true;
  if (ext.data_bar) {
    const db = ext.data_bar;
    const v = Math.max(-1, Math.min(1, db.value));
    const axis = rect.x + 2 + (rect.width - 4) * Math.max(0, Math.min(1, db.axis_position));
    ctx.fillStyle = color(v >= 0 ? db.positive_color : db.negative_color) ?? "#638ec6";
    const w = (rect.width - 4) * Math.abs(v) * (v >= 0 ? 1 - db.axis_position : db.axis_position);
    ctx.fillRect(v >= 0 ? axis : axis - w, rect.y + 3, w, rect.height - 6);
    showValue = db.show_value;
  }
  if (ext.icon) {
    ctx.font = `${px}px sans-serif`;
    ctx.fillStyle = color(ext.icon.color) ?? fg;
    ctx.textBaseline = "middle";
    ctx.fillText(ICON_GLYPH[ext.icon.icon] ?? "●", x0, rect.y + rect.height / 2);
    x0 += px + 4;
    innerW -= px + 4;
    showValue = ext.icon.show_value;
  }
  if (ext.rating) {
    ctx.font = `${px}px sans-serif`;
    ctx.fillStyle = color(ext.rating.color) ?? "#f5b301";
    ctx.textBaseline = "middle";
    const stars = "★".repeat(ext.rating.count) + "☆".repeat(Math.max(0, ext.rating.max - ext.rating.count));
    ctx.fillText(stars, x0, rect.y + rect.height / 2);
    const w = ctx.measureText(stars).width + 4;
    x0 += w;
    innerW -= w;
    showValue = ext.rating.show_value;
  }
  if (showValue && display) {
    ctx.font = css;
    ctx.fillStyle = fg;
    ctx.textBaseline = "top";
    const numeric = isNumeric(display);
    const lines = st.alignment?.wrap_text ? wrapLines(ctx, display, innerW) : display.split("\n");
    const lineH = Math.round(px * 1.3);
    const total = lines.length * lineH;
    const v = st.alignment?.vertical ?? "bottom"; // Excel の既定は下揃え
    // 余白は 1px。Excel 既定の 20px 行に 11pt（19px）が収まるように。収まらないときは上を切らず下を切る
    let y0 = v === "top" ? rect.y + 1 : v === "center" ? rect.y + (rect.height - total) / 2 : rect.y + rect.height - total - 1;
    y0 = Math.max(y0, rect.y + 1);
    const ha = hAlignOf(st, numeric);
    lines.forEach((line, i) => {
      const w = ctx.measureText(line).width;
      const x = ha === "right" ? x0 + innerW - w : ha === "center" ? x0 + (innerW - w) / 2 : x0;
      const y = y0 + i * lineH;
      ctx.fillText(line, x, y);
      if (st.font.u || st.font.strike || opts.underline) {
        ctx.strokeStyle = fg;
        ctx.lineWidth = 1;
        ctx.beginPath();
        if (st.font.u || opts.underline) {
          ctx.moveTo(x, y + px + 1);
          ctx.lineTo(x + w, y + px + 1);
        }
        if (st.font.strike) {
          ctx.moveTo(x, y + px * 0.55);
          ctx.lineTo(x + w, y + px * 0.55);
        }
        ctx.stroke();
      }
    });
  }
  ctx.restore();
}

// 罫線：セルの4辺と斜線を xlsx の線種で描く。edges は結合セルの外周判定（内側の辺は描かない）
export function drawBorders(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  st: CellStyle,
  edges: { top: boolean; bottom: boolean; left: boolean; right: boolean },
  color: ColorResolver
) {
  const b = st.border as {
    left?: BorderItem; right?: BorderItem; top?: BorderItem; bottom?: BorderItem; diagonal?: BorderItem; diagonal_up?: boolean; diagonal_down?: boolean;
  };
  const sides: [BorderItem, boolean, number, number, number, number][] = [
    [b.top, edges.top, rect.x, rect.y, rect.x + rect.width, rect.y],
    [b.bottom, edges.bottom, rect.x, rect.y + rect.height, rect.x + rect.width, rect.y + rect.height],
    [b.left, edges.left, rect.x, rect.y, rect.x, rect.y + rect.height],
    [b.right, edges.right, rect.x + rect.width, rect.y, rect.x + rect.width, rect.y + rect.height],
    [b.diagonal, !!b.diagonal_up, rect.x, rect.y + rect.height, rect.x + rect.width, rect.y], // 斜線（右上がり）
    [b.diagonal, !!b.diagonal_down, rect.x, rect.y, rect.x + rect.width, rect.y + rect.height], // 斜線（右下がり）
  ];
  for (const [it, on, x1, y1, x2, y2] of sides) {
    if (!it || !on) continue;
    const s = it.style;
    ctx.save();
    ctx.strokeStyle = color(it.color) ?? "#000";
    ctx.lineWidth = s === "thick" ? 3 : s.startsWith("medium") ? 2 : 1;
    ctx.setLineDash(s === "dotted" ? [1, 2] : s.includes("dash") ? [4, 2] : []);
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    if (s === "double") {
      const dx = x1 === x2 ? 2 : 0,
        dy = y1 === y2 ? 2 : 0;
      ctx.beginPath();
      ctx.moveTo(x1 + (x1 === rect.x ? dx : -dx), y1 + (y1 === rect.y ? dy : -dy));
      ctx.lineTo(x2 + (x1 === rect.x ? dx : -dx), y2 + (y1 === rect.y ? dy : -dy));
      ctx.stroke();
    }
    ctx.restore();
  }
}

// 小物：コメントの赤三角・ロック・移動先の線
export function drawCommentMark(ctx: CanvasRenderingContext2D, rect: Rect) {
  ctx.save();
  ctx.fillStyle = "#e5484d";
  ctx.beginPath();
  ctx.moveTo(rect.x + rect.width - 9, rect.y);
  ctx.lineTo(rect.x + rect.width, rect.y);
  ctx.lineTo(rect.x + rect.width, rect.y + 9);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}
// Excel のメモ（旧コメント）は紫の三角で、AI 向けコメント（赤）と区別する
export function drawNoteMark(ctx: CanvasRenderingContext2D, rect: Rect) {
  ctx.save();
  ctx.fillStyle = "#7c3aed";
  ctx.beginPath();
  ctx.moveTo(rect.x + rect.width - 9, rect.y + rect.height);
  ctx.lineTo(rect.x + rect.width, rect.y + rect.height);
  ctx.lineTo(rect.x + rect.width, rect.y + rect.height - 9);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}
// 入力規則（リスト）のドロップダウン印
export function drawDropdownMark(ctx: CanvasRenderingContext2D, rect: Rect) {
  ctx.save();
  ctx.fillStyle = "#6b7686";
  ctx.font = "10px sans-serif";
  ctx.textBaseline = "middle";
  ctx.fillText("▾", rect.x + rect.width - 12, rect.y + rect.height / 2);
  ctx.restore();
}
export function drawLockMark(ctx: CanvasRenderingContext2D, rect: Rect) {
  ctx.save();
  ctx.font = "10px sans-serif";
  ctx.fillStyle = "#6b7686";
  ctx.textBaseline = "top";
  ctx.fillText("🔒", rect.x + 2, rect.y + 2);
  ctx.restore();
}
export function drawLine(ctx: CanvasRenderingContext2D, x1: number, y1: number, x2: number, y2: number, color: string, width: number) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.restore();
}
