// IronCalc の行高・列幅の単位と Excel のピクセルの換算（実測値）
//   行高: IronCalc = pt × 1.5625。高さ未指定の行は IronCalc が一律 25.0 を返す（ファイルの既定行高は見ていない）
//   列幅: IronCalc = 文字数 × 9。幅未指定の列は一律 90。Excel のピクセルは Calibri 11 で 文字数 × 7 + 5
export const IC_ROW_DEFAULT = 25.0;
export const IC_COL_DEFAULT = 90.0;
export const EXCEL_ROW_DEFAULT_PX = 20; // 15pt
export const EXCEL_COL_DEFAULT_PX = 64; // 8.43 文字

export const icToPxH = (ic: number, defaultPx = EXCEL_ROW_DEFAULT_PX) => (ic <= 0 ? 0 : Math.abs(ic - IC_ROW_DEFAULT) < 1e-6 ? defaultPx : ((ic / 1.5625) * 4) / 3);
export const pxToIcH = (px: number) => ((px * 3) / 4) * 1.5625;
export const icToPxW = (ic: number, defaultPx = EXCEL_COL_DEFAULT_PX) => (ic <= 0 ? 0 : Math.abs(ic - IC_COL_DEFAULT) < 1e-6 ? defaultPx : Math.round((ic / 9) * 7 + 5));
export const pxToIcW = (px: number) => (Math.max(px - 5, 1) / 7) * 9;
