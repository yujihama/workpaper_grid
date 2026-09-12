// 調書グリッド：このディレクトリごと他アプリへコピーして使う
//   const wb = useWorkbook(createHttpApi("/api"));
//   <WorkpaperGrid wb={wb} />
// 左パネル（コメント一覧・履歴・AI承認）は wb の状態と操作から任意に組める（例: ../App.tsx）
export { WorkpaperGrid } from "./WorkpaperGrid";
export { useWorkbook, type Workbook } from "./useWorkbook";
export { createHttpApi, type WorkbookApi, type AiItem, type AiOp } from "./api";
export { describe, type LogEntry, type LogAction } from "./log";
export { key, same, colName, type Pos, type Comment, type Proposal, type Merge, type Snap, type SheetInfo } from "./types";
import "@glideapps/glide-data-grid/dist/index.css";
import "./styles.css";
