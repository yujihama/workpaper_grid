# 調書グリッド 試作（IronCalc + Glide Data Grid）

生成済みのxlsx調書を右ペインに開き、Excelと同じ感覚で手直しする／コメントを付けてAIに該当箇所だけ再生成させる、を試すための最小構成。

## 構成
| 層 | 部品 | ライセンス |
|---|---|---|
| グリッド表示・編集 | @glideapps/glide-data-grid | MIT |
| 計算エンジン（ブラウザ・即時） | @ironcalc/wasm | MIT / Apache 2.0 |
| 計算エンジン（サーバ・確定＆xlsx保存） | ironcalc (Python) | MIT / Apache 2.0 |
| API | FastAPI + uvicorn | MIT / BSD |

流れ: xlsx → サーバでIronCalc形式へ変換 → ブラウザで表示・編集（再計算はWASM内で即時）
→ 保存時に `flushSendQueue()` の差分だけ送信 → サーバの同じエンジンに適用 → `save_to_xlsx` で書き出し。

## 構成（移植単位）
```
web/src/workpaper-grid/     ← このフォルダごとコピーすれば他アプリで使える
  index.ts                  公開 API（WorkpaperGrid, useWorkbook, createHttpApi, describe, 型）
  types.ts                  型・定数（座標は 1 始まりの実番号）
  remap.ts                  行列/シート操作に伴う座標の付け替え（純関数）
  log.ts                    編集ログの型と人が読む表記（純関数）
  render.ts                 Canvas 描画：文字・罫線・条件付き書式（純関数）
  units.ts                  IronCalc の行高・列幅の単位 ↔ Excel ピクセルの換算
  api.ts                    サーバ契約 WorkbookApi と HTTP 実装
  useWorkbook.ts            モデル・履歴・ログ・スナップショットと全操作（React hook）
  WorkpaperGrid.tsx         グリッド UI（ツールバー・数式バー・検索・タブ・メニュー・付箋）
  styles.css                接頭辞 wg- のスタイル
web/src/App.tsx             試作の殻（左パネル：開く・コメント・履歴・AI 承認）
server/workpaper.py         FastAPI の APIRouter（open / log / sync / ai）
server/app.py               router を /api に include するだけ
```

### 他アプリへの載せ方
フロント（React 18+ / Vite 想定）
```bash
npm i @glideapps/glide-data-grid @ironcalc/wasm     # peer 依存の都合で --legacy-peer-deps が必要な場合あり
```
```tsx
import { useWorkbook, WorkpaperGrid, createHttpApi } from "./workpaper-grid";
const wb = useWorkbook(createHttpApi("/api"));   // サーバの prefix に合わせる
<WorkpaperGrid wb={wb} title={wb.fileName} />    // 高さは親要素に従う（flex 子として height:100%）
```
- `vite.config.ts` に `optimizeDeps: { exclude: ["@ironcalc/wasm"] }` を入れる（wasm を事前バンドルさせない）
- 左パネル相当（コメント一覧・履歴・AI 承認）は `wb.ui.comments / wb.logs / wb.ui.proposals` と `wb.focusCell / wb.setComment / wb.regenerate / wb.approve / wb.reject` で自由に組める。`App.tsx` が例
- サーバを差し替えるときは `WorkbookApi`（open / sync / appendLog / ai の 4 つ）を実装して `useWorkbook` に渡す
- 操作はすべて `wb.*` を実座標（1 始まり）で呼ぶ。グリッド添字との変換は `WorkpaperGrid` の中だけ

サーバ（FastAPI）
```python
pip install fastapi uvicorn python-multipart ironcalc openpyxl
from workpaper import router
app.include_router(router, prefix="/api")
```
- `SESSIONS` / `LOGS` はメモリ。本番は永続化する。`/ai` はダミーなので既存の AI ロジックに置き換える（入出力の形は `AiItem` → `ops[{sheet,row,col,input}]`）

## 動かし方
```bash
# サーバ（Python 3.12）
python -m venv venv && source venv/bin/activate     # Windows: uv venv --python 3.12 venv
pip install -r server/requirements.txt              #          venv\Scripts\python -m pip install -r server/requirements.txt
uvicorn server.app:app --port 8000

# フロント（別ターミナル）
cd web && npm install --legacy-peer-deps && npm run dev   # http://localhost:5173
```
左の「サンプル調書を開く」で `sample_workpaper.xlsx`（文章欄＋計算シート、結合セル・数式あり）が開きます。
「初回生成は済んでいる」想定なので、開いた時点をレビュー開始の状態とみなします。

## 操作（想定する動線）
| やりたいこと | 操作 |
|---|---|
| 手で直す | セルをクリックして入力（Enterで確定）。数式セルも編集可（数式バーで内容が分かる）。開くと数式文字列がそのまま出る |
| まとめて消す | 範囲選択 → Delete |
| コピー／貼り付け | Ctrl+C / Ctrl+V（Excelからの貼り付けも可） |
| 元に戻す／やり直す | Ctrl+Z / Ctrl+Y（右上の ↶ ↷ でも可）。行列操作・AI更新・コメントも戻せる |
| 行の挿入・削除 | 行番号を右クリック → 上に挿入／下に挿入／削除。複数行を選んでいればその行数分 |
| 列の挿入・削除 | 列見出しを右クリック → 左に挿入／右に挿入／削除 |
| セルから行列操作 | セルを右クリック → メニュー下段（範囲選択していればその範囲分） |
| AIに直させる | セルを右クリック →「コメントを付ける…」（または Shift+F2）→ 指示を書いて「付ける」。セル右上に赤い三角、左パネルに一覧 |
| 再生成 | コメントを必要なだけ付けたら、左の「コメントを反映して再生成」。該当セルだけ更新され橙になる |
| AI更新の確認 | 左下の「AIの更新」に 変更前 → 変更後 が並ぶ。クリックでセルへジャンプ。「すべて承認」で確定（黄=変更済み扱い）、「取り消す」で更新前に戻る（コメントも復元） |
| 保存 | 右上「保存（xlsx）」→ 差分をサーバへ送り、`_updated.xlsx` をダウンロード |

色: 元の xlsx の書式のまま（数式セルも色を変えない） ／ 黄=変更済み ／ 橙=AI更新（承認待ち） ／ 赤三角=コメントあり

## 書式の表示
xlsx の書式はサーバ側 IronCalc が保持し、保存後の xlsx にそのまま残る。画面にも次を反映している。

| 書式 | 取得元 | 表示 |
|---|---|---|
| 塗り・フォント色・太字・斜体・サイズ | wasm `getCellStyle` + `resolveColor`（テーマ色も解決） | セルの背景・文字に反映。変更済み(黄)/AI更新(橙)のマークが優先 |
| 罫線 | 同上 `border` | 4辺を線種（thin/medium/thick/double/dotted/dashed）と色で描画 |
| 横位置 | 同上 `alignment.horizontal` | left/center/right。general は数値なら右寄せ |
| 表示形式 | `getFormattedCellValue` | `0.00%` などをそのまま表示。編集時は元の値 |
| 列幅・行高 | `getColumnWidth` / `getRowHeight` を `units.ts` で Excel のピクセルに換算（実測: 行高 = pt×1.5625、列幅 = 文字数×9。未指定の行列は IronCalc が一律 25 / 90 を返すので、`sheetFormatPr` の既定値をサーバから受けて使う） | Excel と同じピクセルで表示。高さを明示していない行（シートの最頻値と同じ高さ）は、Excel と同じくフォントサイズ・折り返し・セル内改行に合わせて自動調整（Calibri 11 で 1 行 20px） |
| 結合セル | サーバで openpyxl から読み `X-Merges` ヘッダで渡す（wasm に API がない） | 左上セルの内容・書式で描き、横結合は 1 セルに見せる。罫線は外周のみ。行列の挿入・削除に追従 |

文字の描画は glide に任せず自前で行っている（`drawCell` で `draw()` を呼ばず、`getCellStyle` の内容で描く）。それにより次も表示できる:

| 書式 | 表示 |
|---|---|
| 縦位置（上/中央/下） | 反映。指定が無ければ Excel 既定の下揃え |
| 折り返し | `wrap_text` のセルだけ文字単位で折り返す。それ以外は 1 行 |
| はみ出し | 折り返し無し・左寄せの文字列が幅に収まらないとき、Excel 同様に右隣の空セルへ流し込む（右隣に値があれば切れる）。glide の `span` で実現 |
| 下線・取り消し線・フォント名 | 反映（フォントは無ければ Noto Sans JP 等にフォールバック） |
| 斜線罫線 | 右上がり／右下がりとも描画 |
| 条件付き書式 | 塗り・文字色は `getCellStyle` が評価済みで返すのでそのまま反映。アイコンセットは記号（▲▼●✓✕ 等）、データバーは棒、評価は★で近似 |
| 枠固定 | 列の固定は反映（行番号列＋固定列）。**行の固定は glide に機能が無く未対応** |

未対応: インデント・縮小して全体表示（IronCalc の alignment に項目が無い）、右寄せ文字列の左へのはみ出し、行の枠固定。

### 非表示シート・非表示行列
- **非表示シート**（hidden / veryHidden）: タブに出さず、タブ列の右端に「非表示シート n: 名前」と薄く表示。参照する数式は普通に計算される。保存後も状態は残る
- **非表示行・列**: `getRowHeight` / `getColumnWidth` が 0 のものをグリッドから抜く。行番号は Excel と同じ実番号（3 の次が 6）で、直前に非表示行がある位置に太線。列見出しも実際の列名
- **操作**: 行番号・列見出しの右クリックに「行/列を非表示」「非表示の n 行/列を再表示（範囲）」。再表示は、非表示を挟んで選択したときに加え、非表示に隣接する 1 行/1 列を右クリックしただけでも出る（Excel より見つけやすく）。ログ・Ctrl+Z 対象
- **複数選択**: 行番号・列見出しとも、ドラッグ／Shift+クリック／Ctrl+クリックで複数選択できる（列見出しのドラッグは glide に無いので `onMouseMove` で実装）
- 行番号列は自前（glide の行マーカーは連番しか出せないため）。クリック／ドラッグで行選択、右クリックで行メニュー
- 検証用: 「非表示ありサンプル」ボタンで `sample_hidden.xlsx`（非表示シート2枚・非表示行 4〜5・非表示列 C）を開ける

### 列幅・行高の変更
- 列幅: 列見出しの境界をドラッグ（Excel 同様）
- 行高: 行番号の境界（上下 4px）をドラッグ（Excel 同様。glide に無いので自前実装。カーソルが ↕ になる位置でつかむ）。自分で変えた行はその高さで固定（自動調整しない）
- どちらもログ・Ctrl+Z 対象で、保存後の xlsx に反映される

## Excel との差分（棚卸し）
IronCalc の読込→保存で何が残るかを要素ごとに検証した結果（`openpyxl` で作った 1 要素ずつのブックを通した）:

| 要素 | IronCalc 読込 | 保存後 | 画面 |
|---|---|---|---|
| 値・数式・表示形式・塗り・罫線・フォント（太字/斜体/下線/取り消し線/色/サイズ） | ○ | 残る | 表示（下線・取り消し線は未描画） |
| 結合セル・列幅・行高・非表示行列・非表示シート | ○ | 残る | 表示 |
| 枠固定（freeze panes） | ○ | 残る | 未反映（glide は先頭列固定のみ可能） |
| 条件付き書式 | ○ | 残る | 未反映（アイコン・データバー含む） |
| 名前定義 | ○ | 残る | — |
| **Excel のメモ（旧コメント）** | **× 開けない**（`Zip Error: specified file not found in archive`） | — | — |
| **グラフ・図形・画像** | ○ | **消える** | 非表示 |
| オートフィルタ | ○ | **消える** | 非表示 |
| 入力規則（ドロップダウン等） | ○ | **消える** | 非表示 |
| ハイパーリンク | ○ | **消える**（文字だけ残る） | リンク不可 |
| マクロ・ピボット | 未検証 | 未検証 | — |

調書に監査証拠の画像やメモが貼られている運用なら、IronCalc 単独では成り立たない。現実的な対処は
「IronCalc は値・数式・書式の編集エンジンとして使い、保存時は **元の xlsx を openpyxl で開いて IronCalc から差分（値・数式・書式・行列操作）だけを書き戻す**」方式で、図形・メモ・フィルタ・入力規則を元ファイル側に温存すること。ただし行列の挿入・削除で図形やメモの位置合わせが必要になる。

### 追加した Excel 相当の操作
| 操作 | 実装 |
|---|---|
| 数式バー | グリッド上部。選択セルの数式を編集せずに確認でき、直接編集も可。`=` で始まる入力中にセルをクリックすると参照（`C7` / `シート!C7`）が入る |
| 検索・置換 | Ctrl+F / Ctrl+H。現在シートの値・数式を大文字小文字無視で検索、「すべて置換」は数式の中も置換（Excel 同様）。1 操作でログ 1 件 |
| コピー・切り取り・貼り付け | glide の OS クリップボード（テキスト）に加え、Ctrl+C / X の時点で IronCalc の書式付きクリップボードも取る。貼り付け内容が一致すれば書式付きで貼り、切り取りなら元を消す。Excel など外部からの貼り付けはテキストのまま |
| Ctrl+Enter | セル編集中に Ctrl+Enter で確定すると、選択範囲すべてに同じ値が入る |
| オートフィル | 選択範囲右下の■をドラッグ。IronCalc の `autoFillRows/Columns` で連番・数式コピーは Excel と同じ規則 |
| 書式 | ツールバー: 太字/斜体/下線/取り消し線（Ctrl+B/I/U）、左/中/右、折り返し、塗り、文字色、罫線（外枠/格子/なし）、表示形式。`updateRangeStyle` / `setAreaWithBorder` |
| 書式のコピー | 「🖌 書式コピー」を押してから貼り先を選ぶ（Esc で中止）。font/fill/num_fmt/alignment を転写（罫線は対象外） |
| セルの結合・解除 | 右クリック。左上以外の値は Excel 同様に捨てる。結合状態はブラウザが持ち、保存時に `merges` として送ってサーバが openpyxl で xlsx に反映（wasm に結合 API が無いため） |
| シート | タブの右クリック: 名前の変更／追加／削除／左右へ移動。「＋」で追加。マーク・コメント・結合のシート番号も付け替える |
| 行・列の移動 | 行番号／列見出しを **Alt+ドラッグ**（青い線が移動先）。IronCalc の `moveRows/moveColumns`。通常ドラッグは選択なので Alt で区別 |
| 編集ロック | 右クリック →「編集をロック」。🔒 が付き入力・クリア・貼り付けを弾く。**セッション内のみ**（IronCalc に保護属性が無く xlsx には残らない） |

画面側でまだ Excel と違う主なもの:
- インデント・縮小表示、右寄せ文字列の左へのはみ出し、行の枠固定（glide の制約）
- セルのドラッグ移動（範囲の縁を掴む操作）、参照色分け（数式中の参照セルに色枠）
- シート保護の永続化、印刷設定、ズーム

## 元ファイルの温存（メモ・図形・画像・フィルタ・入力規則・リンク）
IronCalc は自分のモデルに無いものを書き出せず、メモ付きブックは読込すら失敗する。そこで IronCalc を「計算・編集エンジン」に限定し、
保存は **元の xlsx を土台に、IronCalc の出力からセル部分だけを移植する**（`server/xlsx_pkg.py`）。

| 段階 | 内容 |
|---|---|
| 開く | `strip_for_engine`: メモ（comments + VML）を zip レベルで取り除いたコピーを IronCalc に渡す。元ファイルは `ORIGINALS[sid]` に保管 |
| 表示 | `GET /api/assets/{sid}`: 元ファイルからメモ・画像（data URL）・グラフ/図形（位置と名前）・リンク・入力規則を読み、グリッドに重ねる。メモは紫の三角＋選択時に本文、リンクは青下線＋Ctrl+クリック、入力規則（リスト）は ▾ で候補、画像は実物、グラフ/図形は枠と名前 |
| 保存 | `graft`: 元 xlsx の各シート XML の `dimension / sheetFormatPr / cols / sheetData / mergeCells / conditionalFormatting` を IronCalc 出力のものに差し替え、`styles.xml` `sharedStrings.xml` `metadata.xml` も IronCalc のものに置換。`calcChain` は削除（Excel が作り直す）。それ以外（`drawing / legacyDrawing / autoFilter / dataValidations / hyperlinks / pageSetup / vbaProject …`）は元のまま |
| 構造操作 | 行列の挿入・削除・移動はブラウザが `structOps` として持ち（undo/redo 対応）、保存時にサーバが温存要素の位置に再生する: `autoFilter@ref`、`dataValidation@sqref`、`hyperlink@ref`、メモの `ref` と VML の `Row/Column`、図形アンカーの `row/col` |
| シート操作 | 各シートが元ファイルのどのシートかを `origNames` で追跡（改名しても対応が切れない）。新規シート・変更履歴は IronCalc 出力の XML を持ち込み、削除したシートは外す |

検証（`web/public/sample_rich.xlsx`: メモ 2・グラフ・画像・フィルタ・入力規則・リンク・結合）: 開ける／表示できる／編集＋行挿入＋結合後の保存で全パーツが残り、
フィルタ `A2:C8→A2:C10`、入力規則 `D3:D8→D3:D10`、リンク `A12→A14`、メモ `C5→C7`、図形アンカーが行挿入に追従。保存後のファイルは openpyxl と IronCalc で再読込可。
**Excel 本体での「修復」ダイアログの有無は未確認**なので、実際の調書で確かめること。

未対応: メモの新規作成・本文編集の書き戻し（既存メモは温存のみ）、グラフの描画（枠のみ）、画像・図形の移動、入力規則の list 以外（数値範囲等）の検証。

## 編集ログ（監査証跡）
調書は「誰が・いつ・どこを・何から何に・なぜ」変えたかが残る必要があるため、操作をそのままログにする。

- **単位**: 1操作 = 1件（Ctrl+Z の単位と同じ）。貼り付けや AI 更新のように複数セルが変わる操作は 1件の中に `cells[]` として展開
- **追記のみ**: 元に戻す／やり直すも「取り消した」という記録として追加する（消さない）。保存も記録する
- **記録項目**: `seq, ts, actor(user/ai), action, sheet, range, cells[{addr, before, after, note}], note, ref`
  `action` = edit / paste / clear / insert_rows / delete_rows / insert_cols / delete_cols / comment / comment_remove / ai_update / ai_approve / undo / redo / save
  AI 更新は根拠になったコメントを `cells[].note` に持つ。`ref` は承認・undo・redo が指す元の操作番号
- **保管**: ブラウザで保持しつつ、操作のたびに `POST /api/log/{sid}` へ送る。サーバは受信時刻と直前行のハッシュを含めた `hash` を付けて `<tmp>/workpaper_logs/<sid>.jsonl` に追記（本番は DB＋利用者IDを付与）
- **同梱**: 保存時に xlsx へ「変更履歴」シートを書き出す（1セル1行、`=` 始まりは文字列化）。調書ファイル単体でも履歴を追える
- **参照**: 左パネル「履歴」タブに新しい順で表示。行をクリックすると該当セルへジャンプ。`GET /api/log/{sid}` で JSON 取得

## 確認できたこと
- シート間の数式（手続!B11 = サンプリング!B6）が入力直後に更新される（再計算は0.1ms未満）
- 行挿入で下の数式参照がずれずに保存される。コメント・変更マークも行列操作に追従する
- 結合セル・罫線・塗り・列幅・数式は保存後のxlsxに残る
- コメント → `/api/ai`（複数まとめて送る）→ ops（sheet,row,col,input）→ 適用 → 橙表示 → 承認/取り消し
- 1回の操作（貼り付けの複数セル、行挿入、AI更新）が Ctrl+Z 1回で戻る

## 対応済み（当初の要検討項目）
- 結合セルの表示・操作: サーバが openpyxl で結合範囲を読んで `X-Merges` で渡し、ブラウザが保持して描画。結合・解除もでき、保存時に xlsx へ反映（「書式の表示」「追加した Excel 相当の操作」参照）
- 行高・列幅: ドラッグで変更でき、保存にも反映。セル内改行・折り返し・縦位置は自前描画で表示（「書式の表示」参照）
- 編集ロック: 右クリックで「編集をロック」。ただしセッション内のみ（下記）

## 未対応・要検討
- **編集ロックの永続化**: IronCalc に保護属性が無いため xlsx には残らない。テンプレート側で入力可セルを定義して、開くときにロックを復元する設計が必要
- **コメントはブラウザ内のみ**: AI への指示として使い切る前提で、xlsx のメモとしては書き出していない。再生成後は消える
- **Excel のメモ付きブックが開けない**（IronCalc の読込エラー）。**グラフ・図形・画像・オートフィルタ・入力規則・ハイパーリンクは保存で消える**（「Excel との差分（棚卸し）」で実測済み）。マクロ・ピボットは未検証。対処は「元 xlsx へ差分を書き戻す保存方式」への変更
- 表示: インデント・縮小表示、右寄せ文字列の左へのはみ出し、行の枠固定（glide の制約）
- サーバのセッション・ログはメモリ／JSONL 保持（本番は永続化・排他制御・利用者 ID の付与が必要）
- `/api/ai` はダミー。実運用では既存 AI ロジックの出力を ops 形式に整形する（`server/workpaper.py` の `ai_edit`）
