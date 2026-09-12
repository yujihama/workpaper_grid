import { useMemo, useState } from "react";
import { createHttpApi, describe, key, same, useWorkbook, WorkpaperGrid } from "./workpaper-grid";

/**
 * 試作の殻。右は WorkpaperGrid、左はこのアプリ固有のパネル（ファイルを開く・コメント一覧・履歴・AI 承認）。
 * 他アプリへ載せるときは workpaper-grid/ をコピーし、この左パネル相当を自分の画面で組む。
 */
export default function App() {
  const api = useMemo(() => createHttpApi("/api"), []);
  const wb = useWorkbook(api);
  const [leftTab, setLeftTab] = useState<"comments" | "history">("comments");
  const { ui, sheet, selected } = wb;

  // 「初回生成済み」の想定：生成された調書をワンクリックで開く
  const openSample = async (name = "sample_workpaper.xlsx") => {
    const r = await fetch(`/${name}`);
    if (!r.ok) return alert(`サンプル調書（web/public/${name}）が見つかりません`);
    await wb.open(new File([await r.blob()], name));
  };
  const hhmmss = (iso: string) => new Date(iso).toLocaleTimeString("ja-JP", { hour12: false });
  // ログの「手続!C8」表記からセルへ
  const focusAddr = (a: string) => {
    const m = /^(.+)!([A-Z]+)(\d+)$/.exec(a);
    if (!m) return;
    const si = wb.sheets.findIndex((s) => s.name === m[1]);
    if (si < 0) return;
    let col = 0;
    for (const ch of m[2]) col = col * 26 + (ch.charCodeAt(0) - 64);
    wb.focusCell({ sheet: si, row: Number(m[3]), col });
  };

  return (
    <div className="app">
      <section className="left">
        <div className="left-head">
          <h1>調書グリッド 試作</h1>
          <p>生成済みの調書を右で確認し、直接編集するか、コメントを付けてAIに該当箇所だけ再生成させます。</p>
          <div className="row">
            <button onClick={() => openSample()} disabled={!wb.ready || wb.busy}>
              サンプル調書を開く
            </button>
            <button onClick={() => openSample("sample_hidden.xlsx")} disabled={!wb.ready || wb.busy} title="非表示シート・非表示行列を含むサンプル">
              非表示ありサンプル
            </button>
            <label className={`btn ${!wb.ready || wb.busy ? "off" : ""}`}>
              xlsxを開く…
              <input type="file" accept=".xlsx" hidden disabled={!wb.ready || wb.busy} onChange={(e) => e.target.files?.[0] && wb.open(e.target.files[0])} />
            </label>
          </div>
          {!wb.ready && <p>計算エンジンを読み込み中…</p>}
        </div>

        <div className="left-tabs">
          <button className={leftTab === "comments" ? "on" : ""} onClick={() => setLeftTab("comments")}>
            コメント <span className="count">{ui.comments.length}</span>
          </button>
          <button className={leftTab === "history" ? "on" : ""} onClick={() => setLeftTab("history")}>
            履歴 <span className="count">{wb.logs.length}</span>
          </button>
        </div>

        {leftTab === "comments" ? (
          <div className="comments">
            {ui.comments.length === 0 ? (
              <p className="hint">
                セルを右クリック →「コメントを付ける」（または Shift+F2）で、AIに直してほしい箇所と内容を指定します。複数付けてからまとめて再生成できます。
              </p>
            ) : (
              <ul>
                {ui.comments.map((c) => (
                  <li key={key(c.sheet, c.row, c.col)} className={selected && same(c, { sheet, ...selected }) ? "on" : ""} onClick={() => wb.focusCell(c)}>
                    <span className="addr">{wb.addr(c)}</span>
                    <span className="text">{c.text}</span>
                    <button
                      className="x"
                      title="コメントを削除"
                      onClick={(e) => {
                        e.stopPropagation();
                        wb.setComment(c, "");
                      }}
                    >
                      ×
                    </button>
                  </li>
                ))}
              </ul>
            )}
            <button className="primary" onClick={wb.regenerate} disabled={!ui.comments.length || wb.regen || ui.proposals.length > 0}>
              {wb.regen ? "再生成中…" : "コメントを反映して再生成"}
            </button>
          </div>
        ) : (
          <div className="history">
            {wb.logs.length === 0 ? (
              <p className="hint">操作するたびに 1 件ずつ記録されます。元に戻す・やり直すも記録として残り、保存時に xlsx の「変更履歴」シートにも書き出されます。</p>
            ) : (
              <ul>
                {[...wb.logs].reverse().map((e) => (
                  <li key={e.seq} onClick={() => e.cells[0] && focusAddr(e.cells[0].addr)}>
                    <span className="addr">#{e.seq}</span>
                    <span className="val">
                      <span className={`who ${e.actor === "ai" ? "ai" : e.action === "undo" || e.action === "redo" || e.action === "save" ? "sys" : ""}`}>
                        {e.actor === "ai" ? "AI" : "手動"}
                      </span>
                      {describe(e)}
                      <br />
                      <span className="when">{hhmmss(e.ts)}</span>
                    </span>
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}

        {ui.proposals.length > 0 && (
          <div className="review">
            <h2>
              AIの更新 <span className="count">{ui.proposals.length}</span>
            </h2>
            <p className="hint">橙のセルが更新箇所です。内容を確認して承認するか、まとめて取り消せます。</p>
            <ul>
              {ui.proposals.map((p) => (
                <li key={key(p.sheet, p.row, p.col)} onClick={() => wb.focusCell(p)}>
                  <span className="addr">{wb.addr(p)}</span>
                  <span className="text">
                    <s>{p.before || "（空）"}</s> → <b>{p.after}</b>
                  </span>
                </li>
              ))}
            </ul>
            <div className="row">
              <button className="primary" onClick={wb.approve}>
                すべて承認
              </button>
              <button onClick={wb.reject}>取り消す</button>
            </div>
          </div>
        )}
      </section>

      <aside className="pane">
        <WorkpaperGrid wb={wb} title={wb.fileName || "調書"} />
      </aside>
    </div>
  );
}
