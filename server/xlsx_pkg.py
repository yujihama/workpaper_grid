"""xlsx パッケージ（zip + OOXML）を直接扱う。IronCalc が持たない要素を残すための層。

- strip_for_engine: メモ（comments + VML）を取り除いた IronCalc 用コピーを作る（IronCalc はメモ付きブックを開けない）
- read_notes / read_assets: 元ファイルからメモ・画像・グラフ・図形・リンク・入力規則を読む（画面表示用）
- graft: 元ファイルを土台に、IronCalc の出力からセル部分（sheetData / cols / 結合 / 条件付き書式 / styles / sharedStrings）だけを移植し、
         図形・画像・メモ・フィルタ・入力規則・リンク・マクロ等は元のまま残す。行列の挿入・削除・移動はそれらの位置にも再生する
"""
from __future__ import annotations

import base64
import io
import posixpath
import re
import zipfile
from dataclasses import dataclass, field

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REL_WS, REL_COMMENTS, REL_VML, REL_DRAWING, REL_HYPERLINK, REL_IMAGE, REL_CHART, REL_SST, REL_STYLES, REL_CALC = (
    REL + "worksheet", REL + "comments", REL + "vmlDrawing", REL + "drawing", REL + "hyperlink", REL + "image", REL + "chart",
    REL + "sharedStrings", REL + "styles", REL + "calcChain")
CT_SST = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_META = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"

# CT_Worksheet の子要素の順序（スキーマで固定。差し替え時にこの順に挿入する）
WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
            "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
            "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
            "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
            "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def local(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def resolve(base_part: str, target: str) -> str:
    """rels の Target を zip 内のパーツ名に直す（絶対 /xl/... と相対 ../drawings/... の両方）"""
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def rels_path(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


# ---- パッケージ --------------------------------------------------------------
class Pkg:
    def __init__(self, data: bytes):
        self.parts: dict[str, bytes] = {}
        self.order: list[str] = []
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for n in z.namelist():
                if n.endswith("/"):
                    continue
                self.parts[n] = z.read(n)
                self.order.append(n)

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            names = [n for n in self.order if n in self.parts] + [n for n in self.parts if n not in self.order]
            if "[Content_Types].xml" in names:  # 先頭に置くのが慣例
                names.remove("[Content_Types].xml")
                names.insert(0, "[Content_Types].xml")
            for n in names:
                z.writestr(n, self.parts[n])
        return buf.getvalue()

    def xml(self, part: str):
        return etree.fromstring(self.parts[part])

    def put_xml(self, part: str, root) -> None:
        self.parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if part not in self.order:
            self.order.append(part)

    def remove(self, part: str) -> None:
        self.parts.pop(part, None)
        self.parts.pop(rels_path(part), None)

    # rels
    def rels(self, part: str) -> list[dict]:
        p = rels_path(part)
        if p not in self.parts:
            return []
        out = []
        for e in self.xml(p).findall(q(NS_PR, "Relationship")):
            out.append({"id": e.get("Id"), "type": e.get("Type"), "target": e.get("Target"), "mode": e.get("TargetMode"),
                        "part": None if e.get("TargetMode") == "External" else resolve(part, e.get("Target"))})
        return out

    def remove_rels(self, part: str, pred) -> None:
        p = rels_path(part)
        if p not in self.parts:
            return
        root = self.xml(p)
        for e in list(root):
            if pred(e):
                root.remove(e)
        self.put_xml(p, root)

    def add_rel(self, part: str, rel_type: str, target: str, rid: str | None = None) -> str:
        p = rels_path(part)
        root = self.xml(p) if p in self.parts else etree.Element(q(NS_PR, "Relationships"), nsmap={None: NS_PR})
        ids = {e.get("Id") for e in root}
        if rid is None:
            i = 1
            while f"rId{i}" in ids:
                i += 1
            rid = f"rId{i}"
        etree.SubElement(root, q(NS_PR, "Relationship"), Id=rid, Type=rel_type, Target=target)
        self.put_xml(p, root)
        return rid

    # content types
    def set_override(self, part: str, ct: str) -> None:
        root = self.xml("[Content_Types].xml")
        pn = "/" + part
        for e in root.findall(q(NS_CT, "Override")):
            if e.get("PartName") == pn:
                e.set("ContentType", ct)
                break
        else:
            etree.SubElement(root, q(NS_CT, "Override"), PartName=pn, ContentType=ct)
        self.put_xml("[Content_Types].xml", root)

    def drop_override(self, part: str) -> None:
        root = self.xml("[Content_Types].xml")
        for e in root.findall(q(NS_CT, "Override")):
            if e.get("PartName") == "/" + part:
                root.remove(e)
        self.put_xml("[Content_Types].xml", root)

    # workbook
    def sheets(self) -> list[dict]:
        """workbook.xml の順に (name, sheetId, rId, part)"""
        wb = self.xml("xl/workbook.xml")
        rels = {r["id"]: r for r in self.rels("xl/workbook.xml")}
        out = []
        for s in wb.find(q(NS_MAIN, "sheets")):
            rid = s.get(q(NS_R, "id"))
            out.append({"name": s.get("name"), "sheetId": s.get("sheetId"), "rid": rid, "part": rels[rid]["part"], "state": s.get("state")})
        return out


# ---- 1. IronCalc 用にメモを取り除く ---------------------------------------------
def strip_for_engine(data: bytes) -> bytes:
    pkg = Pkg(data)
    for sh in pkg.sheets():
        part = sh["part"]
        drop = [r for r in pkg.rels(part) if r["type"] in (REL_COMMENTS, REL_VML)]
        if not drop:
            continue
        for r in drop:
            if r["part"]:
                pkg.parts.pop(r["part"], None)
                pkg.drop_override(r["part"])
        ids = {r["id"] for r in drop}
        pkg.remove_rels(part, lambda e: e.get("Id") in ids)
        # シート XML の <legacyDrawing r:id="..."/> を消す（文字列処理で他の部分に触らない）
        xml = pkg.parts[part]
        xml = re.sub(rb"<(?:\w+:)?legacyDrawing\b[^>]*/>", b"", xml)
        pkg.parts[part] = xml
    # 旧形式の VML 既定 (Default Extension="vml") は残っていても害はない
    return pkg.to_bytes()


# ---- 2. 読み取り（画面表示用） ------------------------------------------------------
def _cell_ref(ref: str) -> tuple[int, int]:
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return int(m.group(2)), col


def _col_name(c: int) -> str:
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def read_notes(data: bytes) -> list[dict]:
    """Excel のメモ（旧コメント）: [{sheet, row, col, author, text}]"""
    pkg = Pkg(data)
    out = []
    for si, sh in enumerate(pkg.sheets()):
        for r in pkg.rels(sh["part"]):
            if r["type"] != REL_COMMENTS or r["part"] not in pkg.parts:
                continue
            root = pkg.xml(r["part"])
            au = root.find(q(NS_MAIN, "authors"))
            authors = [a.text or "" for a in au] if au is not None else []
            lst = root.find(q(NS_MAIN, "commentList"))
            for c in lst if lst is not None else []:
                row, col = _cell_ref(c.get("ref"))
                text = "".join(t.text or "" for t in c.iter(q(NS_MAIN, "t")))
                a = int(c.get("authorId", "0"))
                out.append({"sheet": si, "row": row, "col": col, "author": authors[a] if a < len(authors) else "", "text": text})
    return out


def _anchor(el) -> dict | None:
    if el is None:
        return None
    g = lambda t: (el.find(q(NS_XDR, t)).text if el.find(q(NS_XDR, t)) is not None else "0")
    return {"col": int(g("col")) + 1, "row": int(g("row")) + 1, "colOff": int(g("colOff")), "rowOff": int(g("rowOff"))}  # 1始まりに


def read_assets(data: bytes) -> dict:
    """画像・グラフ・図形（位置とプレビュー）、リンク、入力規則。座標は 1 始まり、オフセットは EMU"""
    pkg = Pkg(data)
    images, charts, shapes, hyperlinks, validations, defaults = [], [], [], [], [], []
    for si, sh in enumerate(pkg.sheets()):
        rels = {r["id"]: r for r in pkg.rels(sh["part"])}
        root = pkg.xml(sh["part"])
        # 既定の行高（pt）・列幅（文字数）。IronCalc は未指定の行列に独自の既定値を返すので、画面はこちらを使う
        fp = root.find(q(NS_MAIN, "sheetFormatPr"))
        row_pt = float(fp.get("defaultRowHeight", "15")) if fp is not None else 15.0
        if fp is not None and fp.get("defaultColWidth"):
            col_chars = float(fp.get("defaultColWidth"))
        else:
            col_chars = float(fp.get("baseColWidth", "8")) + 0.71 if fp is not None else 8.43  # 余白込みで 8.43 文字
        defaults.append({"sheet": si, "rowPx": row_pt * 4 / 3, "colPx": int(col_chars * 7 + 5)})
        for h in root.iter(q(NS_MAIN, "hyperlink")):
            rid = h.get(q(NS_R, "id"))
            target = rels[rid]["target"] if rid and rid in rels else ("#" + h.get("location", ""))
            hyperlinks.append({"sheet": si, "ref": h.get("ref"), "target": target, "display": h.get("display") or ""})
        for dv in root.iter(q(NS_MAIN, "dataValidation")):
            f1 = dv.find(q(NS_MAIN, "formula1"))
            validations.append({"sheet": si, "sqref": dv.get("sqref"), "type": dv.get("type"), "formula1": f1.text if f1 is not None else ""})
        for r in rels.values():
            if r["type"] != REL_DRAWING or r["part"] not in pkg.parts:
                continue
            dpart = r["part"]
            drels = {x["id"]: x for x in pkg.rels(dpart)}
            droot = pkg.xml(dpart)
            for anc in droot:
                kind = local(anc.tag)
                if kind not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                    continue
                frm = _anchor(anc.find(q(NS_XDR, "from")))
                to = _anchor(anc.find(q(NS_XDR, "to")))
                if frm is None:
                    continue
                ext_el = anc.find(q(NS_XDR, "ext"))  # oneCellAnchor は大きさを EMU で持つ
                ext = {"cx": int(ext_el.get("cx", "0")), "cy": int(ext_el.get("cy", "0"))} if ext_el is not None else None
                pic = anc.find(q(NS_XDR, "pic"))
                gf = anc.find(q(NS_XDR, "graphicFrame"))
                sp = anc.find(q(NS_XDR, "sp"))
                if pic is not None:
                    blip = pic.find(f".//{q(NS_A, 'blip')}")
                    emb = blip.get(q(NS_R, "embed")) if blip is not None else None
                    mp = drels[emb]["part"] if emb in drels else None
                    if mp and mp in pkg.parts:
                        suffix = mp.rsplit(".", 1)[-1].lower()
                        mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp", "emf": "image/emf", "svg": "image/svg+xml"}.get(suffix, "application/octet-stream")
                        images.append({"sheet": si, "from": frm, "to": to, "ext": ext, "dataUrl": f"data:{mime};base64," + base64.b64encode(pkg.parts[mp]).decode()})
                elif gf is not None:
                    name = gf.find(f".//{q(NS_XDR, 'cNvPr')}")
                    charts.append({"sheet": si, "from": frm, "to": to, "ext": ext, "name": name.get("name") if name is not None else "chart", "kind": "chart"})
                elif sp is not None:
                    name = sp.find(f".//{q(NS_XDR, 'cNvPr')}")
                    text = "".join(t.text or "" for t in sp.iter(q(NS_A, "t")))
                    shapes.append({"sheet": si, "from": frm, "to": to, "ext": ext, "name": name.get("name") if name is not None else "shape", "text": text, "kind": "shape"})
    return {"images": images, "charts": charts, "shapes": shapes, "hyperlinks": hyperlinks, "validations": validations, "defaults": defaults}


# ---- 3. 構造操作の再生（行列の挿入・削除・移動を、温存した要素の位置に適用） ----------------
@dataclass
class StructOp:
    kind: str      # insert_rows / delete_rows / insert_cols / delete_cols / move_rows / move_cols
    at: int        # 1 始まり
    n: int
    delta: int = 0


def _shift(v: int, op: StructOp, axis: str) -> int | None:
    """1 始まりの行番号/列番号を op でずらす。削除で消えたら None"""
    if not op.kind.endswith("rows" if axis == "row" else "cols"):
        return v
    at, n, d = op.at, op.n, op.delta
    if op.kind.startswith("insert"):
        return v + n if v >= at else v
    if op.kind.startswith("delete"):
        return v if v < at else (None if v < at + n else v - n)
    # move
    if at <= v < at + n:
        return v + d
    if d > 0 and at + n <= v < at + n + d:
        return v - n
    if d < 0 and at + d <= v < at:
        return v + n
    return v


def _shift_range(ref: str, op: StructOp) -> str | None:
    """'A1' / 'A1:C5' / 'A:C' / '3:5' をずらす。全体が削除されたら None、一部なら縮める"""
    def one(cell: str):
        m = re.match(r"^\$?([A-Z]*)\$?(\d*)$", cell)
        col = _cell_ref(m.group(1) + "1")[1] if m.group(1) else None
        row = int(m.group(2)) if m.group(2) else None
        return row, col

    def fmt(row, col):
        return (_col_name(col) if col else "") + (str(row) if row else "")

    parts = ref.split(":")
    (r1, c1) = one(parts[0])
    (r2, c2) = one(parts[-1])
    axis = "row" if op.kind.endswith("rows") else "col"
    a1, a2 = (r1, r2) if axis == "row" else (c1, c2)
    if a1 is None:  # この軸の指定が無い（列全体など）ならそのまま
        return ref
    if op.kind.startswith("delete"):
        lo, hi = a1, a2
        if lo >= op.at and hi < op.at + op.n:
            return None
        n1 = a1 if a1 < op.at else (op.at if a1 < op.at + op.n else a1 - op.n)
        n2 = a2 if a2 < op.at else (op.at - 1 if a2 < op.at + op.n else a2 - op.n)
    else:
        n1, n2 = _shift(a1, op, axis), _shift(a2, op, axis)
        if n1 is None or n2 is None:
            return None
        if n1 > n2:
            n1, n2 = n2, n1
    if axis == "row":
        r1, r2 = n1, n2
    else:
        c1, c2 = n1, n2
    return fmt(r1, c1) if len(parts) == 1 else f"{fmt(r1, c1)}:{fmt(r2, c2)}"


def _shift_sqref(sqref: str, ops: list[StructOp]) -> str | None:
    out = []
    for ref in sqref.split():
        r = ref
        for op in ops:
            r = _shift_range(r, op)
            if r is None:
                break
        if r:
            out.append(r)
    return " ".join(out) or None


def _apply_ops_to_sheet(pkg: Pkg, part: str, root, ops: list[StructOp]) -> None:
    """温存した要素（フィルタ・入力規則・リンク・メモ・図形アンカー）の位置に構造操作を再生する"""
    if not ops:
        return
    af = root.find(q(NS_MAIN, "autoFilter"))
    if af is not None:
        nr = _shift_sqref(af.get("ref"), ops)
        if nr:
            af.set("ref", nr)
        else:
            root.remove(af)
    dvs = root.find(q(NS_MAIN, "dataValidations"))
    if dvs is not None:
        for dv in list(dvs):
            nr = _shift_sqref(dv.get("sqref"), ops)
            if nr:
                dv.set("sqref", nr)
            else:
                dvs.remove(dv)
        if len(dvs) == 0:
            root.remove(dvs)
        else:
            dvs.set("count", str(len(dvs)))
    hls = root.find(q(NS_MAIN, "hyperlinks"))
    if hls is not None:
        for h in list(hls):
            nr = _shift_sqref(h.get("ref"), ops)
            if nr:
                h.set("ref", nr)
            else:
                hls.remove(h)
        if len(hls) == 0:
            root.remove(hls)
    for r in pkg.rels(part):
        if r["type"] == REL_COMMENTS and r["part"] in pkg.parts:
            croot = pkg.xml(r["part"])
            lst = croot.find(q(NS_MAIN, "commentList"))
            for c in list(lst):
                nr = _shift_sqref(c.get("ref"), ops)
                if nr:
                    c.set("ref", nr)
                else:
                    lst.remove(c)
            pkg.put_xml(r["part"], croot)
        elif r["type"] == REL_VML and r["part"] in pkg.parts:
            # VML の ClientData は 0 始まりの <x:Row>/<x:Column>（接頭辞は書き手により異なる）。番号だけ文字列で書き換える
            def fix(m):
                pre, tag, v = m.group(1) or "", m.group(2), int(m.group(3))
                axis = "row" if tag == "Row" else "col"
                nv = v + 1
                for op in ops:
                    nv2 = _shift(nv, op, axis)
                    nv = nv2 if nv2 is not None else op.at  # 消えた行のメモは削除位置に寄せる
                return f"<{pre}{tag}>{max(nv, 1) - 1}</{pre}{tag}>"
            pkg.parts[r["part"]] = re.sub(r"<(\w+:)?(Row|Column)>(\d+)</\1?\2>", fix, pkg.parts[r["part"]].decode("utf-8", "ignore")).encode()
        elif r["type"] == REL_DRAWING and r["part"] in pkg.parts:
            droot = pkg.xml(r["part"])
            for anc in droot:
                for tag in ("from", "to"):
                    a = anc.find(q(NS_XDR, tag))
                    if a is None:
                        continue
                    for axis, el in (("row", a.find(q(NS_XDR, "row"))), ("col", a.find(q(NS_XDR, "col")))):
                        if el is None:
                            continue
                        v = int(el.text) + 1
                        for op in ops:
                            nv = _shift(v, op, axis)
                            v = nv if nv is not None else op.at
                        el.text = str(max(v, 1) - 1)
            pkg.put_xml(r["part"], droot)


# ---- 4. 差し替え保存 -----------------------------------------------------------------
def _insert_ordered(root, el) -> None:
    """CT_Worksheet の順序を守って el を挿入する"""
    name = local(el.tag)
    pos = WS_ORDER.index(name) if name in WS_ORDER else len(WS_ORDER)
    idx = len(root)
    for i, child in enumerate(root):
        cn = local(child.tag)
        if (WS_ORDER.index(cn) if cn in WS_ORDER else len(WS_ORDER)) > pos:
            idx = i
            break
    root.insert(idx, el)


def _replace_children(dst_root, src_root, tags: list[str]) -> None:
    for t in tags:
        for e in dst_root.findall(q(NS_MAIN, t)):
            dst_root.remove(e)
        for e in src_root.findall(q(NS_MAIN, t)):
            _insert_ordered(dst_root, e)


def _set_merges(root, merges: list[tuple[int, int, int, int]]) -> None:
    for e in root.findall(q(NS_MAIN, "mergeCells")):
        root.remove(e)
    if not merges:
        return
    mc = etree.Element(q(NS_MAIN, "mergeCells"), count=str(len(merges)))
    for r1, c1, r2, c2 in merges:
        etree.SubElement(mc, q(NS_MAIN, "mergeCell"), ref=f"{_col_name(c1)}{r1}:{_col_name(c2)}{r2}")
    _insert_ordered(root, mc)


def graft(original: bytes, engine_out: bytes, orig_names: list[str | None], merges: list[list[int]], ops: list[dict]) -> bytes:
    """original を土台に engine_out（IronCalc の出力）のセル部分を移植した xlsx を返す。
    orig_names[i]: 出力 i 番目のシートが元ファイルのどのシート名だったか（新規シートは None）
    merges: [sheet, r1, c1, r2, c2]（出力側のシート番号）
    ops: [{sheet_orig_name, kind, at, n, delta}] 構造操作（時系列）
    """
    o, n = Pkg(original), Pkg(engine_out)
    o_sheets = {s["name"]: s for s in o.sheets()}
    n_sheets = n.sheets()

    # 共通パーツ: styles / sharedStrings / metadata は IronCalc のものに置き換える（セルの s= 添字がこれを指す）
    o.parts["xl/styles.xml"] = n.parts["xl/styles.xml"]
    wb_rels = o.rels("xl/workbook.xml")
    if "xl/sharedStrings.xml" in n.parts:
        o.parts["xl/sharedStrings.xml"] = n.parts["xl/sharedStrings.xml"]
        if not any(r["type"] == REL_SST for r in wb_rels):
            o.add_rel("xl/workbook.xml", REL_SST, "sharedStrings.xml")
        o.set_override("xl/sharedStrings.xml", CT_SST)
    if "xl/metadata.xml" in n.parts:
        o.parts["xl/metadata.xml"] = n.parts["xl/metadata.xml"]
        if not any(r["target"].endswith("metadata.xml") for r in wb_rels):
            o.add_rel("xl/workbook.xml", REL + "sheetMetadata", "metadata.xml")
        o.set_override("xl/metadata.xml", CT_META)
    # calcChain は Excel が作り直す（残すと不整合で修復ダイアログが出る）
    for r in wb_rels:
        if r["type"] == REL_CALC:
            o.remove(r["part"])
            o.drop_override(r["part"])
            o.remove_rels("xl/workbook.xml", lambda e, rid=r["id"]: e.get("Id") == rid)

    used_parts: set[str] = set()
    new_sheets = []  # (name, part, rId)
    by_orig: dict[str, list[StructOp]] = {}
    for op in ops:
        if op.get("sheet_orig_name"):
            by_orig.setdefault(op["sheet_orig_name"], []).append(StructOp(op["kind"], int(op["at"]), int(op["n"]), int(op.get("delta", 0))))
    merges_by_sheet: dict[int, list[tuple[int, int, int, int]]] = {}
    for si, r1, c1, r2, c2 in merges:
        merges_by_sheet.setdefault(si, []).append((r1, c1, r2, c2))

    for i, ns in enumerate(n_sheets):
        oname = orig_names[i] if i < len(orig_names) else None
        nroot = n.xml(ns["part"])
        if oname and oname in o_sheets:
            part = o_sheets[oname]["part"]
            oroot = o.xml(part)
            _replace_children(oroot, nroot, ["dimension", "sheetFormatPr", "cols", "sheetData", "mergeCells", "conditionalFormatting"])
            _apply_ops_to_sheet(o, part, oroot, by_orig.get(oname, []))
            if i in merges_by_sheet or oroot.find(q(NS_MAIN, "mergeCells")) is not None:
                _set_merges(oroot, merges_by_sheet.get(i, []))
            o.put_xml(part, oroot)
            used_parts.add(part)
            new_sheets.append((ns["name"], part, o_sheets[oname]["rid"], o_sheets[oname]["sheetId"], ns["state"]))
        else:
            # 新規シート（IronCalc が作ったもの・変更履歴）は出力側の XML をそのまま持ち込む
            k = 1
            while f"xl/worksheets/sheet{k}.xml" in o.parts or f"xl/worksheets/sheet{k}.xml" in used_parts:
                k += 1
            part = f"xl/worksheets/sheet{k}.xml"
            _set_merges(nroot, merges_by_sheet.get(i, [])) if i in merges_by_sheet else None
            o.put_xml(part, nroot)
            o.set_override(part, CT_WS)
            rid = o.add_rel("xl/workbook.xml", REL_WS, "worksheets/" + posixpath.basename(part))
            used_parts.add(part)
            new_sheets.append((ns["name"], part, rid, None, ns["state"]))

    # 出力に無い元シートは外す（削除されたシート）
    for name, s in o_sheets.items():
        if s["part"] not in used_parts:
            o.remove(s["part"])
            o.drop_override(s["part"])
            o.remove_rels("xl/workbook.xml", lambda e, rid=s["rid"]: e.get("Id") == rid)

    # workbook.xml: <sheets> を出力の順・名前に合わせ、definedNames は IronCalc のもの（行列操作で更新済み）にする
    wb = o.xml("xl/workbook.xml")
    sheets_el = wb.find(q(NS_MAIN, "sheets"))
    for e in list(sheets_el):
        sheets_el.remove(e)
    max_id = max([int(s["sheetId"]) for s in o_sheets.values()] + [0])
    for name, part, rid, sheet_id, state in new_sheets:
        if sheet_id is None:
            max_id += 1
            sheet_id = str(max_id)
        e = etree.SubElement(sheets_el, q(NS_MAIN, "sheet"), name=name, sheetId=sheet_id)
        e.set(q(NS_R, "id"), rid)
        if state and state != "visible":
            e.set("state", state)
    ndn = n.xml("xl/workbook.xml").find(q(NS_MAIN, "definedNames"))
    odn = wb.find(q(NS_MAIN, "definedNames"))
    if odn is not None:
        wb.remove(odn)
    if ndn is not None and len(ndn):
        idx = list(wb).index(sheets_el) + 1
        wb.insert(idx, ndn)
    o.put_xml("xl/workbook.xml", wb)
    return o.to_bytes()
