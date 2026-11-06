"""正本Word（手で直した版）の別紙１の表を、items.js の内容で組み直す。

表の行・セル・段落の書式は正本の行をそのまま型として使い、文字と赤字だけを入れ替える。
使い方: python3 -I rebuild_table.py <展開したdocxのdocument.xml> <export_rows.jsの出力JSON>
"""
import copy
import json
import re
import sys

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
q = lambda t: "{%s}%s" % (W, t)
RED = "FF0000"

doc_path, rows_path = sys.argv[1], sys.argv[2]
tree = etree.parse(doc_path)
data = json.load(open(rows_path, encoding="utf-8"))

tbl = tree.findall(".//" + q("tbl"))[2]
trs = tbl.findall(q("tr"))
text = lambda el: "".join(el.itertext())


def find_row(pred):
    for i, tr in enumerate(trs):
        if pred(tr):
            return i, tr
    raise SystemExit("型にする行が見つからない")


# ---- 型にする行 ----
_, sec_star = find_row(lambda tr: text(tr).startswith("【学校教育"))
_, sec_norm = find_row(lambda tr: text(tr).startswith("【教務】"))
i, star_first = find_row(lambda tr: "★対話の設定" in text(tr))
star_cont = trs[i + 1]
i, norm_first = find_row(lambda tr: "継続（R7-1）" in text(tr))
norm_cont = trs[i + 1]
i_free, _ = find_row(lambda tr: text(tr).startswith("【自由記述】"))
free_sec = trs[i_free]
free_first, free_cont = trs[i_free + 1], trs[i_free + 3]
# 備考の段落の型（★タグ・区分・補足）
note_paras = star_first.findall(q("tc"))[3].findall(q("p"))
NOTE_T = {"tag": note_paras[0], "note": note_paras[1], "small": note_paras[2]}


def set_runs(p, marked):
    """段落 p の文字を marked（«»が赤字）に置き換える。最初の文字ランの書式を引き継ぐ。"""
    runs = p.findall(q("r"))
    tmpl = None
    for r in runs:
        if r.find(q("t")) is not None:
            tmpl = r
            break
    if tmpl is None and runs:
        tmpl = runs[0]
    rpr = tmpl.find(q("rPr")) if tmpl is not None else None
    for ch in list(p):
        if ch.tag != q("pPr"):
            p.remove(ch)
    for seg in [s for s in re.split(r"(«[^»]*»)", marked) if s]:
        red = seg.startswith("«")
        s = seg[1:-1] if red else seg
        r = etree.SubElement(p, q("r"))
        if rpr is not None:
            nr = copy.deepcopy(rpr)
            for c in nr.findall(q("color")):
                nr.remove(c)
            if red:
                col = etree.Element(q("color"))
                col.set(q("val"), RED)
                # 色は rPr の中で sz より前に置く（スキーマ順）
                anchor = None
                for name in ("sz", "szCs", "highlight", "u", "effect", "vertAlign", "lang", "eastAsianLayout"):
                    anchor = nr.find(q(name))
                    if anchor is not None:
                        break
                if anchor is not None:
                    anchor.addprevious(col)
                else:
                    nr.append(col)
            r.append(nr)
        t = etree.SubElement(r, q("t"))
        t.text = s
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def set_keep(p, keep):
    ppr = p.find(q("pPr"))
    if ppr is None:
        ppr = etree.Element(q("pPr"))
        p.insert(0, ppr)
    for k in ppr.findall(q("keepNext")):
        ppr.remove(k)
    if keep:
        k = etree.Element(q("keepNext"))
        # pStyle の後ろ、それ以外の要素の前
        ps = ppr.find(q("pStyle"))
        if ps is not None:
            ps.addnext(k)
        else:
            ppr.insert(0, k)


def cell_paras(tc, n=1):
    ps = tc.findall(q("p"))
    for extra in ps[n:]:
        tc.remove(extra)
    return tc.findall(q("p"))[:n]


def set_vmerge(tc, mode):
    """mode: None（結合なし）/'restart'/'continue'"""
    tcpr = tc.find(q("tcPr"))
    for v in tcpr.findall(q("vMerge")):
        tcpr.remove(v)
    if mode:
        v = etree.Element(q("vMerge"))
        if mode == "restart":
            v.set(q("val"), "restart")
        gs = tcpr.find(q("gridSpan"))
        (gs if gs is not None else tcpr.find(q("tcW"))).addnext(v)


def keep_all(tr, keep):
    for p in tr.iter(q("p")):
        set_keep(p, keep)


new_rows = []
for r in data["rows"]:
    if r["type"] == "sec":
        tr = copy.deepcopy(sec_star if r["star"] else sec_norm)
        set_runs(cell_paras(tr.findall(q("tc"))[0])[0], r["text"])
        keep_all(tr, True)
    elif r["type"] == "row":
        if r["first"]:
            tr = copy.deepcopy(star_first if r["star"] else norm_first)
        else:
            tr = copy.deepcopy(star_cont if r["star"] else norm_cont)
        tcs = tr.findall(q("tc"))
        multi = r["n"] > 1
        mode = ("restart" if r["first"] else "continue") if multi else None
        set_vmerge(tcs[0], mode)
        set_vmerge(tcs[3], mode)
        set_runs(cell_paras(tcs[0])[0], str(r["no"]) if r["first"] else "")
        set_runs(cell_paras(tcs[1])[0], r["label"])
        set_runs(cell_paras(tcs[2])[0], r["text"])
        tc3 = tcs[3]
        for p in tc3.findall(q("p")):
            tc3.remove(p)
        if r["first"]:
            for kind, s in r["notes"]:
                p = copy.deepcopy(NOTE_T[kind])
                set_runs(p, s)
                tc3.append(p)
        else:
            p = copy.deepcopy(NOTE_T["note"])
            set_runs(p, "")
            tc3.append(p)
        keep_all(tr, not r["last"])
    else:  # free
        tr = copy.deepcopy(free_first if r["first"] else free_cont)
        tcs = tr.findall(q("tc"))
        set_runs(cell_paras(tcs[1])[0], r["label"])
        set_runs(cell_paras(tcs[2])[0], r["text"])
        keep_all(tr, not r["last"])
    new_rows.append(tr)
    if r["type"] == "row" and r is [x for x in data["rows"] if x["type"] == "row"][-1]:
        new_rows.append(copy.deepcopy(free_sec))

# 見出し行（1行目）以外を入れ替える
for tr in trs[1:]:
    tbl.remove(tr)
for tr in new_rows:
    tbl.append(tr)

tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
print("rows:", len(new_rows))
