# -*- coding: utf-8 -*-
"""
pivot_engine.py — Excel 数据透视汇总引擎（零依赖，纯标准库）

自包含（不 import 其它组件）。把一张表按若干维度列分组，对度量列做
sum / count / avg / max / min 聚合，输出汇总表（可写回 .xlsx / .csv）。
替代手工 Excel 透视表，RPA 月度报表、销售汇总、对账统计通用。

作者：支一郎 (Zhiyilang)  ·  许可证：MIT
"""
from __future__ import annotations

import csv
from collections import OrderedDict
from typing import Dict, List, Optional, Sequence
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

# ===========================================================================
# 读
# ===========================================================================

def read_csv(path: str, encoding: str = "utf-8-sig") -> List[Dict[str, str]]:
    with open(path, "r", encoding=encoding, newline="") as fh:
        rows = []
        for raw in csv.DictReader(fh):
            row = {(k or "").strip(): _norm(v) for k, v in raw.items() if k is not None}
            if any(row.values()):
                rows.append(row)
        return rows


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def read_xlsx(path: str) -> List[Dict[str, str]]:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = _shared(z, names)
        target = _first_sheet(z, names)
        if not target or target not in names:
            return []
        cells = _sheet_rows(z.read(target), shared)
    rows: List[Dict[str, str]] = []
    headers: List[str] = []
    for i, cs in enumerate(cells):
        vals = [c["text"] for c in cs]
        if i == 0:
            headers = [str(h).strip() for h in vals]
            continue
        row = {h: (_norm(vals[j]) if j < len(vals) else "") for j, h in enumerate(headers)}
        if any(row.values()):
            rows.append(row)
    return rows


def read_any(path: str) -> List[Dict[str, str]]:
    return read_xlsx(path) if path.lower().endswith(".xlsx") else read_csv(path)


def _shared(z, names):
    n = next((x for x in names if x.endswith("sharedStrings.xml")), None)
    if not n:
        return []
    root = ET.fromstring(z.read(n))
    return ["".join(t.text or "" for t in si.iter(NS + "t")) for si in root.findall(NS + "si")]


def _first_sheet(z, names):
    order = _order(z, names)
    return order[0] if order else next((x for x in names if x.startswith("xl/worksheets/sheet")), None)


def _order(z, names):
    wb = next((n for n in names if n.endswith("workbook.xml")), None)
    if not wb:
        return []
    root = ET.fromstring(z.read(wb))
    id2t = {}
    rels = next((n for n in names if n.endswith("workbook.xml.rels")), None)
    if rels:
        for rel in ET.fromstring(z.read(rels)):
            id2t[rel.get("Id")] = rel.get("Target")
    out = []
    for s in root.iter(NS + "sheet"):
        t = id2t.get(s.get(RNS + "id") or s.get("r:id"), "")
        out.append(t if t.startswith("xl/") else "xl/" + t.lstrip("/"))
    return out


def _sheet_rows(data, shared):
    root = ET.fromstring(data)
    rows = []
    for row in root.iter(NS + "row"):
        cs = []
        for c in row.findall(NS + "c"):
            t = c.get("t")
            v = c.find(NS + "v")
            is_el = c.find(NS + "is")
            if t == "s":
                try:
                    txt = shared[int(v.text or 0)]
                except (ValueError, IndexError):
                    txt = ""
            elif t == "inlineStr" and is_el is not None:
                txt = "".join(x.text or "" for x in is_el.iter(NS + "t"))
            else:
                txt = (v.text or "") if v is not None else ""
            cs.append({"t": t, "text": txt})
        rows.append(cs)
    return rows


# ===========================================================================
# 写
# ===========================================================================

def write_xlsx(rows: Sequence[Dict[str, str]], path: str,
               columns: Optional[Sequence[str]] = None) -> int:
    cols = list(columns) if columns else _union(rows)
    shared: List[str] = []
    def sid(t):
        shared.append(t)
        return len(shared) - 1

    def L(i):
        s = ""; i += 1
        while i:
            i, r = divmod(i - 1, 26)
            s = chr(65 + r) + s
        return s

    sheet = ["".join(f'<c r="{L(i)}1" t="s"><v>{sid(str(h))}</v></c>' for i, h in enumerate(cols))]
    for ri, row in enumerate(rows, start=2):
        cells = []
        for ci, col in enumerate(cols):
            val = _norm(row.get(col, ""))
            num = _num(val)
            if num is not None:
                cells.append(f'<c r="{L(ci)}{ri}"><v>{num}</v></c>')
            else:
                cells.append(f'<c r="{L(ci)}{ri}" t="s"><v>{sid(val)}</v></c>')
        sheet.append("".join(cells))
    sheet_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
                 + "".join(f'<row r="{i+1}">{c}</row>' for i, c in enumerate(sheet))
                 + '</sheetData></worksheet>')
    ss = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          f'count="{len(shared)}" uniqueCount="{len(shared)}">'
          + "".join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in shared) + '</sst>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wbr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wbr)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/sharedStrings.xml", ss)
    return len(rows)


def write_csv(rows, path, columns=None, encoding="utf-8-sig"):
    cols = list(columns) if columns else _union(rows)
    with open(path, "w", encoding=encoding, newline="") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        for r in rows:
            w.writerow([_norm(r.get(c, "")) for c in cols])
    return len(rows)


# ===========================================================================
# 透视
# ===========================================================================

_AGGS = ("sum", "count", "avg", "max", "min")


def pivot(rows: Sequence[Dict[str, str]],
          group_by: Sequence[str],
          aggs: Sequence[Dict[str, str]]) -> Dict:
    """按 group_by 分组，对 aggs 里指定的列做聚合。

    aggs 每项：{"col": "金额", "func": "sum", "as": "金额合计"}
      func ∈ sum/count/avg/max/min
    返回 {"summary": [行], "group_cols": [...], "value_cols": [...], "stats": {...}}
    """
    groups: "OrderedDict[str, Dict]" = OrderedDict()
    for r in rows:
        key = tuple(_norm(r.get(g, "")) for g in group_by)
        bucket = groups.setdefault(key, {"__rows__": 0})
        bucket["__rows__"] += 1
        for a in aggs:
            col, func = a["col"], a["func"]
            if func not in _AGGS:
                raise ValueError(f"未知聚合 func={func}")
            asname = a.get("as") or f"{col}_{func}"
            vals = bucket.setdefault(asname, [])
            if func == "count":
                vals.append(1)
            else:
                num = _num(_norm(r.get(col, "")))
                if num is not None:
                    vals.append(num)

    summary = []
    for key, bucket in groups.items():
        out = {g: k for g, k in zip(group_by, key)}
        for a in aggs:
            asname = a.get("as") or f"{a['col']}_{a['func']}"
            vals = bucket.get(asname, [])
            out[asname] = _reduce(a["func"], vals)
        summary.append(out)

    value_cols = [a.get("as") or f"{a['col']}_{a['func']}" for a in aggs]
    return {
        "summary": summary,
        "group_cols": list(group_by),
        "value_cols": value_cols,
        "stats": {"input_rows": len(rows), "groups": len(summary)},
    }


def _reduce(func, vals):
    if func == "count":
        return len(vals)
    if not vals:
        return 0
    if func == "sum":
        return _clean(sum(vals))
    if func == "avg":
        return _clean(sum(vals) / len(vals))
    if func == "max":
        return _clean(max(vals))
    if func == "min":
        return _clean(min(vals))


def _union(rows):
    cols = []
    for r in rows:
        for c in r:
            if c not in cols:
                cols.append(c)
    return cols


# ===========================================================================
# 工具
# ===========================================================================

def _norm(v) -> str:
    return "" if v is None else " ".join(str(v).split())


def _num(s):
    try:
        if s == "":
            return None
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def _clean(f):
    return int(f) if float(f).is_integer() else round(f, 4)
