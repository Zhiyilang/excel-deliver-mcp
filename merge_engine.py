# -*- coding: utf-8 -*-
"""
merge_engine.py — Excel / CSV 多表合并引擎（零依赖，纯标准库）

自包含（不 import 其它组件），方便单独上架 / 分发。
- 读：.csv（utf-8-sig 兼容 BOM）与 .xlsx（zipfile+xml 最小解析）
- 写：.csv 与 .xlsx（最小生成器，输出可被 Excel / 本引擎再次读取）
- merge()：把多个文件堆叠合并，可选按主键去重、表头归一化。

作者：支一郎 (Zhiyilang)  ·  许可证：MIT
"""
from __future__ import annotations

import csv
import io
import os
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Sequence

# ===========================================================================
# 读
# ===========================================================================

def read_csv(path: str, encoding: str = "utf-8-sig") -> List[Dict[str, str]]:
    with open(path, "r", encoding=encoding, newline="") as fh:
        return _rows_from_reader(csv.DictReader(fh))


def _rows_from_reader(reader) -> List[Dict[str, str]]:
    rows = []
    for raw in reader:
        row = {(k or "").strip(): _norm(v) for k, v in raw.items() if k is not None}
        if any(row.values()):
            rows.append(row)
    return rows


def read_xlsx(path: str) -> List[Dict[str, str]]:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = _shared_strings(z, names)
        target = _first_sheet(z, names)
        if not target or target not in names:
            return []
        cells_per_row = _sheet_rows(z.read(target), shared)
    rows: List[Dict[str, str]] = []
    headers: List[str] = []
    for r_i, cells in enumerate(cells_per_row):
        vals = [c["text"] for c in cells]
        if r_i == 0:
            headers = [str(h).strip() for h in vals]
            continue
        row = {h: (_norm(vals[i]) if i < len(vals) else "") for i, h in enumerate(headers)}
        if any(row.values()):
            rows.append(row)
    return rows


def read_any(path: str) -> List[Dict[str, str]]:
    return read_xlsx(path) if path.lower().endswith(".xlsx") else read_csv(path)


# ---- xlsx 内部解析 ----
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _shared_strings(z, names) -> List[str]:
    n = next((x for x in names if x.endswith("sharedStrings.xml")), None)
    if not n:
        return []
    root = ET.fromstring(z.read(n))
    return ["".join(t.text or "" for t in si.iter(NS + "t")) for si in root.findall(NS + "si")]


def _first_sheet(z, names) -> Optional[str]:
    order = _sheet_order(z, names)
    return order[0] if order else next((x for x in names if x.startswith("xl/worksheets/sheet")), None)


def _sheet_order(z, names) -> List[str]:
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
        tgt = id2t.get(s.get(RNS + "id") or s.get("r:id"), "")
        out.append(tgt if tgt.startswith("xl/") else "xl/" + tgt.lstrip("/"))
    return out


def _sheet_rows(data: bytes, shared: List[str]) -> List[List[Dict]]:
    root = ET.fromstring(data)
    rows = []
    for row in root.iter(NS + "row"):
        cells = []
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
            cells.append({"t": t, "text": txt})
        rows.append(cells)
    return rows


# ===========================================================================
# 写
# ===========================================================================

def write_csv(rows: Sequence[Dict[str, str]], path: str,
              columns: Optional[Sequence[str]] = None, encoding: str = "utf-8-sig") -> int:
    cols = list(columns) if columns else _union_headers(rows)
    with open(path, "w", encoding=encoding, newline="") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        for r in rows:
            w.writerow([_norm(r.get(c, "")) for c in cols])
    return len(rows)


def write_xlsx(rows: Sequence[Dict[str, str]], path: str,
               columns: Optional[Sequence[str]] = None, sheet_name: str = "Sheet1") -> int:
    cols = list(columns) if columns else _union_headers(rows)
    shared: List[str] = []
    def sid(text):
        shared.append(text)
        return len(shared) - 1

    def col_letter(idx):
        s = ""; idx += 1
        while idx:
            idx, r = divmod(idx - 1, 26)
            s = chr(65 + r) + s
        return s

    sheet_rows = ["".join(f'<c r="{col_letter(i)}1" t="s"><v>{sid(str(h))}</v></c>'
                          for i, h in enumerate(cols))]
    for ri, row in enumerate(rows, start=2):
        cells = []
        for ci, col in enumerate(cols):
            val = _norm(row.get(col, ""))
            num = _as_number(val)
            if num is not None:
                cells.append(f'<c r="{col_letter(ci)}{ri}"><v>{num}</v></c>')
            else:
                cells.append(f'<c r="{col_letter(ci)}{ri}" t="s"><v>{sid(val)}</v></c>')
        sheet_rows.append("".join(cells))

    from xml.sax.saxutils import escape
    sheet_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                 '<sheetData>'
                 + "".join(f'<row r="{i+1}">{c}</row>' for i, c in enumerate(sheet_rows))
                 + '</sheetData></worksheet>')
    shared_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                  f'count="{len(shared)}" uniqueCount="{len(shared)}">'
                  + "".join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in shared)
                  + '</sst>')
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
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          f'<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
               '</Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/sharedStrings.xml", shared_xml)
    return len(rows)


# ===========================================================================
# 合并
# ===========================================================================

def merge(paths: Sequence[str], key: Optional[str] = None, dedup: bool = False,
          case_insensitive_key: bool = True) -> Dict:
    """把多个文件合并成一张表。

    - key/dedup：给定主键且 dedup=True 时，按主键保留后出现的行（通常最新）。
    - 返回 {'rows', 'columns', 'sources': {path: 行数}, 'stats'}。
    """
    per_file: Dict[str, List[Dict[str, str]]] = {}
    for p in paths:
        per_file[p] = read_any(p)

    columns = _union_headers([r for rows in per_file.values() for r in rows])

    merged: List[Dict[str, str]] = []
    if dedup and key:
        by_key: Dict[str, Dict[str, str]] = {}
        for p in paths:
            for r in per_file[p]:
                kv = _norm(r.get(key, ""))
                kv = kv.lower() if case_insensitive_key else kv
                r["__source__"] = os.path.basename(p)
                by_key[kv] = r  # 后者覆盖前者
        merged = list(by_key.values())
    else:
        for p in paths:
            for r in per_file[p]:
                r["__source__"] = os.path.basename(p)
                merged.append(r)

    return {
        "rows": merged,
        "columns": [c for c in columns if c != "__source__"],
        "sources": {os.path.basename(p): len(per_file[p]) for p in paths},
        "stats": {"input_rows": sum(len(v) for v in per_file.values()),
                  "output_rows": len(merged),
                  "files": len(paths), "dedup": bool(dedup and key)},
    }


def _union_headers(rows) -> List[str]:
    cols: List[str] = []
    for r in rows:
        for c in r:
            if c not in cols and c != "__source__":
                cols.append(c)
    return cols


# ===========================================================================
# 工具
# ===========================================================================

def _norm(v) -> str:
    if v is None:
        return ""
    return " ".join(str(v).split())


def _as_number(s: str):
    try:
        if s == "":
            return None
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None
