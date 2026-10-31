# -*- coding: utf-8 -*-
"""
clean_engine.py — Excel / CSV 数据清洗引擎（零依赖，纯标准库）

自包含（不 import 其它组件），方便单独上架 / 分发到 UiBot Store。
- 读：.csv（utf-8-sig 兼容 BOM）与 .xlsx（zipfile + xml 最小解析）
- 写：.csv 与 .xlsx（最小生成器，输出可被 Excel / 本引擎再次读取）
- 清洗：去首尾/多余空格、删空行、去重（按全行或指定主键）、填充空值、表头归一

作者：支一郎 (Zhiyilang)  ·  许可证：MIT
"""
from __future__ import annotations

import csv
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
# 清洗
# ===========================================================================

def _trim_cell(v: str) -> str:
    """去首尾空格并把内部连续空白折叠为单个空格。"""
    return " ".join(str(v).split())


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


def _union_headers(rows) -> List[str]:
    cols: List[str] = []
    for r in rows:
        for c in r:
            if c not in cols and c != "__source__":
                cols.append(c)
    return cols


def trim_whitespace(rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    """清洗每个单元格的首尾/多余空格，并归一表头。"""
    if not rows:
        return []
    cols = list(rows[0].keys())
    cols = [_trim_cell(c) for c in cols]
    out = []
    for r in rows:
        out.append({c: _trim_cell(r.get(oc, "")) for oc, c in zip(rows[0].keys(), cols)})
    return out


def drop_blank_rows(rows: Sequence[Dict[str, str]]) -> List[Dict[str, str]]:
    """删除所有单元格均为空的行。"""
    return [r for r in rows if any((r.get(c, "") or "").strip() for c in r)]


def drop_duplicates(rows: Sequence[Dict[str, str]],
                    keys: Optional[Sequence[str]] = None) -> List[Dict[str, str]]:
    """去重。keys=None 时按整行去重（保留首次出现）；给定 keys 时按主键去重（保留末次）。"""
    if not rows:
        return []
    cols = list(rows[0].keys())
    if not keys:
        seen = set()
        out = []
        for r in rows:
            sig = tuple(_norm(r.get(c, "")) for c in cols)
            if sig in seen:
                continue
            seen.add(sig)
            out.append(r)
        return out
    out: Dict[tuple, Dict[str, str]] = {}
    for r in rows:
        sig = tuple(_norm(r.get(k, "")) for k in keys)
        out[sig] = r  # 后者覆盖前者
    return list(out.values())


def fill_nulls(rows: Sequence[Dict[str, str]], value: str = "") -> List[Dict[str, str]]:
    """把空单元格（含纯空白）填充为给定值。"""
    out = []
    for r in rows:
        out.append({c: (value if not (r.get(c, "") or "").strip() else r.get(c, ""))
                    for c in r})
    return out


def clean(rows: Sequence[Dict[str, str]],
          trim: bool = True,
          drop_blanks: bool = True,
          dedup: bool = False,
          dedup_keys: Optional[Sequence[str]] = None,
          fill: Optional[str] = None) -> List[Dict[str, str]]:
    """清洗流水线：依次执行 去空格 → 删空行 → 填充 → 去重。

    返回清洗后的行列表。所有操作幂等，可重复调用。
    """
    if trim:
        rows = trim_whitespace(rows)
    if drop_blanks:
        rows = drop_blank_rows(rows)
    if fill is not None:
        rows = fill_nulls(rows, fill)
    if dedup:
        rows = drop_duplicates(rows, dedup_keys)
    return rows


def clean_file(in_path: str, out_path: str,
               trim: bool = True, drop_blanks: bool = True,
               dedup: bool = False, dedup_keys: Optional[Sequence[str]] = None,
               fill: Optional[str] = None) -> Dict:
    """读取任意 .csv/.xlsx，清洗后写出 .xlsx，返回统计。

    统计字段：input_rows / output_rows / removed_blank / removed_dup / filled。
    """
    src = read_any(in_path)
    n_in = len(src)
    if trim:
        src = trim_whitespace(src)
    n_after_trim = len(src)
    if drop_blanks:
        before = len(src)
        src = drop_blank_rows(src)
        removed_blank = before - len(src)
    else:
        removed_blank = 0
    if fill is not None:
        before = len(src)
        src = fill_nulls(src, fill)
        filled = sum(1 for r in src for c in r if r.get(c, "") == fill)
    else:
        filled = 0
    if dedup:
        before = len(src)
        src = drop_duplicates(src, dedup_keys)
        removed_dup = before - len(src)
    else:
        removed_dup = 0
    n_out = write_xlsx(src, out_path)
    return {
        "input_rows": n_in,
        "output_rows": n_out,
        "removed_blank": removed_blank,
        "removed_dup": removed_dup,
        "filled": filled,
    }


# 便于命令行 / UiBot 调用的快捷入口
def run(in_path: str, out_path: str = "cleaned.xlsx", **opts) -> Dict:
    return clean_file(in_path, out_path, **opts)
