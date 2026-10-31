#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_deliver_mcp/server.py — 纯标准库 MCP 服务器（stdio 传输）

把 Excel 交付引擎（合并 / 清洗 / 透视 / 拆分 / 双表对账）暴露为 MCP 工具，
让任意兼容 MCP 的 LLM Agent（Claude Desktop / Cursor / 自建客户端）可直接调用。

特性：
- 零第三方依赖（仅 Python 标准库）。
- 实现 MCP 的 stdio 传输：基于 Content-Length 的 JSON-RPC 2.0 帧。
- 支持方法：initialize / tools/list / tools/call / ping。
- 工具执行本身调用本地 vendored 引擎，文件读写在当前工作目录，绝不联网。

运行：
  python server.py            # 作为 MCP server 常驻，从 stdin 读请求
  python eval_mcp.py          # 自测（会作为子进程拉起本 server 并断言）

作者：支一郎 (Zhiyilang) · 许可证：MIT
"""
import json
import os
import sys
import io

# 让本目录的引擎可被导入（与 server.py 同目录）
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import merge_engine as ME   # noqa: E402
import pivot_engine as PE   # noqa: E402
import clean_engine as CE   # noqa: E402


def write_any(path, rows, columns=None):
    """按扩展名选择 csv / xlsx 写出（对齐 merge_engine 的 write_* 签名）。"""
    if str(path).lower().endswith(".csv"):
        return ME.write_csv(rows, path, columns=columns)
    return ME.write_xlsx(rows, path, columns=columns)


# ---------------------------------------------------------------------------
# 工具定义（MCP tools/list 返回）
# ---------------------------------------------------------------------------
TOOLS = [
    {
        "name": "excel_merge",
        "description": "合并多个 .xlsx/.csv 文件为一个总表。支持表头并集、按主键去重、标注来源文件。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "inputs": {"type": "array", "items": {"type": "string"},
                           "description": "待合并的文件路径列表（至少 1 个）"},
                "output": {"type": "string", "description": "输出文件路径，如 merged.xlsx"},
                "key": {"type": "string", "description": "去重主键列名；留空表示按整行去重"},
                "dedup": {"type": "boolean", "description": "是否按主键去重，默认 false"},
            },
            "required": ["inputs", "output"],
        },
    },
    {
        "name": "excel_clean",
        "description": "清洗 Excel/CSV：去首尾空格、删空行、按列去重、空值填充。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "input": {"type": "string", "description": "输入文件路径"},
                "output": {"type": "string", "description": "输出文件路径"},
                "key": {"type": "string", "description": "去重依据列，逗号分隔；留空按整行"},
                "dedup": {"type": "boolean", "description": "是否去重，默认 false"},
                "fill": {"type": "string", "description": "空值统一填充为的文本内容"},
            },
            "required": ["input", "output"],
        },
    },
    {
        "name": "excel_pivot",
        "description": "按维度列分组做聚合（sum/count/avg/max/min），输出透视汇总表。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "input": {"type": "string", "description": "输入文件路径"},
                "output": {"type": "string", "description": "输出文件路径"},
                "by": {"type": "string", "description": "分组列，逗号分隔，如 城市"},
                "agg": {"type": "array", "items": {"type": "string"},
                        "description": "聚合项，格式 列名:func[:别名]，如 金额:sum:金额合计"},
            },
            "required": ["input", "output", "by", "agg"],
        },
    },
    {
        "name": "excel_split",
        "description": "按某列的值将一张总表拆分成多个文件，文件名即该列的值。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "input": {"type": "string", "description": "输入文件路径"},
                "by": {"type": "string", "description": "拆分依据列，如 城市"},
                "outdir": {"type": "string", "description": "输出目录"},
                "xlsx": {"type": "boolean", "description": "true 输出 xlsx，false 输出 csv（默认）"},
            },
            "required": ["input", "by", "outdir"],
        },
    },
    {
        "name": "excel_recon",
        "description": "双表按主键逐行比对，输出三张结果：仅在A/仅在B/数值不一致（精确到单元格）。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "a": {"type": "string", "description": "表 A 文件路径"},
                "b": {"type": "string", "description": "表 B 文件路径"},
                "key": {"type": "string", "description": "主键列名，如 订单号"},
                "outdir": {"type": "string", "description": "输出目录"},
                "xlsx": {"type": "boolean", "description": "true 输出 xlsx，false 输出 csv（默认）"},
            },
            "required": ["a", "b", "key", "outdir"],
        },
    },
]


# ---------------------------------------------------------------------------
# 工具实现
# ---------------------------------------------------------------------------
def _out(text):
    return {"content": [{"type": "text", "text": text}]}


def run_tool(name, args):
    if name == "excel_merge":
        res = ME.merge(args["inputs"], key=args.get("key") or None,
                       dedup=bool(args.get("dedup")))
        rows = res["rows"]
        write_any(args["output"], rows,
                     columns=res["columns"] + (["__source__"] if "__source__" in _cols(rows) else []))
        return _out("合并完成：%d 个文件 → %d 行，已写出 %s" % (len(args["inputs"]), len(rows), args["output"]))

    if name == "excel_clean":
        rows = ME.read_any(args["input"])
        out = CE.clean(rows, trim=True, drop_blanks=True,
                       dedup=bool(args.get("dedup")),
                       dedup_keys=[k.strip() for k in args["key"].split(",")] if args.get("key") else None,
                       fill=args.get("fill"))
        write_any(args["output"], out)
        return _out("清洗完成：%d 行 → %d 行（减少 %d 行），已写出 %s" % (len(rows), len(out), len(rows) - len(out), args["output"]))

    if name == "excel_pivot":
        rows = ME.read_any(args["input"])
        raw_aggs = args["agg"]
        if isinstance(raw_aggs, str):
            raw_aggs = [raw_aggs]
        aggs = []
        for spec in raw_aggs:
            parts = spec.split(":")
            col, func = parts[0].strip(), parts[1].strip()
            aggs.append({"col": col, "func": func, "as": parts[2].strip() if len(parts) > 2 else None})
        res = PE.pivot(rows, group_by=[g.strip() for g in args["by"].split(",")], aggs=aggs)
        write_any(args["output"], res["summary"],
                     columns=res["group_cols"] + res["value_cols"])
        return _out("透视完成：%d 行 → %d 组，已写出 %s" % (res["stats"]["input_rows"], res["stats"]["groups"], args["output"]))

    if name == "excel_split":
        rows = ME.read_any(args["input"])
        outdir = args["outdir"]
        os.makedirs(outdir, exist_ok=True)
        ext = "xlsx" if args.get("xlsx") else "csv"
        buckets = {}
        for r in rows:
            k = _safe(str(r.get(args["by"], "")).strip()) or "未分类"
            buckets.setdefault(k, []).append(r)
        total = 0
        for k, grp in sorted(buckets.items()):
            path = os.path.join(outdir, "%s.%s" % (k, ext))
            write_any(path, grp)
            total += 1
        return _out("拆分完成：%d 行 → %d 个文件，目录 %s" % (len(rows), total, outdir))

    if name == "excel_recon":
        a = ME.read_any(args["a"])
        b = ME.read_any(args["b"])
        key = args["key"]
        outdir = args["outdir"]
        os.makedirs(outdir, exist_ok=True)
        ext = "xlsx" if args.get("xlsx") else "csv"
        da = {str(r.get(key, "")).strip(): r for r in a}
        db = {str(r.get(key, "")).strip(): r for r in b}
        only_a = [da[k] for k in da if k and k not in db]
        only_b = [db[k] for k in db if k and k not in da]
        common = [k for k in da if k and k in db]
        diffs = []
        for k in common:
            ra, rb = da[k], db[k]
            for c in [c for c in _cols(a) if c in _cols(b) and c != key]:
                va, vb = str(ra.get(c, "")).strip(), str(rb.get(c, "")).strip()
                na, nb = _num(va), _num(vb)
                same = (abs(na - nb) < 1e-9) if (na is not None and nb is not None) else (va == vb)
                if not same:
                    diffs.append({"主键": k, "列": c, "表A": va, "表B": vb})
        write_any(os.path.join(outdir, "1_仅在表A.%s" % ext), only_a)
        write_any(os.path.join(outdir, "2_仅在表B.%s" % ext), only_b)
        write_any(os.path.join(outdir, "3_数值不一致.%s" % ext), diffs,
                     columns=["主键", "列", "表A", "表B"])
        return _out("对账完成（主键 %s）：仅A %d / 仅B %d / 不一致 %d，结果在 %s" % (key, len(only_a), len(only_b), len(diffs), outdir))

    raise ValueError("未知工具：%s" % name)


def _cols(rows):
    out = []
    for r in rows:
        for k in r.keys():
            if k not in out:
                out.append(k)
    return out


def _num(s):
    if s is None:
        return None
    t = str(s).strip().replace(",", "").replace("¥", "").replace("￥", "").replace("%", "")
    if t in ("", "-", "--", "N/A", "NA", "null", "None"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _safe(name):
    import re
    t = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip().rstrip(".")
    return t or "未分类"


# ---------------------------------------------------------------------------
# stdio JSON-RPC 传输（Content-Length 分帧）
# ---------------------------------------------------------------------------
def read_message(stream):
    """从 stream 读取一条 JSON-RPC 消息（按 Content-Length 分帧）。返回 dict 或 None。"""
    headers = {}
    while True:
        line = stream.readline()
        if not line:
            return None
        line = line.decode("utf-8") if isinstance(line, bytes) else line
        if line in ("\r\n", "\n", ""):
            break
        if ":" in line:
            k, v = line.split(":", 1)
            headers[k.strip().lower()] = v.strip()
    if "content-length" not in headers:
        return None
    size = int(headers["content-length"])
    body = stream.read(size)
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


def write_message(stream, msg):
    data = json.dumps(msg, ensure_ascii=False).encode("utf-8")
    stream.write(b"Content-Length: %d\r\n\r\n" % len(data))
    stream.write(data)
    stream.flush()


def main():
    # 关键：直接读底层二进制缓冲，勿用 TextIOWrapper。
    # TextIOWrapper 的解码前瞻缓冲会在多次 readline/read 间错位帧（吞掉后续 header 字节），
    # 导致连续多次 tools/call 时服务端卡死。BufferedReader 的 readline 与 read(size) 共用同一缓冲，安全。
    istream = sys.stdin.buffer
    ostream = sys.stdout.buffer  # 直接写底层缓冲（勿二次包裹，否则 flush 不触达 OS 管道）
    while True:
        try:
            req = read_message(istream)
        except Exception:
            break
        if req is None:
            break
        mid = req.get("id")
        method = req.get("method", "")
        params = req.get("params", {}) or {}

        if method == "initialize":
            write_message(ostream, {
                "jsonrpc": "2.0", "id": mid,
                "result": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": "excel-deliver-mcp", "version": "1.0.0"},
                },
            })
        elif method == "notifications/initialized":
            continue
        elif method == "ping":
            write_message(ostream, {"jsonrpc": "2.0", "id": mid, "result": {}})
        elif method == "tools/list":
            write_message(ostream, {
                "jsonrpc": "2.0", "id": mid,
                "result": {"tools": TOOLS},
            })
        elif method == "tools/call":
            try:
                result = run_tool(params.get("name", ""), params.get("arguments", {}) or {})
                write_message(ostream, {
                    "jsonrpc": "2.0", "id": mid,
                    "result": {"content": result["content"], "isError": False},
                })
            except Exception as e:
                write_message(ostream, {
                    "jsonrpc": "2.0", "id": mid,
                    "result": {"content": [{"type": "text", "text": "错误：%s" % e}], "isError": True},
                })
        else:
            if mid is not None:
                write_message(ostream, {
                    "jsonrpc": "2.0", "id": mid,
                    "error": {"code": -32601, "message": "方法未实现：%s" % method},
                })


if __name__ == "__main__":
    main()
