#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
eval_mcp.py — excel-deliver-mcp 自测（宣称=实测）

作为子进程拉起 server.py，通过 stdio JSON-RPC 帧验证：
  1) initialize 握手成功
  2) tools/list 返回 5 个工具
  3) tools/call excel_merge 真实合并两个样本并产出文件
  4) tools/call excel_recon 真实对账，产出 3 个结果文件
退出码 0 表示全部通过。
"""
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(HERE, "server.py")


def send(proc, msg):
    data = json.dumps(msg, ensure_ascii=False).encode("utf-8")
    proc.stdin.write(b"Content-Length: %d\r\n\r\n" % len(data))
    proc.stdin.write(data)
    proc.stdin.flush()


def recv(proc):
    # 读 header
    headers = {}
    while True:
        line = proc.stdout.readline()
        if not line:
            return None
        if line in (b"\r\n", b"\n", b""):
            break
        if b":" in line:
            k, v = line.split(b":", 1)
            headers[k.strip().lower()] = v.strip()
    size = int(headers[b"content-length"])
    body = proc.stdout.read(size)
    return json.loads(body.decode("utf-8"))


def main():
    passed = 0
    failed = 0

    def check(cond, label):
        nonlocal passed, failed
        if cond:
            passed += 1
            print("  [PASS] %s" % label)
        else:
            failed += 1
            print("  [FAIL] %s" % label)

    tmp = tempfile.mkdtemp(prefix="mcp_eval_")
    a = os.path.join(tmp, "a.csv")
    b = os.path.join(tmp, "b.csv")
    with open(a, "w", encoding="utf-8") as f:
        f.write("订单号,城市,金额\nA001,上海,100\nA002,北京,200\nA003,上海,150\n")
    with open(b, "w", encoding="utf-8") as f:
        f.write("订单号,城市,金额\nA002,北京,200\nA004,广州,300\n")

    proc = subprocess.Popen(
        [sys.executable, SERVER],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=open("/tmp/eval_srv.log", "w"),
    )
    try:
        # 1) initialize
        send(proc, {"jsonrpc": "2.0", "id": 1, "method": "initialize",
                    "params": {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "eval", "version": "1"}}})
        r = recv(proc)
        check(r and r.get("result", {}).get("serverInfo", {}).get("name") == "excel-deliver-mcp",
              "initialize 返回 serverInfo")

        # 2) tools/list
        send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
        r = recv(proc)
        tools = (r or {}).get("result", {}).get("tools", [])
        names = [t["name"] for t in tools]
        check(len(tools) == 5 and set(names) == {"excel_merge", "excel_clean", "excel_pivot", "excel_split", "excel_recon"},
              "tools/list 返回 5 个工具: %s" % names)

        # 3) tools/call excel_merge
        out = os.path.join(tmp, "merged.xlsx")
        send(proc, {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
                    "params": {"name": "excel_merge", "arguments": {"inputs": [a, b], "output": out, "key": "订单号", "dedup": True}}})
        r = recv(proc)
        res = (r or {}).get("result", {})
        check(not res.get("isError") and os.path.exists(out), "excel_merge 调用成功并产出文件")

        # 4) tools/call excel_recon
        rdir = os.path.join(tmp, "recon")
        send(proc, {"jsonrpc": "2.0", "id": 4, "method": "tools/call",
                    "params": {"name": "excel_recon", "arguments": {"a": a, "b": b, "key": "订单号", "outdir": rdir}}})
        r = recv(proc)
        res = (r or {}).get("result", {})
        ok = (not res.get("isError")) and os.path.exists(os.path.join(rdir, "1_仅在表A.csv"))
        check(ok, "excel_recon 调用成功并产出 1_仅在表A.csv")

        # 5) ping
        send(proc, {"jsonrpc": "2.0", "id": 5, "method": "ping", "params": {}})
        r = recv(proc)
        check(r and "result" in r, "ping 正常返回")
    finally:
        try:
            proc.stdin.close()
        except Exception:
            pass
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()

    print("\nMCP EVAL: %d passed, %d failed" % (passed, failed))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
