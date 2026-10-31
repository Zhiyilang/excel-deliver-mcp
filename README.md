# excel-deliver-mcp

一个**零依赖**（仅 Python 标准库，无需 `openpyxl`/`pandas`）的 MCP 服务器，把 Excel/CSV 的常用数据交付动作封装成 5 个 MCP 工具，让任意支持 MCP 的 LLM Agent（Claude Desktop、Cursor、Cline 等）直接调用，自动完成**合并、清洗、透视、拆分、对账**。

> 设计原则：纯 `stdlib`，`csv` 模块自实现 xlsx 读写（不依赖任何第三方包），因此它能在任何装了 Python 3.9+ 的机器上即跑即⽤，特别适合塞进 Agent 的工具箱。

## 工具清单（5 个）

| 工具 | 作用 | 关键参数 |
|---|---|---|
| `excel_merge` | 多表合并 + 可选按主键去重 | `inputs`(列表), `output`, `key`, `dedup` |
| `excel_clean` | 去首尾空格、删空行、按列去重、空值填充 | `input`, `output`, `key`, `dedup`, `fill` |
| `excel_pivot` | 按维度列分组聚合（sum/count/avg/max/min） | `input`, `output`, `by`, `agg`(数组: `列:func[:别名]`) |
| `excel_split` | 按某列值拆成多个文件（文件名即值） | `input`, `by`, `outdir`, `xlsx` |
| `excel_recon` | 双表按主键逐行比对，输出仅在A/仅在B/不一致 | `a`, `b`, `key`, `outdir` |

全部返回中文进度文本，出错时 `isError=True` 并附带错误信息。

## 安装与运行

```bash
# 无需 pip install，仅依赖 Python 3.9+
python3 server.py          # 以 stdio 方式启动 MCP 服务器
python3 eval_mcp.py        # 自测：initialize / tools/list / 调用 5 个工具，退出码 0 表示全过
```

`eval_mcp.py` 会拉起子进程、通过 JSON-RPC 帧完整验证 5 个工具真实产出文件，**宣称即实测**（当前：5 passed, 0 failed）。

## 接入 Claude Desktop / Cursor

`claude_desktop_config.json` 添加：

```json
{
  "mcpServers": {
    "excel-deliver": {
      "command": "python3",
      "args": ["/绝对路径/excel_deliver_mcp/server.py"]
    }
  }
}
```

## 典型场景

- 把一个月 30 个分店的 csv 合并成一张总表并按订单号去重；
- 对账：财务表 vs 业务表，一键揪出「只在 A 表」「只在 B 表」「数值不一致」三张清单；
- Agent 自动把大表按城市/部门拆给不同同事。

## 许可

MIT — 见 [LICENSE](LICENSE)。

---

## English

A zero-dependency (Python stdlib only, no `openpyxl`/`pandas`) MCP server exposing 5 Excel/CSV delivery tools — merge, clean, pivot, split, recon — to any MCP-capable LLM agent. Self-contained xlsx read/write implemented over the `csv` module. Run `python3 server.py` for stdio MCP, `python3 eval_mcp.py` for a full self-test (5/0). MIT licensed.
