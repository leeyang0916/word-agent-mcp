# word-agent-mcp

让 AI agent 以**修订模式（Track Changes）**编辑 Word 文档的 MCP server —— 律师版 "Claude Code for Word" 的文档操作层。

## 核心理念

程序员审代码用 diff，律师审文书用**修订**。本项目让 agent 的每一处修改都以标准 Word 修订（红线删除 + 下划线插入）写入 `.docx`，律师在 Word「审阅」界面逐条接受/拒绝——agent 永远不能未经审阅直接改动文书。

修订直接写入 OOXML（`w:ins`/`w:del`），**不需要安装 Word、跨平台**（Mac 开发、Windows 交付零移植）。

## 架构

```
src/word_agent_mcp/
├── server.py          # MCP 工具层（read_document / replace_text / insert_paragraph / ...）
├── revisions.py       # OOXML 修订标记核心（w:ins / w:del 生成、run 拆分、修订扫描）
└── backend/
    ├── base.py        # WordBackend 抽象接口（按 Windows COM 能力超集设计）
    └── ooxml.py       # 文件后端：python-docx + lxml，跨平台
                       # （规划）com.py：Windows COM 后端，实时操控打开中的 Word
```

工具层只依赖 `WordBackend` 接口。阶段 2 在 Windows 上补一个 COM 后端即可获得"实时看着 agent 改文档"的体验，工具层零改动。

## 快速开始

```bash
uv sync                          # 安装依赖（自动准备 Python 3.11+）
uv run pytest                    # 跑测试
uv run python examples/demo.py   # 生成演示合同并修订，用 Word 打开看红线
```

## 接入 Claude Code

```bash
claude mcp add word-agent -- uv --directory /path/to/word-agent-mcp run word-agent-mcp
```

或在项目 `.mcp.json` 里：

```json
{
  "mcpServers": {
    "word-agent": {
      "command": "uv",
      "args": ["--directory", "/path/to/word-agent-mcp", "run", "word-agent-mcp"]
    }
  }
}
```

然后直接对话：「读一下 ~/Desktop/租赁合同.docx，把租金改成 9000 并加一条逾期解约条款」。

## MCP 工具

| 工具 | 说明 |
|---|---|
| `read_document` | 读取全部段落（index / style / 有效文本） |
| `replace_text` | 修订式替换；不唯一时要求更长上下文或 `replace_all` |
| `insert_paragraph` | 在指定段后修订式插入新段落 |
| `delete_paragraph` | 修订式删除段落 |
| `list_revisions` | 列出文档全部修订（自检 / 汇报用） |

安全设计：每次编辑前自动留 `.bak` 备份；所有修改必须过 Word 审阅这一关。

## 已知限制（MVP）

- 只处理正文段落，表格/页眉页脚/脚注暂不支持
- 匹配区间内含制表符、图片等复杂 run 结构时会明确报错（不会静默破坏文档）
- 批注（comments）待加入
- Windows 上 Word 正在打开的文件会被锁定：请让 agent 改副本，或关闭文档后编辑
