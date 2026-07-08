# CLAUDE.md

## 项目定位

律师行业的 "Claude Code for Word"：MCP server，让 AI agent 以**修订模式（Track Changes）**编辑 .docx。核心产品原则：**agent 的任何修改只能以 Word 修订写入，永不直接生效**，由律师在 Word 审阅界面逐条接受/拒绝。

用户（项目所有者）是律师，主力工作平台是 Windows + 桌面版 Word；开发机是 Mac。本仓库是其商业产品的阶段 1（跨平台文件级编辑），阶段 2 将增加 Windows COM 后端实现实时编辑。

## 常用命令

```bash
uv sync                         # 安装依赖（需要 Python 3.11+，uv 自动管理）
uv run pytest -q                # 全量测试（快，<1s，改完必跑）
uv run python examples/demo.py  # 生成演示合同并修订，产出在 examples/output/
uv run word-agent-mcp           # 以 stdio 启动 MCP server
```

## 架构

```
src/word_agent_mcp/
├── server.py        # MCP 工具层（FastMCP）。工具无跨调用状态：每次调用独立完成 打开→编辑→保存
├── revisions.py     # 核心：OOXML 修订标记。w:ins/w:del 生成、run 边界拆分、修订扫描
└── backend/
    ├── base.py      # WordBackend 抽象接口 —— 按 Windows COM 能力超集设计，live 方法默认拒绝
    └── ooxml.py     # 文件后端：python-docx + lxml，跨平台
                     # （规划）com.py：Windows COM 后端，工具层零改动接入
```

- 工具层只依赖 `WordBackend` 接口，禁止在 server.py 里直接操作 XML。
- 修订的本质：`revisions.py` 直接把 `w:ins`/`w:del` 写进 document.xml，不需要 Word 进程。

## 关键约定（改代码前必读）

1. **所有编辑必须走修订**。任何"直接改文本不留修订"的功能都违背产品底线，不要实现。
2. **路径必须绝对**（`server.py` 的 `_abs`）：MCP server 进程 cwd ≠ 用户会话 cwd，相对路径会解析错。报错信息要能引导 agent 自我纠正。
3. **文本替换契约仿照 Claude Code 的 Edit 工具**：old_text 全文不唯一时拒绝执行，要求更长上下文或显式 `replace_all`。
4. **替换实现必须"先定位全部、从后往前改"**（见 `tracked_replace_in_paragraph`）：新文本包含旧文本是律师高频场景（"甲方"→"甲方（出租人）"），逐次找第一处会重复匹配刚插入的内容。
5. **首次编辑自动留 `.bak` 备份**（`OoxmlBackend.__init__`），不要移除。
6. **宁可报错，不可静默破坏文档**：遇到不支持的结构（复杂 run、表格等）明确抛 `RevisionError`。
7. 测试夹具要**刻意把文本拆散到多个 run**（模拟真实 Word 文档的碎片化），只测单 run 的用例没有说服力。

## 测试

- 回归测试在 `tests/test_ooxml_backend.py`，改 `revisions.py` / `backend/` 必须补同场景用例。
- 验证产出时同时检查两层：`list_revisions()` 的语义正确性 + 解压 docx 直查 XML（`w:ins`/`w:del`/`w:delText`）。
- 真实 Word 打开验收（人工）：产出文件在 Word「审阅」里必须显示为标准可接受/拒绝的修订。

## MVP 已知边界

- 只处理正文段落：表格、页眉页脚、脚注、批注（comments）未支持
- 匹配区间内含制表符/图片等复杂 run 会明确报错
- Windows 上 Word 打开中的文件被锁定 → 让 agent 改副本

## 本地联调

已注册到本机 Claude Code / cc-haha（user scope，`~/.claude.json`）：

```bash
claude mcp add --scope user word-agent -- /Users/yang/.local/bin/uv --directory /Users/yang/work/word-agent-mcp run word-agent-mcp
```

注意 GUI 应用不继承 shell PATH，`uv` 必须写绝对路径。测试文档在 `~/word-agent-test/`。

## 路线图

1. ✅ 阶段 1：OOXML 文件后端 + 5 个 MCP 工具（read/replace/insert/delete/list_revisions）
2. 批注工具（add_comment，律师"只评不改"场景）
3. 表格支持（合同里的报价表、当事人信息表）
4. 阶段 2：`backend/com.py` — Windows COM 后端（pywin32），实时编辑打开中的文档；接口已在 `base.py` 预留（`accept_revision`/`reject_revision`/`show_in_word`）
