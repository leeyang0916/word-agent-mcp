"""MCP server：把 Word 文档修订能力暴露给 AI agent（Claude Code / Agent SDK 等）。

设计原则：
- 每个工具调用独立完成 打开→编辑→保存，无跨调用状态，agent 崩溃不会留下半成品。
- 所有编辑一律以修订（Track Changes）写入，最终由律师在 Word 审阅界面逐条接受/拒绝。
- 首次编辑自动在原文件旁留 .bak 备份。
"""

from __future__ import annotations

import json
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from .backend import OoxmlBackend
from .revisions import RevisionError

mcp = FastMCP("word-agent")

DEFAULT_AUTHOR = "AI 律师助手"


def _json(data) -> str:
    return json.dumps(data, ensure_ascii=False, indent=1)


def _abs(path: str) -> str:
    """要求绝对路径：MCP server 进程的工作目录与用户会话目录不同，
    相对路径会解析到错误位置。报错信息引导 agent 自行纠正。"""
    p = Path(path).expanduser()
    if not p.is_absolute():
        raise RevisionError(
            f"必须传文档的绝对路径（收到 {path!r}）。"
            f"请结合用户会话的工作目录拼出完整路径后重试"
        )
    return str(p)


@mcp.tool()
def read_document(path: str) -> str:
    """读取 Word 文档（.docx）的全部段落。path 必须是绝对路径。

    返回 JSON 数组，每段含 index（段落序号，编辑工具以此定位）、style（样式名）、
    text（接受全部现有修订后的有效文本）。编辑前必须先读取文档。
    """
    try:
        doc = OoxmlBackend(_abs(path), backup=False)
        return _json(doc.get_paragraphs())
    except (RevisionError, OSError) as e:
        return f"错误: {e}"


@mcp.tool()
def replace_text(
    path: str,
    old_text: str,
    new_text: str,
    replace_all: bool = False,
    author: str = DEFAULT_AUTHOR,
) -> str:
    """以修订模式（Track Changes）替换文档中的文本。

    修改不会直接生效，而是写成 Word 修订（删除线 + 下划线），由用户在 Word
    审阅界面逐条接受/拒绝。old_text 在全文只出现一次时才会替换；出现多次时
    需提供更长的唯一上下文，或明确 replace_all=True 全部替换。
    new_text 传空字符串表示纯删除。
    """
    try:
        doc = OoxmlBackend(_abs(path), author=author)
        n = doc.replace_text(old_text, new_text, replace_all=replace_all)
        doc.save()
        return f"已以修订方式替换 {n} 处。用户可在 Word 审阅界面接受/拒绝。原文件备份为 {path}.bak"
    except (RevisionError, OSError) as e:
        return f"错误: {e}"


@mcp.tool()
def insert_paragraph(
    path: str,
    after_index: int,
    text: str,
    style: str | None = None,
    author: str = DEFAULT_AUTHOR,
) -> str:
    """在第 after_index 段（从 0 计，见 read_document）之后以修订方式插入新段落。

    style 可选，填文档已有的段落样式名（如 "Heading 2"）。
    """
    try:
        doc = OoxmlBackend(_abs(path), author=author)
        doc.insert_paragraph_after(after_index, text, style=style)
        doc.save()
        return f"已在第 {after_index} 段后以修订方式插入新段落。"
    except (RevisionError, OSError) as e:
        return f"错误: {e}"


@mcp.tool()
def delete_paragraph(path: str, index: int, author: str = DEFAULT_AUTHOR) -> str:
    """以修订方式删除第 index 段（从 0 计，见 read_document）。"""
    try:
        doc = OoxmlBackend(_abs(path), author=author)
        doc.delete_paragraph(index)
        doc.save()
        return f"已以修订方式删除第 {index} 段。"
    except (RevisionError, OSError) as e:
        return f"错误: {e}"


@mcp.tool()
def list_revisions(path: str) -> str:
    """列出文档中现存的全部修订（含此前人工修订），返回 JSON。

    用于编辑后自检，或向用户汇报"本次共产生了哪些修订"。
    """
    try:
        doc = OoxmlBackend(_abs(path), backup=False)
        return _json(doc.list_revisions())
    except (RevisionError, OSError) as e:
        return f"错误: {e}"


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
