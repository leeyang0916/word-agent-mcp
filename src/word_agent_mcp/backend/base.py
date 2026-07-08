"""WordBackend 抽象接口。

接口按能力超集设计（以 Windows COM 的能力为上限）：
- OoxmlBackend（本文件夹 ooxml.py）：直接读写 .docx 文件，跨平台，Mac 上开发测试。
- ComBackend（未来，Windows）：通过 Word COM 实时操控打开的文档，实现 live 能力。

工具层只依赖本接口，切换后端不改工具代码。
"""

from __future__ import annotations

from abc import ABC, abstractmethod


class CapabilityNotSupported(Exception):
    """当前后端不支持该能力（例如文件后端不支持实时预览）。"""


class WordBackend(ABC):
    """一个后端实例对应一份打开的文档。"""

    #: 是否实时操控运行中的 Word（COM 后端为 True）
    is_live: bool = False

    # ---- 读取 ----

    @abstractmethod
    def get_paragraphs(self) -> list[dict]:
        """返回 [{index, style, text}]，text 为接受全部修订后的有效文本。"""

    @abstractmethod
    def list_revisions(self) -> list[dict]:
        """返回文档中的全部修订 [{id, type, author, date, text}]。"""

    # ---- 修订式编辑 ----

    @abstractmethod
    def replace_text(self, old: str, new: str, *, replace_all: bool = False) -> int:
        """以修订方式替换文本，返回替换次数。old 不唯一且未指定 replace_all 时应报错。"""

    @abstractmethod
    def insert_paragraph_after(self, index: int, text: str, *, style: str | None = None) -> None:
        """在第 index 段之后以修订方式插入新段落。"""

    @abstractmethod
    def delete_paragraph(self, index: int) -> None:
        """以修订方式删除第 index 段。"""

    # ---- 持久化 ----

    @abstractmethod
    def save(self, path: str | None = None) -> str:
        """保存（可另存为），返回实际写入路径。"""

    # ---- live 能力（COM 后端实现；文件后端一律拒绝） ----

    def accept_revision(self, revision_id: str) -> None:
        raise CapabilityNotSupported("接受/拒绝修订请在 Word 审阅界面完成（当前为文件后端）")

    def reject_revision(self, revision_id: str) -> None:
        raise CapabilityNotSupported("接受/拒绝修订请在 Word 审阅界面完成（当前为文件后端）")

    def show_in_word(self) -> None:
        raise CapabilityNotSupported("实时预览需要 Windows COM 后端（规划中）")
