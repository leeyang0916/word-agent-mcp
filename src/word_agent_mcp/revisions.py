"""OOXML 修订标记（Track Changes）底层操作。

Word 的修订在 OOXML 里就是普通 XML 标记：
- 插入:  <w:ins w:id=".." w:author=".." w:date=".."><w:r><w:t>新文字</w:t></w:r></w:ins>
- 删除:  <w:del w:id=".." w:author=".." w:date=".."><w:r><w:delText>旧文字</w:delText></w:r></w:del>

不需要 Word 进程参与即可写入，用 Word 打开后表现为标准红线修订，
可逐条接受/拒绝。这是本项目跨平台（Mac 开发 / Windows 交付）的基础。
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field

from docx.oxml.ns import qn
from lxml import etree

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class RevisionError(Exception):
    """修订操作失败（定位不到文本、结构不支持等）。"""


@dataclass
class RevisionContext:
    """一次编辑会话的修订元数据：作者、时间戳、递增的修订 id。"""

    author: str
    date: str  # ISO-8601, e.g. 2026-07-08T10:00:00Z
    _next_id: int = field(default=1)

    @classmethod
    def for_document(cls, body: etree._Element, author: str, date: str) -> "RevisionContext":
        max_id = 0
        for el in body.iter(qn("w:ins"), qn("w:del")):
            raw = el.get(qn("w:id"))
            if raw is not None:
                try:
                    max_id = max(max_id, int(raw))
                except ValueError:
                    pass
        return cls(author=author, date=date, _next_id=max_id + 1)

    def next_id(self) -> str:
        rid = self._next_id
        self._next_id += 1
        return str(rid)

    def make_marker(self, tag: str) -> etree._Element:
        """构造 <w:ins> 或 <w:del> 元素（tag 形如 'w:ins'）。"""
        el = etree.Element(qn(tag))
        el.set(qn("w:id"), self.next_id())
        el.set(qn("w:author"), self.author)
        el.set(qn("w:date"), self.date)
        return el


# ---------------------------------------------------------------------------
# 文本定位：把段落里的 run 拉平成连续字符流
# ---------------------------------------------------------------------------


def _visible_runs(p: etree._Element) -> list[etree._Element]:
    """段落中参与"当前有效文本"的 run：含 w:t 且不在 w:del 之内。

    已被标记删除的文字（w:delText）不算有效文本；
    已有 w:ins 里的 run 算有效文本（那是已插入待接受的内容）。
    """
    runs = []
    for r in p.iter(qn("w:r")):
        if r.find(qn("w:t")) is None:
            continue
        anc = r.getparent()
        in_del = False
        while anc is not None and anc is not p:
            if anc.tag == qn("w:del"):
                in_del = True
                break
            anc = anc.getparent()
        if not in_del:
            runs.append(r)
    return runs


def _run_text(r: etree._Element) -> str:
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def paragraph_effective_text(p: etree._Element) -> str:
    """段落当前有效文本（接受所有修订后的样子）。"""
    return "".join(_run_text(r) for r in _visible_runs(p))


def _set_text(t: etree._Element, value: str) -> None:
    t.text = value
    t.set(XML_SPACE, "preserve")


def _split_run(r: etree._Element, offset: int) -> etree._Element:
    """在 offset 处把 run 一分为二，返回后半个（已插入原 run 之后）。

    仅支持单一 w:t 的 run —— python-docx 与 Word 生成的常规文档均满足；
    含 tab/图片等复杂结构的 run 直接报错，由上层给出清晰提示。
    """
    ts = r.findall(qn("w:t"))
    if len(ts) != 1:
        raise RevisionError("目标文本所在的 run 结构复杂（含制表符/特殊元素），暂不支持在其内部拆分")
    full = ts[0].text or ""
    right = deepcopy(r)
    _set_text(ts[0], full[:offset])
    _set_text(right.findall(qn("w:t"))[0], full[offset:])
    r.addnext(right)
    return right


def _wrap_in_del(r: etree._Element, ctx: RevisionContext) -> etree._Element:
    """把 run 就地包进 <w:del>，w:t 改写为 w:delText。"""
    parent = r.getparent()
    idx = list(parent).index(r)
    del_el = ctx.make_marker("w:del")
    parent.insert(idx, del_el)
    del_el.append(r)
    for t in r.findall(qn("w:t")):
        t.tag = qn("w:delText")
        t.set(XML_SPACE, "preserve")
    return del_el


def count_occurrences(p: etree._Element, needle: str) -> int:
    return paragraph_effective_text(p).count(needle)


def tracked_replace_in_paragraph(
    p: etree._Element, old: str, new: str, ctx: RevisionContext, *, all_occurrences: bool = False
) -> int:
    """在段落内以修订方式把 old 替换为 new，返回替换次数。

    先一次性定位全部出现位置，再从后往前逐处替换：前面的偏移不受影响，
    且新插入的文本（可能包含 old，如"甲方"→"甲方（即买方）"）不会被再次匹配。
    """
    joined = paragraph_effective_text(p)
    starts: list[int] = []
    i = joined.find(old)
    while i >= 0:
        starts.append(i)
        if not all_occurrences:
            break
        i = joined.find(old, i + len(old))
    if not starts:
        raise RevisionError(f"段落中未找到文本: {old!r}")
    for s in reversed(starts):
        _replace_range(p, s, s + len(old), new, ctx)
    return len(starts)


def _replace_range(
    p: etree._Element, start: int, end: int, new: str, ctx: RevisionContext
) -> None:
    """把段落有效文本 [start, end) 区间标记为删除，并在原位插入 new。"""
    runs = _visible_runs(p)
    spans: list[tuple[etree._Element, int, int]] = []
    pos = 0
    for r in runs:
        text = _run_text(r)
        spans.append((r, pos, pos + len(text)))
        pos += len(text)

    # 在匹配边界处拆分 run，使匹配区间恰好由整数个 run 组成。
    # 起点拆分后，原区间的剩余部分在新产生的右半 run 里，终点拆分要以它为基准。
    for r, r_start, r_end in spans:
        target, base = r, r_start
        if r_start < start < r_end:
            target = _split_run(r, start - r_start)
            base = start
        if base < end < r_end:
            _split_run(target, end - base)

    # 拆分后重新扫描，收集完全落在匹配区间内的 run
    runs = _visible_runs(p)
    matched: list[etree._Element] = []
    pos = 0
    for r in runs:
        length = len(_run_text(r))
        if pos >= start and pos + length <= end and length > 0:
            matched.append(r)
        pos += length
    if not matched:
        raise RevisionError(f"内部错误：拆分后未能定位到区间 [{start}, {end})")

    rpr_template = matched[0].find(qn("w:rPr"))
    last_del = None
    for r in matched:
        last_del = _wrap_in_del(r, ctx)

    if new:
        ins_el = ctx.make_marker("w:ins")
        new_r = etree.SubElement(ins_el, qn("w:r"))
        if rpr_template is not None:
            new_r.append(deepcopy(rpr_template))
        t = etree.SubElement(new_r, qn("w:t"))
        _set_text(t, new)
        last_del.addnext(ins_el)


def tracked_delete_paragraph(p: etree._Element, ctx: RevisionContext) -> None:
    """以修订方式删除整个段落（内容 + 段落标记）。"""
    for r in _visible_runs(p):
        _wrap_in_del(r, ctx)
    _mark_paragraph_mark(p, "w:del", ctx)


def tracked_new_paragraph(
    p_after: etree._Element, text: str, ctx: RevisionContext, style: str | None = None
) -> etree._Element:
    """在 p_after 之后以修订方式插入新段落，返回新段落元素。"""
    new_p = etree.Element(qn("w:p"), nsmap=p_after.nsmap)
    if style:
        ppr = etree.SubElement(new_p, qn("w:pPr"))
        pstyle = etree.SubElement(ppr, qn("w:pStyle"))
        pstyle.set(qn("w:val"), style)
    _mark_paragraph_mark(new_p, "w:ins", ctx)
    ins_el = ctx.make_marker("w:ins")
    new_p.append(ins_el)
    r = etree.SubElement(ins_el, qn("w:r"))
    t = etree.SubElement(r, qn("w:t"))
    _set_text(t, text)
    p_after.addnext(new_p)
    return new_p


def _mark_paragraph_mark(p: etree._Element, tag: str, ctx: RevisionContext) -> None:
    """在 pPr/rPr 里标记段落标记本身被插入/删除（Word 合并段落的表示方式）。"""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(ppr, qn("w:rPr"))
    rpr.append(ctx.make_marker(tag))


def list_revisions(body: etree._Element) -> list[dict]:
    """扫描文档中的全部修订，返回结构化清单。"""
    out = []
    for el in body.iter(qn("w:ins"), qn("w:del")):
        kind = "插入" if el.tag == qn("w:ins") else "删除"
        parent = el.getparent()
        if parent is not None and parent.tag == qn("w:rPr"):
            kind += "段落标记"
            text = ""
        else:
            tags = (qn("w:t"),) if el.tag == qn("w:ins") else (qn("w:delText"),)
            text = "".join(t.text or "" for tag in tags for t in el.iter(tag))
        out.append(
            {
                "id": el.get(qn("w:id")),
                "type": kind,
                "author": el.get(qn("w:author")),
                "date": el.get(qn("w:date")),
                "text": text,
            }
        )
    return out
