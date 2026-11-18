# -*- coding: utf-8 -*-
"""Turn hand-typed heading numbers into a Word/WPS multilevel auto-number.

The current document stores section numbers as literal text, so deleting a
chapter leaves the following numbers stale until someone edits them by hand.
This module replaces that text with a ``w:numPr`` reference into a
``numbering.xml`` multilevel definition: the editor then owns the numbers and
renumbers automatically when chapters are added or removed.

The numbering style is inferred from the source text, so 一、/（一）/1.
and 1/1.1 documents both keep their original look.
"""
from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CHINESE_DIGITS = '一二三四五六七八九十百'

# 匹配完整编号前缀（序号 + 分隔符/括号），删除时要整段去掉。
# group(1) 只取序号本身，用于推断 lvlText。
PREFIX_PATTERNS = (
    ('multi', re.compile(r'^(\d{1,3}(?:[.．]\d{1,3})+)')),
    ('chinese_dun', re.compile(r'^([' + CHINESE_DIGITS + r']+)\s*、')),
    ('chinese_paren', re.compile(r'^[（(]([' + CHINESE_DIGITS + r']+)[）)]')),
    ('arabic_dot', re.compile(r'^(\d{1,3})\s*[.．](?!\d)')),
    ('arabic_space', re.compile(r'^(\d{1,3})\s+')),
    ('arabic_paren', re.compile(r'^[（(](\d{1,3})[）)]')),
)

# 每种编号形态对应的 numFmt 与后缀处理。
# suffix: 'nothing' 表示 lvlText 自带分隔符（如"、"），'space' 表示编号后补空格。
FORMAT_TABLE = {
    'chinese_dun': ('chineseCounting', 'nothing'),
    'chinese_paren': ('chineseCounting', 'nothing'),
    'arabic_dot': ('decimal', 'space'),
    'arabic_space': ('decimal', 'space'),
    'arabic_paren': ('decimal', 'nothing'),
}


def _paragraph_text(paragraph):
    return ''.join(node.text or '' for node in paragraph._p.iter(qn('w:t')))


def parse_prefix(text):
    """返回 (kind, 序号文本, 完整编号前缀) 或 None。"""
    stripped = (text or '').strip()
    for kind, pattern in PREFIX_PATTERNS:
        match = pattern.match(stripped)
        if match:
            return kind, match.group(1), match.group(0)
    return None


def strip_prefix(paragraph, prefix):
    """从段落里删掉开头的编号文字，返回是否删除成功。"""
    for run in paragraph.runs:
        text = run.text or ''
        if not text:
            continue
        lead = len(text) - len(text.lstrip())
        body = text[lead:]
        if not body.startswith(prefix):
            return False
        # 编号后面的空白也一并去掉，交给自动编号的 suff 处理
        rest = body[len(prefix):].lstrip()
        run.text = text[:lead] + rest
        return True
    return False


def _lvl_text(kind, level, digits):
    """按编号形态拼出 lvlText，%N 是层级占位符。"""
    if kind == 'multi':
        span = len(re.split(r'[.．]', digits))
        # 多级编号要引用"本级及以上"的层级：三级 3.1 是 %2.%3 而不是 %1.%2，
        # 否则父级计数会落到空着的一级上，各节都显示成 1.x。
        start = max(1, level - span + 1)
        parts = [f'%{index}' for index in range(start, start + span)]
        return '.'.join(parts)
    if kind == 'chinese_dun':
        return f'%{level}、'
    if kind == 'chinese_paren':
        return f'（%{level}）'
    if kind == 'arabic_dot':
        return f'%{level}.'
    if kind == 'arabic_paren':
        return f'（%{level}）'
    return f'%{level}'


def _next_abstract_id(numbering):
    used = [int(node.get(qn('w:abstractNumId')))
            for node in numbering.findall(qn('w:abstractNum'))
            if (node.get(qn('w:abstractNumId')) or '').isdigit()]
    return (max(used) + 1) if used else 0


def _next_num_id(numbering):
    used = [int(node.get(qn('w:numId')))
            for node in numbering.findall(qn('w:num'))
            if (node.get(qn('w:numId')) or '').isdigit()]
    return (max(used) + 1) if used else 1


def build_definition(numbering, level_formats, max_level):
    """写入一个多级 abstractNum 和一个 num 实例，返回 numId。"""
    abstract_id = _next_abstract_id(numbering)
    num_id = _next_num_id(numbering)

    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(abstract_id))
    multi = OxmlElement('w:multiLevelType')
    multi.set(qn('w:val'), 'multilevel')
    abstract.append(multi)

    for level in range(1, max_level + 1):
        kind = level_formats.get(level)
        if kind is None:
            num_fmt, suffix = 'decimal', 'space'
            lvl_text = f'%{level}'
        elif kind == 'multi':
            num_fmt, suffix = 'decimal', 'nothing'
            lvl_text = level_formats[('text', level)]
        else:
            num_fmt, suffix = FORMAT_TABLE.get(kind, ('decimal', 'space'))
            lvl_text = _lvl_text(kind, level, '')

        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(level - 1))
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)
        fmt = OxmlElement('w:numFmt')
        fmt.set(qn('w:val'), num_fmt)
        lvl.append(fmt)
        text = OxmlElement('w:lvlText')
        text.set(qn('w:val'), lvl_text)
        lvl.append(text)
        justification = OxmlElement('w:lvlJc')
        justification.set(qn('w:val'), 'left')
        lvl.append(justification)
        suff = OxmlElement('w:suff')
        suff.set(qn('w:val'), suffix)
        lvl.append(suff)
        ppr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), '0')
        ind.set(qn('w:hanging'), '0')
        ppr.append(ind)
        lvl.append(ppr)
        abstract.append(lvl)

    existing = numbering.findall(qn('w:abstractNum'))
    if existing:
        existing[-1].addnext(abstract)
    else:
        numbering.insert(0, abstract)

    instance = OxmlElement('w:num')
    instance.set(qn('w:numId'), str(num_id))
    ref = OxmlElement('w:abstractNumId')
    ref.set(qn('w:val'), str(abstract_id))
    instance.append(ref)
    numbering.append(instance)
    return num_id


def attach_numbering(paragraph, num_id, level):
    """给段落挂上 numPr，指向刚建好的多级列表。"""
    ppr = paragraph._p.get_or_add_pPr()
    for old in ppr.findall(qn('w:numPr')):
        ppr.remove(old)
    numpr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level - 1))
    numpr.append(ilvl)
    num_id_node = OxmlElement('w:numId')
    num_id_node.set(qn('w:val'), str(num_id))
    numpr.append(num_id_node)
    ppr.append(numpr)


def apply_heading_auto_numbering(doc, headings, log=None, max_level=4):
    """把标题的文字编号换成分级自动编号。

    headings: 按文档顺序的 [(paragraph, level)]，level 从 1 开始。
    返回实际改写的标题数；无法处理时返回 0 且不动文档。
    """
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        if log:
            log("  > 警告：文档缺少 numbering 部件，跳过自动编号。")
        return 0

    entries = []
    level_formats = {}
    for paragraph, level in headings:
        if level < 1 or level > max_level:
            continue
        parsed = parse_prefix(_paragraph_text(paragraph))
        if parsed is None:
            continue
        kind, digits, full = parsed
        entries.append((paragraph, level, full))
        # 每级只取第一个样本决定格式，避免混排把定义弄乱
        if level not in level_formats:
            level_formats[level] = kind
            if kind == 'multi':
                level_formats[('text', level)] = _lvl_text(kind, level, digits)

    if not entries:
        if log:
            log("  > 标题没有可识别的编号，未生成自动编号。")
        return 0

    num_id = build_definition(numbering, level_formats, max_level)
    changed = 0
    for paragraph, level, full_prefix in entries:
        if not strip_prefix(paragraph, full_prefix):
            continue
        attach_numbering(paragraph, num_id, level)
        changed += 1
    if log:
        log(f"  > 已将 {changed} 个标题改为 Word/WPS 自动编号（多级列表）。")
    return changed
