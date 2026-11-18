# -*- coding: utf-8 -*-
"""标题自动编号：按出现顺序重写序号，保留原文的编号格式。

真实文档里存在两套编号体系，本模块要同时支持：
1. 公文式 ``一、`` / ``（一）`` / ``1.`` / ``（1）``
2. 科技文档式多级 ``1`` / ``3.1`` / ``3.1.2``（下级编号含父级序号）

重排只改序号数字，分隔符、括号、尾随空白一律保持原样。
"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn

from wfp_config import DEFAULT_CONFIG
from wfp_core import WordProcessor


def outline_level(paragraph):
    pPr = paragraph._p.pPr
    if pPr is None:
        return None
    element = pPr.find(qn('w:outlineLvl'))
    if element is None:
        return None
    raw = element.get(qn('w:val'))
    if raw is None or not raw.lstrip('-').isdigit():
        return None
    return int(raw) + 1


class HeadingNumberingTests(unittest.TestCase):
    def processor(self, **options):
        config = dict(DEFAULT_CONFIG)
        config.update(options)
        return WordProcessor(config)

    def format_texts(self, texts, styles=None, **options):
        """排版一组段落，返回输出文档的 (文本, 大纲级别) 列表。"""
        styles = styles or [None] * len(texts)
        options.setdefault('auto_heading_numbering', True)
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source.docx'
            output = Path(directory) / 'output.docx'
            doc = Document()
            for text, style in zip(texts, styles):
                if style is None:
                    doc.add_paragraph(text)
                else:
                    doc.add_paragraph(text, style=style)
            doc.save(source)

            processor = self.processor(**options)
            try:
                with patch.object(processor, '_com_available', return_value=False):
                    processor.format_document(str(source), str(output))
                return [(para.text.strip(), outline_level(para))
                        for para in Document(output).paragraphs]
            finally:
                processor._cleanup_temp_files()

    def test_chinese_numbers_close_gap_after_deletion(self):
        """删掉中间一章后，后面的中文序号自动接上。"""
        result = self.format_texts(['一、总体要求', '三、实施步骤'])
        self.assertEqual(result[0][0], '一、总体要求')
        self.assertEqual(result[1][0], '二、实施步骤', '跳号应被重排为连续序号')

    def test_chinese_parenthesised_numbers_keep_brackets(self):
        """括号式中文编号保留中文括号。"""
        result = self.format_texts(
            ['（一）适用范围', '（三）术语'],
            styles=['Heading 2', 'Heading 2'],
        )
        self.assertEqual(result[0][0], '（一）适用范围')
        self.assertEqual(result[1][0], '（二）术语')

    def test_multi_level_numbers_follow_parent(self):
        """多级编号的下级取父级当前序号。"""
        result = self.format_texts(
            ['1 甲', '3.1 子', '2 乙', '5.1 子'],
            styles=['Heading 2', 'Heading 3', 'Heading 2', 'Heading 3'],
        )
        self.assertEqual(result[0][0], '1 甲')
        self.assertEqual(result[1][0], '1.1 子', '首个子节应为 1.1')
        self.assertEqual(result[2][0], '2 乙')
        self.assertEqual(result[3][0], '2.1 子', '父级变 2 后子节回到 2.1')

    def test_independent_child_numbers_reset_per_parent(self):
        """下级编号不含父级时，父级切换后从 1 重新开始。"""
        result = self.format_texts(
            ['一、甲', '5. 子', '9. 子', '二、乙', '7. 子'],
            styles=['Heading 2', 'Heading 3', 'Heading 3', 'Heading 2', 'Heading 3'],
        )
        self.assertEqual(result[0][0], '一、甲')
        self.assertEqual(result[1][0], '1. 子')
        self.assertEqual(result[2][0], '2. 子')
        self.assertEqual(result[3][0], '二、乙')
        self.assertEqual(result[4][0], '1. 子', '新一章下的子节重新从 1 开始')

    def test_heading_without_number_is_left_untouched(self):
        """原本没有编号的标题不强行加编号。"""
        result = self.format_texts(
            ['一、甲', '通用表格组件'],
            styles=['Heading 2', 'Heading 4'],
        )
        self.assertEqual(result[1][0], '通用表格组件')

    def test_numbering_can_be_switched_off(self):
        """关闭开关时保持源文档编号原样。"""
        result = self.format_texts(
            ['一、总体要求', '三、实施步骤'],
            auto_heading_numbering=False,
        )
        self.assertEqual(result[1][0], '三、实施步骤', '关闭后不应重排')

    def test_number_followed_by_space_only_is_renumbered(self):
        """一级编号只有数字加空格、没有句点时也要重排。"""
        result = self.format_texts(
            ['1 目的', '2 职责', '4 基线', '4.1 子'],
            styles=['Heading 2', 'Heading 2', 'Heading 2', 'Heading 3'],
        )
        self.assertEqual(result[0][0], '1 目的')
        self.assertEqual(result[1][0], '2 职责')
        self.assertEqual(result[2][0], '3 基线', '缺号应接上')
        self.assertEqual(result[3][0], '3.1 子', '子节跟随新父级序号')


if __name__ == '__main__':
    unittest.main()
