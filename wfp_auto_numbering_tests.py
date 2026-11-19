# -*- coding: utf-8 -*-
"""标题分级自动编号：文字编号应替换为 numPr 多级列表。

目标形态与 WPS/Word 的"多级列表"一致：编号由文档结构生成，
增删章节时编辑器自动重排，而不是每次跑工具重写文字。
"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn

from wfp_config import DEFAULT_CONFIG
from wfp_core import WordProcessor


def numpr_of(paragraph):
    ppr = paragraph._p.pPr
    if ppr is None:
        return None
    return ppr.find(qn('w:numPr'))


class AutoNumberingTests(unittest.TestCase):
    def processor(self, **options):
        config = dict(DEFAULT_CONFIG)
        config.update(options)
        return WordProcessor(config)

    def format_doc(self, rows, **options):
        """排版一组 (style, text) 段落，返回输出 Document。"""
        options.setdefault('auto_heading_numbering', True)
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source.docx'
            output = Path(directory) / 'output.docx'
            doc = Document()
            for style, text in rows:
                if style is None:
                    doc.add_paragraph(text)
                else:
                    doc.add_paragraph(text, style=style)
            doc.save(source)

            processor = self.processor(**options)
            try:
                with patch.object(processor, '_com_available', return_value=False):
                    processor.format_document(str(source), str(output))
                return Document(output)
            finally:
                processor._cleanup_temp_files()

    def test_heading_text_loses_typed_number(self):
        """编号文字应被移除，改由 numPr 提供。"""
        doc = self.format_doc([
            ('Heading 2', '一、总体要求'),
            ('Heading 2', '三、实施步骤'),
        ])
        texts = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
        self.assertEqual(texts, ['总体要求', '实施步骤'], '文字编号应被清掉')

    def test_headings_carry_numpr(self):
        """标题段落必须挂上 numPr，否则编辑器不会自动编号。"""
        doc = self.format_doc([
            ('Heading 2', '1 总体要求'),
            ('Heading 3', '1.1 适用范围'),
            ('Heading 2', '2 实施步骤'),
        ])
        heads = [para for para in doc.paragraphs if numpr_of(para) is not None]
        self.assertEqual(len(heads), 3, '三个标题都应挂上 numPr')

        levels = []
        for para in heads:
            node = numpr_of(para).find(qn('w:ilvl'))
            levels.append(node.get(qn('w:val')))
        self.assertEqual(levels, ['1', '2', '1'], 'ilvl 应与标题层级对应')

    def test_numbering_xml_gets_multilevel_definition(self):
        """numbering.xml 里要有多级定义，numPr 才有东西可引用。"""
        doc = self.format_doc([
            ('Heading 2', '一、总体要求'),
            ('Heading 3', '（一）适用范围'),
        ])
        numbering = doc.part.numbering_part.element
        abstracts = numbering.findall(qn('w:abstractNum'))
        self.assertTrue(abstracts, '应写入 abstractNum')
        multi = abstracts[-1].find(qn('w:multiLevelType'))
        self.assertIsNotNone(multi)
        self.assertEqual(multi.get(qn('w:val')), 'multilevel')

        texts = [node.get(qn('w:val'))
                 for node in abstracts[-1].findall(qn('w:lvl') + '/' + qn('w:lvlText'))]
        self.assertIn('%1、', texts, '一级应为中文顿号形态')
        self.assertIn('（%2）', texts, '二级应为中文括号形态')

    def test_multi_level_document_keeps_dotted_scheme(self):
        """1 / 3.1 这类多级编号应生成 %1.%2 形态。"""
        doc = self.format_doc([
            ('Heading 2', '1 目的'),
            ('Heading 3', '3.1 术语'),
        ])
        numbering = doc.part.numbering_part.element
        abstract = numbering.findall(qn('w:abstractNum'))[-1]
        texts = [node.get(qn('w:val'))
                 for node in abstract.findall(qn('w:lvl') + '/' + qn('w:lvlText'))]
        self.assertIn('%2.%3', texts, '三级应拼接父级（二级）序号')

    def test_numbering_levels_carry_heading_font(self):
        """编号符号必须带标题层级的字体，否则会继承默认字体。"""
        doc = self.format_doc([
            ('Heading 2', '一、总体要求'),
            ('Heading 3', '（一）适用范围'),
        ])
        numbering = doc.part.numbering_part.element
        abstract = numbering.findall(qn('w:abstractNum'))[-1]
        levels = abstract.findall(qn('w:lvl'))

        for lvl in levels[:3]:
            self.assertIsNotNone(lvl.find(qn('w:rPr')),
                                 '每一级编号都应带 rPr')

        second = levels[1].find(qn('w:rPr')).find(qn('w:rFonts'))
        self.assertIsNotNone(second, '二级编号应有字体定义')
        self.assertEqual(second.get(qn('w:eastAsia')),
                         DEFAULT_CONFIG['h2_font'],
                         '二级编号应使用二级标题字体')

    def test_disabled_leaves_text_untouched(self):
        """关闭开关时保持原文字编号。"""
        doc = self.format_doc(
            [('Heading 2', '三、实施步骤')],
            auto_heading_numbering=False,
        )
        texts = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
        self.assertIn('三、实施步骤', texts)
        self.assertEqual([p for p in doc.paragraphs if numpr_of(p) is not None], [])


if __name__ == '__main__':
    unittest.main()
