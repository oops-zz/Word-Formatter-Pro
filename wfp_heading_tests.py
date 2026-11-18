# -*- coding: utf-8 -*-
"""标题识别：正文里的有序列举不能被当成章节标题；带标题样式的段落必须被识别。

背景（实测）：
1. ``wfp_core.py`` 的 ``RE_HEADING_H3 = ^\\d+\\s*[.．]`` 会把 "1. xxx；" 这类
   正文列举一并判为三级标题并写入 ``w:outlineLvl``。真实样本"软件物料清单
   （SBOM）管理规程"里有 10 条正文列举被这样污染。
2. 同一份文档有 27 条 ``Heading 1/2/3`` 样式段落，但程序从不读 ``w:pStyle``，
   只靠文本前缀猜。文本没有 "一、"/"（一）"/"1." 前缀的标题（如
   "1 目的与适用范围"）全部被当成正文。
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
    """段落直接写入的大纲级别（1 起），未设置返回 None。"""
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


class HeadingRecognitionTests(unittest.TestCase):
    def processor(self, **options):
        config = dict(DEFAULT_CONFIG)
        config.update(options)
        return WordProcessor(config)

    def format_texts(self, texts, styles=None, **options):
        """排版一组段落，返回每段的大纲级别列表。"""
        styles = styles or [None] * len(texts)
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
                return [outline_level(para) for para in Document(output).paragraphs]
            finally:
                processor._cleanup_temp_files()

    def test_ordered_list_items_in_body_are_not_headings(self):
        """以分号/句号收尾的长列举是正文，不能带大纲级别。"""
        levels = self.format_texts([
            '3.1 术语',
            '1. 同一组件同时出现在白名单与黑名单时以黑名单为准，白名单条目同步标记为暂缓准入或移出；',
            '2. 命中黑名单但业务确需使用的，只能走第 6 章例外流程，不得通过修改白名单状态绕过；',
            '3. 例外条目不计入已批准组件统计，须单列。',
        ])
        self.assertEqual(levels[0], 3, '「3.1 术语」应识别为三级标题')
        self.assertIsNone(levels[1], '第一条正文列举不应带大纲级别')
        self.assertIsNone(levels[2], '第二条正文列举不应带大纲级别')
        self.assertIsNone(levels[3], '第三条正文列举不应带大纲级别')

    def test_short_numbered_heading_stays_a_heading(self):
        """收紧判定不能把真正的短编号标题一起排除。"""
        levels = self.format_texts([
            '1. 总则',
            '2. 适用范围',
        ])
        self.assertEqual(levels[0], 3, '「1. 总则」仍是三级标题')
        self.assertEqual(levels[1], 3, '「2. 适用范围」仍是三级标题')

    def test_heading_style_without_number_prefix_is_recognized(self):
        """套了 Heading N 样式但文本没有编号前缀，也要识别成对应级别。"""
        levels = self.format_texts(
            ['软件物料清单管理规程', '目的与适用范围', '术语与依据'],
            styles=['Heading 1', 'Heading 2', 'Heading 3'],
        )
        self.assertEqual(levels[0], 1, 'Heading 1 样式应识别为一级标题')
        self.assertEqual(levels[1], 2, 'Heading 2 样式应识别为二级标题')
        self.assertEqual(levels[2], 3, 'Heading 3 样式应识别为三级标题')

    def test_plain_body_paragraph_without_style_stays_body(self):
        """既无标题样式也无编号前缀的段落仍是正文。"""
        levels = self.format_texts([
            '本规程适用于公司范围内所有自研与外购组件的物料清单管理。',
        ])
        self.assertIsNone(levels[0], '普通正文不应带大纲级别')


if __name__ == '__main__':
    unittest.main()
