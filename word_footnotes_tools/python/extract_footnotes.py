#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
أداة ذكية لاستخراج الحواشي السفلية من مستندات Word وتحويلها إلى نص عادي
Smart tool to extract footnotes from Word documents and convert them to plain text
"""

import sys
import os
import zipfile
import shutil
from lxml import etree
from pathlib import Path
import re


class FootnoteExtractor:
    """فئة لاستخراج الحواشي السفلية من مستندات Word"""

    # تعريف مساحات الأسماء XML
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    def __init__(self, input_file):
        """
        تهيئة المستخرج

        Args:
            input_file: مسار ملف Word المدخل
        """
        self.input_file = Path(input_file)
        if not self.input_file.exists():
            raise FileNotFoundError(f"الملف غير موجود: {input_file}")

        self.footnotes = {}
        self.endnotes = {}

    def extract(self, output_file=None):
        """
        استخراج الحواشي وتحويلها إلى نص بين &&

        Args:
            output_file: مسار ملف الإخراج (اختياري)

        Returns:
            مسار ملف الإخراج
        """
        if output_file is None:
            output_file = self.input_file.parent / f"{self.input_file.stem}_extracted{self.input_file.suffix}"
        else:
            output_file = Path(output_file)

        # نسخ الملف إلى موقع مؤقت
        temp_dir = Path('temp_word_extraction')
        temp_dir.mkdir(exist_ok=True)

        try:
            # فك ضغط ملف Word
            with zipfile.ZipFile(self.input_file, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # قراءة الحواشي السفلية
            self._read_footnotes(temp_dir)
            self._read_endnotes(temp_dir)

            # تعديل المستند الرئيسي
            self._modify_document(temp_dir)

            # إعادة ضغط الملف
            self._zip_directory(temp_dir, output_file)

            print(f"✓ تم الاستخراج بنجاح!")
            print(f"  - عدد الحواشي السفلية: {len(self.footnotes)}")
            print(f"  - عدد الهوامش النهائية: {len(self.endnotes)}")
            print(f"  - ملف الإخراج: {output_file}")

            return output_file

        finally:
            # حذف الملفات المؤقتة
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    def _read_footnotes(self, temp_dir):
        """قراءة الحواشي السفلية من footnotes.xml"""
        footnotes_file = temp_dir / 'word' / 'footnotes.xml'
        if not footnotes_file.exists():
            return

        tree = etree.parse(str(footnotes_file))
        root = tree.getroot()

        for footnote in root.findall('.//w:footnote', self.NAMESPACES):
            footnote_id = footnote.get('{' + self.NAMESPACES['w'] + '}id')
            if footnote_id and footnote_id not in ['-1', '0']:  # تجاهل الحواشي الخاصة
                text = self._extract_text_from_element(footnote)
                self.footnotes[footnote_id] = text

    def _read_endnotes(self, temp_dir):
        """قراءة الهوامش النهائية من endnotes.xml"""
        endnotes_file = temp_dir / 'word' / 'endnotes.xml'
        if not endnotes_file.exists():
            return

        tree = etree.parse(str(endnotes_file))
        root = tree.getroot()

        for endnote in root.findall('.//w:endnote', self.NAMESPACES):
            endnote_id = endnote.get('{' + self.NAMESPACES['w'] + '}id')
            if endnote_id and endnote_id not in ['-1', '0']:
                text = self._extract_text_from_element(endnote)
                self.endnotes[endnote_id] = text

    def _extract_text_from_element(self, element):
        """استخراج النص من عنصر XML"""
        texts = []
        for text_elem in element.findall('.//w:t', self.NAMESPACES):
            if text_elem.text:
                texts.append(text_elem.text)
        return ''.join(texts).strip()

    def _modify_document(self, temp_dir):
        """تعديل المستند الرئيسي لاستبدال مراجع الحواشي بالنص"""
        doc_file = temp_dir / 'word' / 'document.xml'
        if not doc_file.exists():
            raise FileNotFoundError("ملف document.xml غير موجود")

        tree = etree.parse(str(doc_file))
        root = tree.getroot()

        # البحث عن مراجع الحواشي السفلية
        for footnote_ref in root.findall('.//w:footnoteReference', self.NAMESPACES):
            footnote_id = footnote_ref.get('{' + self.NAMESPACES['w'] + '}id')
            if footnote_id in self.footnotes:
                self._replace_reference_with_text(footnote_ref, self.footnotes[footnote_id])

        # البحث عن مراجع الهوامش النهائية
        for endnote_ref in root.findall('.//w:endnoteReference', self.NAMESPACES):
            endnote_id = endnote_ref.get('{' + self.NAMESPACES['w'] + '}id')
            if endnote_id in self.endnotes:
                self._replace_reference_with_text(endnote_ref, self.endnotes[endnote_id])

        # حفظ المستند المعدل
        tree.write(str(doc_file), xml_declaration=True, encoding='UTF-8', standalone=True)

    def _replace_reference_with_text(self, ref_element, footnote_text):
        """استبدال مرجع الحاشية بالنص بين &&"""
        # الحصول على العنصر الأب (w:r - run)
        run = ref_element.getparent()
        if run is None:
            return

        # إنشاء عنصر نص جديد
        new_run = etree.Element('{' + self.NAMESPACES['w'] + '}r')
        text_elem = etree.SubElement(new_run, '{' + self.NAMESPACES['w'] + '}t')
        text_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        text_elem.text = f" &&{footnote_text}&& "

        # إدراج النص الجديد بعد الـ run الحالي
        parent = run.getparent()
        index = list(parent).index(run)
        parent.insert(index + 1, new_run)

        # حذف الـ run الذي يحتوي على مرجع الحاشية
        parent.remove(run)

    def _zip_directory(self, source_dir, output_file):
        """ضغط المجلد إلى ملف Word"""
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(source_dir)
                    zipf.write(file_path, arcname)


def main():
    """الدالة الرئيسية"""
    if len(sys.argv) < 2:
        print("الاستخدام: python extract_footnotes.py <input_file.docx> [output_file.docx]")
        print("Usage: python extract_footnotes.py <input_file.docx> [output_file.docx]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        extractor = FootnoteExtractor(input_file)
        output = extractor.extract(output_file)
        print(f"\n✓ تم بنجاح! الملف الجديد: {output}")
    except Exception as e:
        print(f"✗ خطأ: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
