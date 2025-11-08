#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
أداة ذكية لإعادة الحواشي السفلية إلى مستندات Word من النص العادي
Smart tool to restore footnotes to Word documents from plain text
"""

import sys
import os
import zipfile
import shutil
from lxml import etree
from pathlib import Path
import re


class FootnoteRestorer:
    """فئة لإعادة الحواشي السفلية إلى مستندات Word"""

    # تعريف مساحات الأسماء XML
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    def __init__(self, input_file):
        """
        تهيئة المستعيد

        Args:
            input_file: مسار ملف Word المدخل
        """
        self.input_file = Path(input_file)
        if not self.input_file.exists():
            raise FileNotFoundError(f"الملف غير موجود: {input_file}")

        self.footnote_counter = 1

    def restore(self, output_file=None):
        """
        إعادة الحواشي من النص بين && إلى حواشي سفلية عادية

        Args:
            output_file: مسار ملف الإخراج (اختياري)

        Returns:
            مسار ملف الإخراج
        """
        if output_file is None:
            output_file = self.input_file.parent / f"{self.input_file.stem}_restored{self.input_file.suffix}"
        else:
            output_file = Path(output_file)

        # نسخ الملف إلى موقع مؤقت
        temp_dir = Path('temp_word_restoration')
        temp_dir.mkdir(exist_ok=True)

        try:
            # فك ضغط ملف Word
            with zipfile.ZipFile(self.input_file, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # تعديل المستند الرئيسي وإنشاء الحواشي
            footnotes_data = self._modify_document(temp_dir)

            # إنشاء/تحديث ملف الحواشي
            if footnotes_data:
                self._create_footnotes_file(temp_dir, footnotes_data)
                self._update_relationships(temp_dir)

            # إعادة ضغط الملف
            self._zip_directory(temp_dir, output_file)

            print(f"✓ تمت الإعادة بنجاح!")
            print(f"  - عدد الحواشي المستعادة: {len(footnotes_data)}")
            print(f"  - ملف الإخراج: {output_file}")

            return output_file

        finally:
            # حذف الملفات المؤقتة
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    def _modify_document(self, temp_dir):
        """تعديل المستند الرئيسي لاستبدال النص بمراجع الحواشي"""
        doc_file = temp_dir / 'word' / 'document.xml'
        if not doc_file.exists():
            raise FileNotFoundError("ملف document.xml غير موجود")

        tree = etree.parse(str(doc_file))
        root = tree.getroot()

        footnotes_data = []

        # البحث عن جميع عناصر النص
        for para in root.findall('.//w:p', self.NAMESPACES):
            self._process_paragraph(para, footnotes_data)

        # حفظ المستند المعدل
        tree.write(str(doc_file), xml_declaration=True, encoding='UTF-8', standalone=True)

        return footnotes_data

    def _process_paragraph(self, paragraph, footnotes_data):
        """معالجة فقرة واحدة للبحث عن نص الحواشي"""
        runs = paragraph.findall('.//w:r', self.NAMESPACES)

        for run in runs:
            text_elems = run.findall('.//w:t', self.NAMESPACES)

            for text_elem in text_elems:
                if text_elem.text:
                    # البحث عن نص بين &&
                    matches = list(re.finditer(r'&&([^&]+)&&', text_elem.text))

                    if matches:
                        # معالجة كل تطابق
                        self._process_matches(run, text_elem, matches, footnotes_data, paragraph)

    def _process_matches(self, run, text_elem, matches, footnotes_data, paragraph):
        """معالجة التطابقات وإنشاء مراجع الحواشي"""
        original_text = text_elem.text
        parent = run.getparent()
        run_index = list(parent).index(run)

        # تقسيم النص ومعالجة كل جزء
        last_end = 0
        elements_to_insert = []

        for match in matches:
            start, end = match.span()
            footnote_text = match.group(1)

            # النص قبل الحاشية
            if start > last_end:
                before_text = original_text[last_end:start]
                if before_text:
                    new_run = self._create_text_run(before_text)
                    elements_to_insert.append(new_run)

            # إنشاء مرجع الحاشية
            footnote_id = str(self.footnote_counter)
            footnotes_data.append({
                'id': footnote_id,
                'text': footnote_text
            })

            footnote_run = self._create_footnote_reference(footnote_id)
            elements_to_insert.append(footnote_run)

            self.footnote_counter += 1
            last_end = end

        # النص بعد آخر حاشية
        if last_end < len(original_text):
            after_text = original_text[last_end:]
            if after_text:
                new_run = self._create_text_run(after_text)
                elements_to_insert.append(new_run)

        # حذف الـ run الأصلي وإدراج العناصر الجديدة
        if elements_to_insert:
            parent.remove(run)
            for i, elem in enumerate(elements_to_insert):
                parent.insert(run_index + i, elem)

    def _create_text_run(self, text):
        """إنشاء عنصر run يحتوي على نص"""
        run = etree.Element('{' + self.NAMESPACES['w'] + '}r')
        text_elem = etree.SubElement(run, '{' + self.NAMESPACES['w'] + '}t')
        text_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        text_elem.text = text
        return run

    def _create_footnote_reference(self, footnote_id):
        """إنشاء مرجع حاشية سفلية"""
        run = etree.Element('{' + self.NAMESPACES['w'] + '}r')

        # إضافة خصائص الحاشية
        rPr = etree.SubElement(run, '{' + self.NAMESPACES['w'] + '}rPr')
        rStyle = etree.SubElement(rPr, '{' + self.NAMESPACES['w'] + '}rStyle')
        rStyle.set('{' + self.NAMESPACES['w'] + '}val', 'FootnoteReference')

        # إضافة مرجع الحاشية
        footnote_ref = etree.SubElement(run, '{' + self.NAMESPACES['w'] + '}footnoteReference')
        footnote_ref.set('{' + self.NAMESPACES['w'] + '}id', footnote_id)

        return run

    def _create_footnotes_file(self, temp_dir, footnotes_data):
        """إنشاء أو تحديث ملف footnotes.xml"""
        footnotes_file = temp_dir / 'word' / 'footnotes.xml'

        # إنشاء العنصر الجذر
        root = etree.Element(
            '{' + self.NAMESPACES['w'] + '}footnotes',
            nsmap={'w': self.NAMESPACES['w'], 'r': self.NAMESPACES['r']}
        )

        # إضافة الحاشيتين الافتراضيتين (separator و continuation separator)
        self._add_default_footnotes(root)

        # إضافة الحواشي الفعلية
        for footnote_data in footnotes_data:
            self._add_footnote(root, footnote_data['id'], footnote_data['text'])

        # كتابة الملف
        tree = etree.ElementTree(root)
        tree.write(str(footnotes_file), xml_declaration=True, encoding='UTF-8', standalone=True, pretty_print=True)

    def _add_default_footnotes(self, root):
        """إضافة الحواشي الافتراضية"""
        # Separator footnote
        separator = etree.SubElement(root, '{' + self.NAMESPACES['w'] + '}footnote')
        separator.set('{' + self.NAMESPACES['w'] + '}type', 'separator')
        separator.set('{' + self.NAMESPACES['w'] + '}id', '-1')

        p = etree.SubElement(separator, '{' + self.NAMESPACES['w'] + '}p')
        r = etree.SubElement(p, '{' + self.NAMESPACES['w'] + '}r')
        sep = etree.SubElement(r, '{' + self.NAMESPACES['w'] + '}separator')

        # Continuation separator
        cont_separator = etree.SubElement(root, '{' + self.NAMESPACES['w'] + '}footnote')
        cont_separator.set('{' + self.NAMESPACES['w'] + '}type', 'continuationSeparator')
        cont_separator.set('{' + self.NAMESPACES['w'] + '}id', '0')

        p = etree.SubElement(cont_separator, '{' + self.NAMESPACES['w'] + '}p')
        r = etree.SubElement(p, '{' + self.NAMESPACES['w'] + '}r')
        cont_sep = etree.SubElement(r, '{' + self.NAMESPACES['w'] + '}continuationSeparator')

    def _add_footnote(self, root, footnote_id, text):
        """إضافة حاشية سفلية"""
        footnote = etree.SubElement(root, '{' + self.NAMESPACES['w'] + '}footnote')
        footnote.set('{' + self.NAMESPACES['w'] + '}id', footnote_id)

        p = etree.SubElement(footnote, '{' + self.NAMESPACES['w'] + '}p')

        # إضافة رقم الحاشية
        pPr = etree.SubElement(p, '{' + self.NAMESPACES['w'] + '}pPr')
        pStyle = etree.SubElement(pPr, '{' + self.NAMESPACES['w'] + '}pStyle')
        pStyle.set('{' + self.NAMESPACES['w'] + '}val', 'FootnoteText')

        # إضافة مرجع الحاشية في البداية
        r = etree.SubElement(p, '{' + self.NAMESPACES['w'] + '}r')
        rPr = etree.SubElement(r, '{' + self.NAMESPACES['w'] + '}rPr')
        rStyle = etree.SubElement(rPr, '{' + self.NAMESPACES['w'] + '}rStyle')
        rStyle.set('{' + self.NAMESPACES['w'] + '}val', 'FootnoteReference')
        footnoteRef = etree.SubElement(r, '{' + self.NAMESPACES['w'] + '}footnoteRef')

        # إضافة النص
        r_text = etree.SubElement(p, '{' + self.NAMESPACES['w'] + '}r')
        t = etree.SubElement(r_text, '{' + self.NAMESPACES['w'] + '}t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = ' ' + text

    def _update_relationships(self, temp_dir):
        """تحديث ملف العلاقات لإضافة الحواشي"""
        rels_file = temp_dir / 'word' / '_rels' / 'document.xml.rels'

        if not rels_file.exists():
            # إنشاء ملف جديد
            rels_dir = temp_dir / 'word' / '_rels'
            rels_dir.mkdir(exist_ok=True)

        # قراءة أو إنشاء ملف العلاقات
        if rels_file.exists():
            tree = etree.parse(str(rels_file))
            root = tree.getroot()
        else:
            root = etree.Element(
                '{http://schemas.openxmlformats.org/package/2006/relationships}Relationships'
            )
            tree = etree.ElementTree(root)

        # التحقق من وجود علاقة الحواشي
        footnotes_rel_exists = False
        for rel in root.findall(
            './/{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
        ):
            if rel.get('Target') == 'footnotes.xml':
                footnotes_rel_exists = True
                break

        # إضافة علاقة الحواشي إذا لم تكن موجودة
        if not footnotes_rel_exists:
            # الحصول على أكبر rId
            max_id = 0
            for rel in root.findall(
                './/{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
            ):
                rel_id = rel.get('Id', '')
                if rel_id.startswith('rId'):
                    try:
                        num = int(rel_id[3:])
                        max_id = max(max_id, num)
                    except:
                        pass

            new_rel = etree.SubElement(
                root,
                '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
            )
            new_rel.set('Id', f'rId{max_id + 1}')
            new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes')
            new_rel.set('Target', 'footnotes.xml')

        tree.write(str(rels_file), xml_declaration=True, encoding='UTF-8', standalone=True)

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
        print("الاستخدام: python restore_footnotes.py <input_file.docx> [output_file.docx]")
        print("Usage: python restore_footnotes.py <input_file.docx> [output_file.docx]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        restorer = FootnoteRestorer(input_file)
        output = restorer.restore(output_file)
        print(f"\n✓ تم بنجاح! الملف الجديد: {output}")
    except Exception as e:
        print(f"✗ خطأ: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
