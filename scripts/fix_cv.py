#!/usr/bin/env python3
"""Sửa CV: font Calibri → Times New Roman, lề trái 25mm → 30mm"""
import sys
from docx import Document
from docx.shared import Mm, Pt
from docx.oxml.ns import qn

def fix_cv(input_path, output_path):
    doc = Document(input_path)

    # Sửa lề trang
    for section in doc.sections:
        section.left_margin = Mm(30)

    # Sửa font tất cả paragraph và run
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.font.name == 'Calibri' or run.font.name is None:
                run.font.name = 'Times New Roman'
                # Sửa cả font Đông Á
                r = run._element
                rPr = r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = r.makeelement(qn('w:rPr'), {})
                    r.insert(0, rPr)
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = rPr.makeelement(qn('w:rFonts'), {})
                    rPr.insert(0, rFonts)
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # Sửa font trong bảng (nếu có)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name == 'Calibri' or run.font.name is None:
                            run.font.name = 'Times New Roman'
                            r = run._element
                            rPr = r.find(qn('w:rPr'))
                            if rPr is None:
                                rPr = r.makeelement(qn('w:rPr'), {})
                                r.insert(0, rPr)
                            rFonts = rPr.find(qn('w:rFonts'))
                            if rFonts is None:
                                rFonts = rPr.makeelement(qn('w:rFonts'), {})
                                rPr.insert(0, rFonts)
                            rFonts.set(qn('w:ascii'), 'Times New Roman')
                            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                            rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    doc.save(output_path)
    print(f"Đã sửa CV: {output_path}")

if __name__ == '__main__':
    input_path = sys.argv[1] if len(sys.argv) > 1 else 'output/CV_PhamTuanAnh_AIEngineerIntern.docx'
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'output/CV_PhamTuanAnh_AIEngineerIntern_FIXED.docx'
    fix_cv(input_path, output_path)
