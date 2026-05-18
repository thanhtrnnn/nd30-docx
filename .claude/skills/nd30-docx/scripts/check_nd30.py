#!/usr/bin/env python3
"""
Kiểm tra file .docx tuân thủ thể thức NĐ30/2020/NĐ-CP.
Sử dụng: python check_nd30.py --check <file.docx>
"""
import sys
import os
import zipfile
import xml.etree.ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def mm_to_emu(mm):
    return int(mm * 36000)

def emu_to_mm(emu):
    return emu / 36000

def twip_to_mm(twip):
    return twip * 25.4 / 1440

def extract_text(para):
    """Extract all text from a paragraph element."""
    texts = []
    for t in para.findall('.//w:t', NS):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def check_docx(filepath):
    """Check a .docx file against ND30 rules."""
    results = {
        'errors': [],
        'warnings': [],
        'passed': [],
        'file': filepath,
    }

    if not os.path.exists(filepath):
        results['errors'].append(f"File not found: {filepath}")
        return results

    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            # Read document.xml
            with z.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()

            # --- A. Page setup ---
            sect_pr = root.find('.//w:sectPr', NS)
            if sect_pr is not None:
                pg_sz = sect_pr.find('w:pgSz', NS)
                if pg_sz is not None:
                    w = int(pg_sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', 0))
                    h = int(pg_sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}h', 0))
                    w_mm = twip_to_mm(w)
                    h_mm = twip_to_mm(h)
                    # A4: 210mm x 297mm (tolerance 2mm)
                    if abs(w_mm - 210) > 2 or abs(h_mm - 297) > 2:
                        results['errors'].append(f"Paper size: {w_mm:.0f}mm x {h_mm:.0f}mm (expected A4: 210x297mm)")
                    else:
                        results['passed'].append(f"Paper size: A4 ({w_mm:.0f}mm x {h_mm:.0f}mm)")

                pg_mar = sect_pr.find('w:pgMar', NS)
                if pg_mar is not None:
                    top = twip_to_mm(int(pg_mar.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top', 0)))
                    bottom = twip_to_mm(int(pg_mar.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom', 0)))
                    left = twip_to_mm(int(pg_mar.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left', 0)))
                    right = twip_to_mm(int(pg_mar.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right', 0)))

                    # Check margins (tolerance 0.5mm)
                    margin_checks = [
                        ("Top margin", top, 20, 25),
                        ("Bottom margin", bottom, 20, 25),
                        ("Left margin", left, 30, 35),
                        ("Right margin", right, 15, 20),
                    ]
                    for name, val, min_val, max_val in margin_checks:
                        if val < min_val - 0.5 or val > max_val + 0.5:
                            results['errors'].append(f"{name}: {val:.1f}mm (expected {min_val}-{max_val}mm)")
                        else:
                            results['passed'].append(f"{name}: {val:.1f}mm")

            # --- B. Check font ---
            fonts_found = set()
            for rPr in root.findall('.//w:rPr/w:rFonts', NS):
                ascii_font = rPr.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', '')
                if ascii_font:
                    fonts_found.add(ascii_font)

            if fonts_found and 'Times New Roman' not in fonts_found:
                results['warnings'].append(f"Font: found {fonts_found}, expected Times New Roman")
            elif 'Times New Roman' in fonts_found:
                results['passed'].append("Font: Times New Roman")

            # --- C. Check national emblem ---
            all_text = []
            for para in root.findall('.//w:p', NS):
                t = extract_text(para)
                if t.strip():
                    all_text.append(t.strip())

            full_text = ' '.join(all_text).upper()

            if 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' in full_text or 'CONG HOA XA HOI CHU NGHIA VIET NAM' in full_text:
                results['passed'].append("National emblem text: present")
            else:
                results['warnings'].append("National emblem text: not found (may be in header table)")

            # --- D. Check motto ---
            if 'ĐỘC LẬP - TỰ DO - HẠNH PHÚC' in full_text or 'DOC LAP - TU DO - HANH PHUC' in full_text:
                results['passed'].append("Motto text: present")
            else:
                results['warnings'].append("Motto text: not found")

            # --- E. Check document number ---
            has_so = any('Số:' in t or 'Số:' in t or 'So:' in t.upper() for t in all_text)
            if has_so:
                results['passed'].append("Document number (Số:): present")
            else:
                results['warnings'].append("Document number (Số:): not found")

            # --- F. Check date ---
            has_date = any('ngày' in t.lower() or 'ngay' in t.lower() for t in all_text)
            if has_date:
                results['passed'].append("Date field: present")
            else:
                results['warnings'].append("Date field: not found")

            # --- G. Check noi nhan ---
            has_noi_nhan = any('nơi nhận' in t.lower() or 'noi nhan' in t.lower() for t in all_text)
            if has_noi_nhan:
                results['passed'].append("Recipients (Nơi nhận): present")
            else:
                results['warnings'].append("Recipients (Nơi nhận): not found")

            # --- H. Check "Kính gửi" formatting ---
            for para in root.findall('.//w:p', NS):
                t = extract_text(para)
                if 'Kính gửi' in t or 'kính gửi' in t:
                    # Check italic
                    is_italic = False
                    for rPr in para.findall('.//w:rPr', NS):
                        i_elem = rPr.find('w:i', NS)
                        if i_elem is not None:
                            val = i_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
                            if val.lower() not in ('false', '0'):
                                is_italic = True
                    if is_italic:
                        results['errors'].append("Kính gửi: italic (NĐ30 requires 'đứng' = not italic)")
                    else:
                        results['passed'].append("Kính gửi: not italic (correct)")

                    # Check first-line indent
                    pPr = para.find('w:pPr', NS)
                    has_indent = False
                    if pPr is not None:
                        ind = pPr.find('w:ind', NS)
                        if ind is not None:
                            fl = ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine', '0')
                            if int(fl) >= 500:
                                has_indent = True
                    if has_indent:
                        results['passed'].append("Kính gửi: first-line indent ≥ 500 DXA")
                    else:
                        results['warnings'].append("Kính gửi: no first-line indent (expected ~567 DXA / 1cm)")
                    break

            # --- I. Check separator lines (underscore characters) ---
            for para in root.findall('.//w:p', NS):
                t = extract_text(para).strip()
                if t and all(c == '_' for c in t) and 10 <= len(t) <= 35:
                    results['passed'].append(f"Separator line: underscore chars ({len(t)} chars)")
                    break
            else:
                results['warnings'].append("Separator line: no underscore separator found")

            # --- J. Check line spacing ---
            spacing_issues = []
            for para in root.findall('.//w:p', NS):
                t = extract_text(para).strip()
                if not t or len(t) < 5:
                    continue
                pPr = para.find('w:pPr', NS)
                if pPr is not None:
                    sp = pPr.find('w:spacing', NS)
                    if sp is not None:
                        line = sp.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', '')
                        rule = sp.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', '')
                        if line and rule == 'exact':
                            line_val = int(line)
                            if line_val > 360:
                                spacing_issues.append(f"Line spacing {line_val} DXA > 360 (1.5 lines)")
            if not spacing_issues:
                results['passed'].append("Line spacing: within spec (single to 1.5 lines)")
            else:
                results['warnings'].append(f"Line spacing: {len(spacing_issues)} paragraph(s) exceed 1.5 lines")

            # --- K. Check "Nơi nhận" label formatting ---
            for para in root.findall('.//w:p', NS):
                t = extract_text(para).strip()
                if t.startswith('Nơi nhận'):
                    # Check italic + bold
                    is_italic = False
                    is_bold = False
                    for rPr in para.findall('.//w:rPr', NS):
                        if rPr.find('w:i', NS) is not None:
                            is_italic = True
                        if rPr.find('w:b', NS) is not None:
                            is_bold = True
                    if is_italic and is_bold:
                        results['passed'].append("Nơi nhận label: italic + bold (correct)")
                    else:
                        results['warnings'].append(f"Nơi nhận label: italic={is_italic}, bold={is_bold} (expected both)")
                    break

    except zipfile.BadZipFile:
        results['errors'].append(f"Not a valid .docx file: {filepath}")
    except Exception as e:
        results['errors'].append(f"Error reading file: {e}")

    return results

def print_report(results):
    """Print a formatted report."""
    print(f"\n{'='*60}")
    print(f"ND30 COMPLIANCE CHECK: {os.path.basename(results['file'])}")
    print(f"{'='*60}")

    if results['passed']:
        print(f"\n✓ PASSED ({len(results['passed'])}):")
        for item in results['passed']:
            print(f"  ✓ {item}")

    if results['warnings']:
        print(f"\n⚠ WARNINGS ({len(results['warnings'])}):")
        for item in results['warnings']:
            print(f"  ⚠ {item}")

    if results['errors']:
        print(f"\n✗ ERRORS ({len(results['errors'])}):")
        for item in results['errors']:
            print(f"  ✗ {item}")

    total = len(results['passed']) + len(results['warnings']) + len(results['errors'])
    print(f"\nSummary: {len(results['passed'])}/{total} passed, {len(results['warnings'])} warnings, {len(results['errors'])} errors")
    print(f"{'='*60}\n")

def main():
    if len(sys.argv) < 3 or sys.argv[1] != '--check':
        print("Usage: python check_nd30.py --check <file.docx>")
        print("       python check_nd30.py --check <directory>")
        sys.exit(1)

    target = sys.argv[2]
    if os.path.isdir(target):
        for f in sorted(os.listdir(target)):
            if f.endswith('.docx') and not f.startswith('~'):
                results = check_docx(os.path.join(target, f))
                print_report(results)
    else:
        results = check_docx(target)
        print_report(results)
        sys.exit(1 if results['errors'] else 0)

if __name__ == '__main__':
    main()
