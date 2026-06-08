#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
聘书批量生成 - 2026年压力管道检验人员培训
操作模板 run 级替换，保留格式
"""

import sys, re, copy, os, zipfile, shutil, tempfile
from lxml import etree
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')

# ========== CONFIG ==========
MEETING_NAME = "2026年压力管道检验人员培训"
START_DATE = "5月18日"
END_DATE = "5月26日"
LOCATION = "河北省廊坊市"
ADDRESS = "廊坊新世界酒店（河北省廊坊市广阳区建设路158号）"
LETTER_DATE = "2026年5月10日"

SCHEDULE_PATH = "C:/Users/dbz13/Documents/xwechat_files/wxid_ql8hr345qxsh41_0416/msg/file/2026-05/2026年压力管道检验员培训课程安排(1).docx"
TEMPLATE_PATH = "C:/Users/dbz13/Documents/xwechat_files/wxid_ql8hr345qxsh41_0416/msg/file/2026-05/聘书模板.docx"
OUTPUT_PATH = "C:/Users/dbz13/Documents/xwechat_files/wxid_ql8hr345qxsh41_0416/msg/file/2026-05/聘书_全部.docx"

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': W_NS}


# ========== STEP 1: Parse schedule ==========
def get_cell_text(tc):
    texts = []
    for p in tc.findall('.//w:p', NSMAP):
        t_nodes = p.findall('.//w:t', NSMAP)
        pt = ''.join(t.text or '' for t in t_nodes)
        if pt.strip():
            texts.append(pt.strip())
    return '\n'.join(texts)


def parse_schedule(docx_path):
    doc = Document(docx_path)
    root = etree.fromstring(doc.part.blob)
    tables = root.findall('.//w:tbl', NSMAP)
    if not tables:
        print("No table found!")
        return []
    table = tables[0]
    rows = table.findall('./w:tr', NSMAP)

    experts = []
    carry_date = ''
    carry_course = ''
    carry_expert = {}

    for row_idx, row in enumerate(rows):
        if row_idx == 0:
            continue
        cells = row.findall('./w:tc', NSMAP)
        if len(cells) < 4:
            continue

        # --- Date ---
        date_vmerge = cells[0].find('.//w:vMerge', NSMAP)
        date_is_restart = date_vmerge is not None and date_vmerge.get(f'{{{W_NS}}}val') == 'restart'
        date_is_continue = date_vmerge is not None and not date_is_restart

        if date_is_restart or date_vmerge is None:
            carry_date = get_cell_text(cells[0])
        date_str = carry_date

        # --- Time ---
        time_text = get_cell_text(cells[1])

        # --- Course ---
        course_vmerge = cells[2].find('.//w:vMerge', NSMAP)
        course_is_restart = course_vmerge is not None and course_vmerge.get(f'{{{W_NS}}}val') == 'restart'
        course_is_continue = course_vmerge is not None and not course_is_restart

        if course_is_restart or course_vmerge is None:
            carry_course = get_cell_text(cells[2])
        course_text = carry_course

        # Skip header-like or empty rows
        if not time_text or '报到' in time_text or '开班' in time_text:
            continue
        if not course_text or '报到' in course_text or '开班' in course_text:
            continue

        # --- Expert ---
        expert_vmerge = cells[3].find('.//w:vMerge', NSMAP)
        expert_is_restart = expert_vmerge is not None and expert_vmerge.get(f'{{{W_NS}}}val') == 'restart'
        expert_is_continue = expert_vmerge is not None and not expert_is_restart

        if expert_is_restart or expert_vmerge is None:
            raw = get_cell_text(cells[3])
            if not raw.strip():
                continue
            carry_expert = parse_expert(raw)

        if not carry_expert.get('name'):
            continue

        # Period
        period = ''
        if '上午' in time_text:
            period = '上午'
        elif '下午' in time_text:
            period = '下午'
        else:
            tm = re.search(r'(\d+):(\d+)\s*-\s*(\d+):(\d+)', time_text)
            if tm:
                end_h = int(tm.group(3))
                period = '上午' if end_h <= 12 else '下午'

        # Date string (handle optional space like "5月 19日")
        dm = re.search(r'(\d+月\s*\d+日)', date_str)
        date_clean = dm.group(1) if dm else date_str.strip()
        date_clean = re.sub(r'\s+', '', date_clean)  # remove spaces

        experts.append({
            'name': carry_expert['name'],
            'unit': carry_expert['unit'],
            'course': course_text,
            'date': date_clean,
            'period': period,
        })

    return experts


def parse_expert(raw):
    lines = [l.strip() for l in raw.split('\n') if l.strip()]
    if not lines:
        return {'name': '', 'unit': ''}
    raw_name = lines[0]
    # Handle double parentheses like （（主任））
    name = raw_name
    while re.search(r'[（(][^（）()]*[）)]', name):
        name = re.sub(r'[（(][^（）()]*[）)]', '', name)
    name = name.strip()
    unit_parts = []
    for line in lines[1:]:
        digits = re.sub(r'\D', '', line)
        if len(digits) >= 11 and re.match(r'^[\d\s\-/]+$', line.replace(' ', '')):
            continue
        unit_parts.append(line)
    unit = ''.join(unit_parts)
    return {'name': name, 'unit': unit}


# ========== STEP 2: Group ==========
def date_sort_key(d):
    m = re.match(r'(\d+)月\s*(\d+)日', d)
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def group_teachers(experts):
    groups = {}
    for e in experts:
        key = (e['name'], e['unit'])
        if key not in groups:
            groups[key] = {'name': e['name'], 'unit': e['unit'], 'sessions': []}
        groups[key]['sessions'].append(e)

    result = []
    for g in groups.values():
        s = sorted(g['sessions'], key=lambda x: date_sort_key(x['date']))
        # Build date string
        date_periods = {}
        for sess in s:
            d = sess['date']
            if d not in date_periods:
                date_periods[d] = set()
            date_periods[d].add(sess['period'])

        sorted_dates = sorted(date_periods.keys(), key=date_sort_key)
        parts = []
        for d in sorted_dates:
            ps = date_periods[d]
            if '上午' in ps and '下午' in ps:
                parts.append(f"{d}（全天）")
            elif '上午' in ps:
                parts.append(f"{d}（上午）")
            elif '下午' in ps:
                parts.append(f"{d}（下午）")
            else:
                parts.append(d)
        date_str = '、'.join(parts)

        # Dedup courses
        seen = set()
        courses = []
        for sess in s:
            c = sess['course'].strip()
            if c and c not in seen:
                seen.add(c)
                courses.append(c)
        course_text = '\n'.join(courses)

        result.append({
            'name': g['name'],
            'unit': g['unit'],
            'dates': date_str,
            'courses': course_text,
        })
    return result


# ========== STEP 3: Template XML manipulation ==========
def load_template_xml(path):
    """Load template docx and return (zip, files dict, document.xml string)."""
    files = {}
    with zipfile.ZipFile(path, 'r') as z:
        for name in z.namelist():
            files[name] = z.read(name)
    doc_xml = files['word/document.xml'].decode('utf-8')
    return files, doc_xml


def replace_placeholders(xml, teacher, common):
    """Replace all placeholders in document.xml string."""
    subs = {
        '（会议名字）': common['meeting_name'],
        '（对方单位）': teacher['unit'],
        '（起止时间）': common['start_date'],
        '（结束时间）': common['end_date'],
        '（地点）': common['location'],
        '（老师名字）': teacher['name'],
        '（上课时间）': teacher['dates'],
        '（课程内容）': teacher['courses'],
        '（授课地址）': common['address'],
        '（发函时间）': common['letter_date'],
    }
    for ph, val in subs.items():
        xml = xml.replace(ph, escape_xml(val))
    return xml


def escape_xml(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def build_final_docx(files_orig, doc_xml, teachers, common):
    """Concatenate all teachers' document bodies with page breaks."""
    # Extract body from template
    body_match = re.search(r'<w:body[^>]*>([\s\S]*)</w:body>', doc_xml)
    if not body_match:
        raise ValueError("No body found")
    template_body = body_match.group(1)

    all_body_parts = []
    for i, teacher in enumerate(teachers):
        xml = replace_placeholders(template_body, teacher, common)
        # Remove sectPr from intermediate bodies (keep only last)
        if i < len(teachers) - 1:
            xml = re.sub(r'<w:sectPr[\s\S]*?/w:sectPr>', '', xml)
            # Add page break at end
            xml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        all_body_parts.append(xml)

    final_body = ''.join(all_body_parts)
    final_xml = doc_xml.replace(template_body, final_body)

    files = copy.copy(files_orig)
    files['word/document.xml'] = final_xml.encode('utf-8')

    with zipfile.ZipFile(OUTPUT_PATH, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f"Saved: {OUTPUT_PATH}")


# ========== MAIN ==========
def main():
    print("=" * 50)
    print(f"培训: {MEETING_NAME}")
    print(f"时间: {START_DATE} ~ {END_DATE}")
    print(f"地点: {ADDRESS}")
    print("=" * 50)

    # Step 1
    print("\n[1/3] Parsing schedule...")
    experts = parse_schedule(SCHEDULE_PATH)
    print(f"  → {len(experts)} session entries")

    # Step 2
    print("\n[2/3] Grouping by teacher...")
    teachers = group_teachers(experts)
    print(f"  → {len(teachers)} teachers:")
    for t in teachers:
        print(f"    · {t['name']:　<6} | {t['unit']}")

    # Step 3
    print("\n[3/3] Generating 聘书...")
    common = {
        'meeting_name': MEETING_NAME,
        'start_date': START_DATE,
        'end_date': END_DATE,
        'location': LOCATION,
        'address': ADDRESS,
        'letter_date': LETTER_DATE,
    }

    files, doc_xml = load_template_xml(TEMPLATE_PATH)
    build_final_docx(files, doc_xml, teachers, common)

    print(f"\n✅ Done! {len(teachers)} 份聘书 → {OUTPUT_PATH}")


if __name__ == '__main__':
    main()
