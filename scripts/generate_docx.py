"""将论文全文.md转换为DOCX格式，适配《矿业研究与开发》投稿要求"""

import re
import os
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = __import__('lxml.etree', fromlist=['etree']).etree.SubElement(tcPr, qn('w:tcBorders'))
    for edge, val in kwargs.items():
        element = tcBorders.find(qn(f'w:{edge}'))
        if element is None:
            element = __import__('lxml.etree', fromlist=['etree']).etree.SubElement(tcBorders, qn(f'w:{edge}'))
        element.set(qn('w:val'), val.get('val', 'single'))
        element.set(qn('w:sz'), val.get('sz', '4'))
        element.set(qn('w:color'), val.get('color', '000000'))
        element.set(qn('w:space'), val.get('space', '0'))


def create_document():
    doc = Document()

    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 设置页边距
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)

    return doc


def add_heading(doc, text, level):
    """添加标题"""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        if level == 0:
            run.font.size = Pt(16)
        elif level == 1:
            run.font.size = Pt(14)
        elif level == 2:
            run.font.size = Pt(12)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER if level <= 1 else WD_ALIGN_PARAGRAPH.LEFT
    return heading


def add_paragraph_with_formatting(doc, text, bold=False, font_size=Pt(10.5), font_name='宋体', alignment=None):
    """添加带格式的段落"""
    p = doc.add_paragraph()
    if alignment:
        p.alignment = alignment

    # 处理加粗标记
    parts = re.split(r'\*\*(.*?)\*\*', text)
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        run.font.name = font_name
        run.font.size = font_size
        run.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        if i % 2 == 1 or bold:
            run.bold = True

    return p


def add_table(doc, headers, rows):
    """添加表格"""
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # 表头
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header.strip()
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.name = '宋体'
                run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 数据行
    for row_idx, row in enumerate(rows):
        for col_idx, cell_text in enumerate(row):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = cell_text.strip()
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(9)
                    run.font.name = '宋体'
                    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    return table


def convert_md_to_docx(md_path, docx_path):
    """主转换函数"""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    doc = create_document()
    lines = content.split('\n')
    i = 0
    in_table = False
    table_headers = []
    table_rows = []

    while i < len(lines):
        line = lines[i].strip()

        # 跳过空行
        if not line:
            i += 1
            continue

        # 跳过分隔线
        if line == '---':
            i += 1
            continue

        # 跳过图片标记（Markdown图片在docx中需要特殊处理）
        if line.startswith('!['):
            # 提取图片说明
            match = re.match(r'!\[(.*?)\]\((.*?)\)', line)
            if match:
                alt_text = match.group(1)
                img_path = match.group(2)
                # 尝试插入图片
                full_img_path = os.path.join(os.path.dirname(md_path), img_path)
                if os.path.exists(full_img_path):
                    try:
                        doc.add_picture(full_img_path, width=Inches(5.5))
                        last_paragraph = doc.paragraphs[-1]
                        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    except Exception:
                        p = doc.add_paragraph(f'[图片: {alt_text}]')
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p = doc.add_paragraph(f'[图片: {alt_text}]')
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            i += 1
            continue

        # 标题
        if line.startswith('# '):
            # 主标题
            text = line[2:].strip()
            add_heading(doc, text, level=0)
            i += 1
            continue

        if line.startswith('## '):
            text = line[3:].strip()
            add_heading(doc, text, level=1)
            i += 1
            continue

        if line.startswith('### '):
            text = line[4:].strip()
            add_heading(doc, text, level=2)
            i += 1
            continue

        # 表格标题（加粗）
        if line.startswith('**表') and line.endswith('**'):
            text = line.strip('*').strip()
            add_paragraph_with_formatting(doc, text, bold=True, font_size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            i += 1
            continue

        # 图片标题（加粗）
        if line.startswith('**图') and line.endswith('**'):
            text = line.strip('*').strip()
            add_paragraph_with_formatting(doc, text, bold=True, font_size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            i += 1
            continue

        # Fig./Table 英文标题
        if line.startswith('**Fig.') or line.startswith('**Table'):
            text = line.strip('*').strip()
            p = add_paragraph_with_formatting(doc, text, bold=True, font_size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            i += 1
            continue

        # LaTeX公式
        if line.startswith('$$'):
            formula_lines = [line[2:]]
            i += 1
            while i < len(lines) and not lines[i].strip().endswith('$$'):
                formula_lines.append(lines[i].strip())
                i += 1
            if i < len(lines):
                formula_lines.append(lines[i].strip()[:-2])
            formula_text = ' '.join(formula_lines).strip()
            # 提取标签
            tag_match = re.search(r'\\tag\{(\d+)\}', formula_text)
            tag = f'({tag_match.group(1)})' if tag_match else ''
            formula_clean = re.sub(r'\\tag\{\d+\}', '', formula_text).strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f'{formula_clean}  {tag}')
            run.font.name = 'Cambria Math'
            run.font.size = Pt(10)
            run.italic = True
            i += 1
            continue

        # 表格
        if '|' in line and not line.startswith('**'):
            # 收集表格数据
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if cells and all(c.replace('-', '').replace(':', '') == '' or set(c.replace('-', '').replace(':', '')).issubset({'-'}) for c in cells):
                # 分隔行，跳过
                i += 1
                continue
            if not in_table:
                in_table = True
                table_headers = cells
            else:
                table_rows.append(cells)
            # 检查下一行是否还是表格
            if i + 1 < len(lines) and '|' in lines[i + 1] and not lines[i + 1].strip().startswith('**'):
                i += 1
                continue
            else:
                # 表格结束，创建表格
                if table_headers and table_rows:
                    add_table(doc, table_headers, table_rows)
                in_table = False
                table_headers = []
                table_rows = []
                i += 1
                continue

        # 注释行
        if line.startswith('注：'):
            text = line[1:].strip()
            p = add_paragraph_with_formatting(doc, text, font_size=Pt(9))
            p.paragraph_format.first_line_indent = Cm(0.74)
            i += 1
            continue

        # 参考文献
        if line.startswith('[') and ']' in line:
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = '宋体'
            run.font.size = Pt(9)
            run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            i += 1
            continue

        # 普通段落
        # 收集连续非空行作为一个段落
        para_lines = [line]
        while i + 1 < len(lines) and lines[i + 1].strip() and not lines[i + 1].strip().startswith('#') and not lines[i + 1].strip().startswith('![') and not lines[i + 1].strip().startswith('$$') and not lines[i + 1].strip().startswith('|') and not lines[i + 1].strip().startswith('**表') and not lines[i + 1].strip().startswith('**图') and not lines[i + 1].strip().startswith('**Fig') and not lines[i + 1].strip().startswith('**Table') and lines[i + 1].strip() != '---':
            i += 1
            para_lines.append(lines[i].strip())

        full_text = ' '.join(para_lines)
        add_paragraph_with_formatting(doc, full_text)
        i += 1

    doc.save(docx_path)
    print(f"已生成: {docx_path}")


if __name__ == '__main__':
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    md_path = os.path.join(project_root, '论文全文.md')
    docx_path = os.path.join(project_root, '论文全文.docx')
    convert_md_to_docx(md_path, docx_path)
