"""将论文全文.md转换为DOCX格式，严格遵循《矿业研究与开发》投稿模板"""

import re
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_run_font(run, cn_font='宋体', en_font='Times New Roman', size=Pt(10.5)):
    """设置run的中英文字体"""
    run.font.name = en_font
    run.font.size = size
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = r.makeelement(qn('w:rPr'), {})
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)


def set_paragraph_format(p, line_spacing=1.5, first_indent=None, space_before=Pt(0), space_after=Pt(0)):
    """设置段落格式"""
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = space_before
    pf.space_after = space_after
    if first_indent is not None:
        pf.first_line_indent = first_indent


def add_title(doc, text):
    """添加论文标题：黑体，小二号(18pt)，居中"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, cn_font='黑体', en_font='Times New Roman', size=Pt(18))
    set_paragraph_format(p, line_spacing=1.5, space_after=Pt(6))
    return p


def add_author_info(doc, text):
    """添加作者信息：宋体，五号"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, size=Pt(10.5))
    set_paragraph_format(p, line_spacing=1.5)
    return p


def add_abstract_label(doc, label='摘要'):
    """添加摘要标签"""
    p = doc.add_paragraph()
    run = p.add_run(label)
    set_run_font(run, cn_font='黑体', en_font='Times New Roman', size=Pt(10.5))
    run.bold = True
    set_paragraph_format(p, line_spacing=1.5)
    return p


def add_abstract_text(doc, text):
    """添加摘要正文：楷体，小五号(9pt)"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, cn_font='楷体', en_font='Times New Roman', size=Pt(9))
    set_paragraph_format(p, line_spacing=1.5)
    return p


def add_keywords(doc, label, text):
    """添加关键词"""
    p = doc.add_paragraph()
    run_label = p.add_run(label)
    set_run_font(run_label, cn_font='黑体', en_font='Times New Roman', size=Pt(10.5))
    run_label.bold = True
    run_text = p.add_run(text)
    set_run_font(run_text, size=Pt(10.5))
    set_paragraph_format(p, line_spacing=1.5)
    return p


def add_heading_custom(doc, text, level):
    """添加章节标题
    level 1: 黑体，四号(14pt)，如 '0 引言', '1 相关工作'
    level 2: 黑体，小四号(12pt)，如 '1.1 井下目标检测研究进展'
    level 3: 黑体，五号(10.5pt)，如 '2.1 YOLOv11n基础架构'
    """
    size_map = {1: Pt(14), 2: Pt(12), 3: Pt(10.5)}
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, cn_font='黑体', en_font='Times New Roman', size=size_map.get(level, Pt(12)))
    run.bold = True
    space_before = Pt(12) if level == 1 else Pt(6)
    set_paragraph_format(p, line_spacing=1.5, space_before=space_before, space_after=Pt(6))
    return p


def add_body_paragraph(doc, text):
    """添加正文段落：宋体，五号(10.5pt)，首行缩进2字符，1.5倍行距"""
    p = doc.add_paragraph()

    # 处理加粗标记
    parts = re.split(r'\*\*(.*?)\*\*', text)
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        set_run_font(run, size=Pt(10.5))
        if i % 2 == 1:
            run.bold = True

    set_paragraph_format(p, line_spacing=1.5, first_indent=Cm(0.74))
    return p


def add_table_caption(doc, cn_text, en_text):
    """添加表格标题：中英双语，居中，黑体，小五号"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(cn_text)
    set_run_font(run, cn_font='黑体', en_font='Times New Roman', size=Pt(9))
    run.bold = True
    set_paragraph_format(p, line_spacing=1.5, space_before=Pt(6), space_after=Pt(3))

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(en_text)
    set_run_font(run2, en_font='Times New Roman', size=Pt(9))
    run2.bold = True
    set_paragraph_format(p2, line_spacing=1.5, space_after=Pt(3))
    return p2


def add_figure_caption(doc, cn_text, en_text):
    """添加图片标题：中英双语，居中，宋体，小五号"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(cn_text)
    set_run_font(run, cn_font='宋体', en_font='Times New Roman', size=Pt(9))
    run.bold = True
    set_paragraph_format(p, line_spacing=1.5, space_after=Pt(3))

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(en_text)
    set_run_font(run2, en_font='Times New Roman', size=Pt(9))
    run2.bold = True
    set_paragraph_format(p2, line_spacing=1.5, space_after=Pt(6))
    return p2


def add_image(doc, img_path, width=Cm(12)):
    """插入图片"""
    if os.path.exists(img_path):
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(img_path, width=width)
            return p
        except Exception:
            pass
    return None


def add_three_line_table(doc, headers, rows):
    """添加三线表"""
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 设置表格样式为三线表
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = tbl.makeelement(qn('w:tblPr'), {})
        tbl.insert(0, tblPr)

    # 表头行
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header.strip())
        set_run_font(run, size=Pt(9))
        run.bold = True

    # 数据行
    for row_idx, row in enumerate(rows):
        for col_idx, cell_text in enumerate(row):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(cell_text.strip())
            set_run_font(run, size=Pt(9))

    # 设置三线表边框（仅顶线、表头底线、表格底线）
    set_three_line_borders(table)

    return table


def set_three_line_borders(table):
    """设置三线表样式"""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = tbl.makeelement(qn('w:tblPr'), {})
        tbl.insert(0, tblPr)

    # 移除默认边框
    for borders in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(borders)

    # 添加三线表边框
    borders_elem = tblPr.makeelement(qn('w:tblBorders'), {})

    for edge in ['top', 'bottom']:
        elem = borders_elem.makeelement(qn(f'w:{edge}'), {
            qn('w:val'): 'single',
            qn('w:sz'): '12',  # 粗线
            qn('w:space'): '0',
            qn('w:color'): '000000',
        })
        borders_elem.append(elem)

    # 表头底线（中等粗细）
    insideH = borders_elem.makeelement(qn('w:insideH'), {
        qn('w:val'): 'single',
        qn('w:sz'): '6',
        qn('w:space'): '0',
        qn('w:color'): '000000',
    })
    borders_elem.append(insideH)

    # 去掉竖线
    for edge in ['left', 'right', 'insideV']:
        elem = borders_elem.makeelement(qn(f'w:{edge}'), {
            qn('w:val'): 'none',
            qn('w:sz'): '0',
            qn('w:space'): '0',
            qn('w:color'): 'auto',
        })
        borders_elem.append(elem)

    tblPr.append(borders_elem)

    # 仅对第一行设置下边框
    if len(table.rows) > 0:
        row = table.rows[0]
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = tc.makeelement(qn('w:tcPr'), {})
                tc.insert(0, tcPr)
            tcBorders = tcPr.makeelement(qn('w:tcBorders'), {})
            bottom = tcBorders.makeelement(qn('w:bottom'), {
                qn('w:val'): 'single',
                qn('w:sz'): '6',
                qn('w:space'): '0',
                qn('w:color'): '000000',
            })
            tcBorders.append(bottom)
            tcPr.append(tcBorders)


def add_reference(doc, text):
    """添加参考文献：宋体，小五号(9pt)，无首行缩进"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=Pt(9))
    set_paragraph_format(p, line_spacing=1.5, first_indent=Cm(0), space_before=Pt(0), space_after=Pt(6))
    return p


def add_formula(doc, formula_text, tag=''):
    """添加公式：居中，Cambria Math"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text = f'{formula_text}  {tag}' if tag else formula_text
    run = p.add_run(text)
    run.font.name = 'Cambria Math'
    run.font.size = Pt(10)
    run.italic = True
    set_paragraph_format(p, line_spacing=1.5, space_before=Pt(6), space_after=Pt(6))
    return p


def add_footnote(doc, text):
    """添加注释行：宋体，小五号"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=Pt(9))
    set_paragraph_format(p, line_spacing=1.5, first_indent=Cm(0.74))
    return p


def convert_md_to_docx(md_path, docx_path):
    """主转换函数"""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    doc = Document()

    # 设置默认样式
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.paragraph_format.line_spacing = 1.5

    # 设置页边距
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)

    lines = content.split('\n')
    i = 0
    in_table = False
    table_headers = []
    table_rows = []
    project_root = os.path.dirname(md_path)

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

        # 主标题
        if line.startswith('# ') and not line.startswith('## '):
            text = line[2:].strip()
            add_title(doc, text)
            i += 1
            continue

        # 二级标题（章节标题 level 1）
        if line.startswith('## ') and not line.startswith('### '):
            text = line[3:].strip()
            # 跳过 "参考文献（References）" 这样的标题，稍后处理
            if '参考文献' in text:
                add_heading_custom(doc, '参考文献', level=1)
            elif text == 'Abstract':
                add_heading_custom(doc, 'Abstract', level=1)
            else:
                add_heading_custom(doc, text, level=1)
            i += 1
            continue

        # 三级标题（节标题 level 2）
        if line.startswith('### ') and not line.startswith('#### '):
            text = line[4:].strip()
            add_heading_custom(doc, text, level=2)
            i += 1
            continue

        # 四级标题（小节标题 level 3）
        if line.startswith('#### '):
            text = line[5:].strip()
            add_heading_custom(doc, text, level=3)
            i += 1
            continue

        # 图片
        if line.startswith('!['):
            match = re.match(r'!\[(.*?)\]\((.*?)\)', line)
            if match:
                alt_text = match.group(1)
                img_path = os.path.join(project_root, match.group(2))
                add_image(doc, img_path)
            i += 1
            continue

        # 图片标题（中英双语）
        if line.startswith('**图') and line.endswith('**'):
            cn_text = line.strip('*').strip()
            # 下一行可能是英文标题
            en_text = ''
            if i + 1 < len(lines) and lines[i + 1].strip().startswith('**Fig.'):
                i += 1
                en_text = lines[i].strip().strip('*').strip()
            add_figure_caption(doc, cn_text, en_text)
            i += 1
            continue

        # 单独的Fig.标题行（如果没被上面捕获）
        if line.startswith('**Fig.') and line.endswith('**'):
            text = line.strip('*').strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_run_font(run, en_font='Times New Roman', size=Pt(9))
            run.bold = True
            set_paragraph_format(p, line_spacing=1.5, space_after=Pt(6))
            i += 1
            continue

        # 表格标题（中英双语）
        if line.startswith('**表') and line.endswith('**'):
            cn_text = line.strip('*').strip()
            en_text = ''
            if i + 1 < len(lines) and lines[i + 1].strip().startswith('**Table'):
                i += 1
                en_text = lines[i].strip().strip('*').strip()
            add_table_caption(doc, cn_text, en_text)
            i += 1
            continue

        # 单独的Table标题行
        if line.startswith('**Table') and line.endswith('**'):
            text = line.strip('*').strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_run_font(run, en_font='Times New Roman', size=Pt(9))
            run.bold = True
            set_paragraph_format(p, line_spacing=1.5, space_after=Pt(3))
            i += 1
            continue

        # LaTeX行间公式
        if line.startswith('$$'):
            formula_lines = [line[2:]]
            i += 1
            while i < len(lines) and not lines[i].strip().endswith('$$'):
                formula_lines.append(lines[i].strip())
                i += 1
            if i < len(lines):
                formula_lines.append(lines[i].strip()[:-2])
            formula_text = ' '.join(formula_lines).strip()
            tag_match = re.search(r'\\tag\{(\d+)\}', formula_text)
            tag = f'({tag_match.group(1)})' if tag_match else ''
            formula_clean = re.sub(r'\\tag\{\d+\}', '', formula_text).strip()
            add_formula(doc, formula_clean, tag)
            i += 1
            continue

        # 表格
        if '|' in line and not line.startswith('**'):
            cells = [c.strip() for c in line.split('|') if c.strip()]
            # 跳过分隔行
            if cells and all(c.replace('-', '').replace(':', '').strip() == '' for c in cells):
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
                if table_headers and table_rows:
                    add_three_line_table(doc, table_headers, table_rows)
                in_table = False
                table_headers = []
                table_rows = []
                i += 1
                continue

        # 注释行
        if line.startswith('注') and (line.startswith('注：') or line.startswith('注¹') or line.startswith('注²')):
            add_footnote(doc, line)
            i += 1
            continue

        # 参考文献
        if re.match(r'^\[\d+\]', line):
            add_reference(doc, line)
            i += 1
            continue

        # 普通正文段落
        # 收集连续非空、非特殊格式行作为一个段落
        para_lines = [line]
        while i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if not next_line:
                break
            if next_line.startswith('#') or next_line.startswith('![') or next_line.startswith('$$'):
                break
            if next_line.startswith('|') and not next_line.startswith('**'):
                break
            if next_line.startswith('**表') or next_line.startswith('**图') or next_line.startswith('**Fig') or next_line.startswith('**Table'):
                break
            if re.match(r'^\[\d+\]', next_line):
                break
            if next_line == '---':
                break
            i += 1
            para_lines.append(next_line)

        full_text = ' '.join(para_lines)
        add_body_paragraph(doc, full_text)
        i += 1

    doc.save(docx_path)
    print(f"DOCX已生成: {docx_path}")


if __name__ == '__main__':
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    md_path = os.path.join(project_root, '论文全文.md')
    docx_path = os.path.join(project_root, '论文全文.docx')
    convert_md_to_docx(md_path, docx_path)
