"""PPT操作共享工具模块"""
import copy
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# 主题颜色常量
THEME_COLOR = RGBColor(25, 137, 141)
RED_COLOR = RGBColor(255, 0, 0)

def to_float(val):
    """安全转换为浮点数"""
    try:
        return float(val) if val is not None else 0.0
    except:
        return 0.0

def delete_table_column(table, col_idx):
    """删除PPT表格中的某一列"""
    tbl = table._tbl
    tblGrid = tbl.tblGrid
    gridCols = tblGrid.findall(qn("a:gridCol"))
    if col_idx < len(gridCols):
        tblGrid.remove(gridCols[col_idx])
    
    for tr in tbl.findall(qn("a:tr")):
        tcs = tr.findall(qn("a:tc"))
        if col_idx < len(tcs):
            tr.remove(tcs[col_idx])

def delete_table_row(table, row_idx):
    """删除表格行"""
    if row_idx < 0 or row_idx >= len(table.rows):
        return
    tr = table.rows[row_idx]._tr
    tr.getparent().remove(tr)

def insert_table_row(table, target_idx, source_idx):
    """在指定位置插入新行"""
    tbl = table._tbl
    source_tr = table.rows[source_idx]._tr
    new_tr = copy.deepcopy(source_tr)
    
    for tc in new_tr.tc_lst:
        if tc.tcPr is not None:
            for tag in ["a:vMerge", "a:gridSpan"]:
                elem = tc.tcPr.find(qn(tag))
                if elem is not None:
                    tc.tcPr.remove(elem)
    
    if target_idx < len(table.rows):
        table.rows[target_idx]._tr.addprevious(new_tr)
    else:
        tbl.append(new_tr)
    
    new_row = table.rows[target_idx]
    for cell in new_row.cells:
        if cell.text_frame:
            cell.text_frame.text = ""
    return new_row

def duplicate_slide(prs, index):
    """复制幻灯片"""
    source = prs.slides[index]
    try:
        layout = source.slide_layout
    except:
        layout = prs.slide_layouts[0]
    
    dest = prs.slides.add_slide(layout)
    for shape in source.shapes:
        new_el = copy.deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return dest

def format_cell(cell, text, font_size=14, is_bold=False, color_rgb=None):
    """格式化单元格"""
    if cell is None:
        return
    tf = cell.text_frame
    tf.text = text
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        if not p.runs:
            p.add_run()
        for run in p.runs:
            run.text = text
            run.font.name = '微软雅黑'
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            run.font.color.rgb = color_rgb if color_rgb else RGBColor(0, 0, 0)

def clean_cell_merge_info(cell):
    """清理单元格合并信息"""
    try:
        tc = cell._tc
        if tc.tcPr is not None:
            for tag in ["a:vMerge", "a:gridSpan"]:
                elem = tc.tcPr.find(qn(tag))
                if elem is not None:
                    tc.tcPr.remove(elem)
    except:
        pass

def replace_text_in_slide(slide, old_txt, new_txt):
    """替换幻灯片中的文本"""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                if old_txt in p.text:
                    p.text = p.text.replace(old_txt, new_txt)
