#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.7 - VDA弯曲数据处理模块

软件名称：育材堂报告助手
版本号：V3.7
开发单位：育材堂
开发完成日期：2024年

模块功能：
    提供VDA弯曲试验数据的提取、统计计算和PPT报告生成功能。

主要功能：
    - 从Excel/CSV文件提取VDA弯曲试验数据
    - 自动识别试样编号和分组
    - 单位自动转换（N→kN）
    - 计算平均值和标准差统计
    - 动态生成PPT报告表格

数据提取字段：
    - 试样编号、公称厚度、最大力
    - 压头位移、弯曲角度

Copyright (c) 2024 育材堂. All rights reserved.
"""

import pandas as pd
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from ppt_utils import (
    to_float, delete_table_column, delete_table_row, 
    duplicate_slide, THEME_COLOR
)

def process_vda_report(excel_path, ppt_template, output_path, force_unit='kN', include_disp=True):
    """
    处理VDA弯曲数据并生成PPT报告
    简化版：默认单位 kN，不再尝试修改表头文字
    """
    print(f"--- 开始处理 VDA 弯曲报告: {excel_path} (单位: {force_unit}) ---")
    
    # 1. 读取数据
    try:
        try:
            df = pd.read_csv(excel_path, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(excel_path, encoding='gbk')
            except:
                df = pd.read_csv(excel_path, encoding='gb18030')
    except:
        try:
            df = pd.read_excel(excel_path, sheet_name="2. VDA弯曲")
        except:
            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                return f"读取Excel/CSV失败: {str(e)}"

    # 2. 列名映射
    col_map = {
        "试样编号": "SampleID",
        "公称厚度t0": "Thickness",
        "最大力Fm": "MaxForce",
        "压头位移S": "Displacement",
        "角度": "Angle"
    }
    
    for col, std_name in col_map.items():
        if col in df.columns:
            df.rename(columns={col: std_name}, inplace=True)
        else:
            for c in df.columns:
                if col in c:
                    df.rename(columns={c: std_name}, inplace=True)
                    break
    
    if 'SampleID' in df.columns:
        df = df.dropna(subset=['SampleID'])
        df = df[df['SampleID'].astype(str).str.strip() != '']

    required = ["SampleID", "Thickness", "MaxForce", "Displacement", "Angle"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return f"错误: Excel中找不到这些列: {missing}，请检查表头。"

    # 3. 分组逻辑
    project_id = os.path.splitext(os.path.basename(excel_path))[0]
    
    def parse_group(sid):
        sid = str(sid).strip()
        if '-' in sid:
            parts = sid.rsplit('-', 1)
            return parts[0], parts[1]
        else:
            return sid, "1"

    df['GroupName'] = df['SampleID'].apply(lambda x: parse_group(x)[0])
    df['Number'] = df['SampleID'].apply(lambda x: parse_group(x)[1])

    # 4. 准备PPT页面
    prs = Presentation(ppt_template)
    unique_groups = df['GroupName'].unique()
    total_groups = len(unique_groups)
    
    if total_groups == 0:
        return "错误：未识别到任何有效的分组数据，请检查“试样编号”列。"

    groups_per_slide = 4
    num_slides_needed = (total_groups + groups_per_slide - 1) // groups_per_slide
    
    if num_slides_needed > 1:
        for _ in range(num_slides_needed - 1):
            duplicate_slide(prs, 0)
    
    group_chunks = [unique_groups[i:i + groups_per_slide] for i in range(0, total_groups, groups_per_slide)]

    # 5. 循环填充每一页
    for slide_idx, chunk_groups in enumerate(group_chunks):
        if slide_idx >= len(prs.slides): break
            
        slide = prs.slides[slide_idx]
        replace_text_in_slide(slide, "项目号", project_id)
        
        tables = [s.table for s in slide.shapes if s.has_table]
        if not tables: continue
        main_table = tables[0]

        # 根据选项删除“压头位移”列
        if not include_disp:
            if len(main_table.columns) >= 5:
                delete_table_column(main_table, 4)
        
        # 填充数据
        process_table_chunk(main_table, chunk_groups, df, force_unit, include_disp)

    # 6. 保存
    try:
        prs.save(output_path)
        return f"成功生成报告！\n共 {total_groups} 组数据，{num_slides_needed} 页。\n已保存至: {output_path}"
    except Exception as e:
        return f"保存PPT失败 (请关闭已打开的同名PPT): {str(e)}"

# ================= 辅助函数 =================

def process_table_chunk(table, groups, full_df, unit, include_disp):
    HEADER_ROWS = 1
    DEFAULT_DATA_ROWS = 3
    BLOCK_SIZE = DEFAULT_DATA_ROWS + 1 
    row_offset = 0 
    
    for i in range(4):
        base_start_row = HEADER_ROWS + (i * BLOCK_SIZE)
        current_start_row = base_start_row + row_offset
        
        if i < len(groups):
            group_name = groups[i]
            group_data = full_df[full_df['GroupName'] == group_name].copy()
            
            try:
                group_data['Number'] = group_data['Number'].astype(int)
                group_data.sort_values('Number', inplace=True)
            except: pass
            
            n_data = len(group_data)
            diff = n_data - DEFAULT_DATA_ROWS
            current_stats_row_idx = current_start_row + DEFAULT_DATA_ROWS
            
            if diff > 0:
                for _ in range(diff):
                    add_table_row(table, current_stats_row_idx - 1)
                    row_offset += 1
                    current_stats_row_idx += 1
            elif diff < 0:
                rows_to_del = abs(diff)
                for _ in range(rows_to_del):
                    delete_table_row(table, current_stats_row_idx - 1)
                    current_stats_row_idx -= 1
                    row_offset -= 1
            
            fill_group_data(table, group_data, group_name, current_start_row, n_data, unit, include_disp)
            stats_idx = current_start_row + n_data
            fill_stats_row(table, group_data, stats_idx, unit, include_disp)
            
        else:
            if current_start_row >= len(table.rows): continue
            try:
                for _ in range(BLOCK_SIZE):
                    if current_start_row < len(table.rows):
                        delete_table_row(table, current_start_row)
                        row_offset -= 1
            except: pass

def fill_group_data(table, data, group_name, start_row, n_rows, unit, include_disp):
    if n_rows > 1:
        try:
            cell_top = table.cell(start_row, 0)
            cell_bottom = table.cell(start_row + n_rows - 1, 0)
            cell_top.merge(cell_bottom)
        except: pass
    
    cell_name = table.cell(start_row, 0)
    cell_name.text = str(group_name)
    format_cell(cell_name, 12)
    
    for idx, (_, row) in enumerate(data.iterrows()):
        r = start_row + idx
        format_cell_text(table, r, 1, str(row['Number']))
        format_cell_text(table, r, 2, f"{to_float(row['Thickness']):.2f}")
        
        # 注意：这里保留了单位转换逻辑
        val_f = to_float(row['MaxForce'])
        if unit == 'kN': val_f /= 1000.0
        format_cell_text(table, r, 3, f"{val_f:.1f}")
        
        current_col = 4
        if include_disp:
            format_cell_text(table, r, current_col, f"{to_float(row['Displacement']):.2f}")
            current_col += 1
        
        if current_col < len(table.columns):
            format_cell_text(table, r, current_col, f"{to_float(row['Angle']):.2f}")

def fill_stats_row(table, data, r_idx, unit, include_disp):
    c_lbl = table.cell(r_idx, 1)
    c_lbl.text = "平均值±标准差"
    format_cell(c_lbl, 12, bold=True)
    custom_color = (25, 137, 141)
    
    set_stat_cell(table.cell(r_idx, 2), data['Thickness'], 1, 1.0, custom_color)
    
    # 统计行单位转换
    f_vals = pd.to_numeric(data['MaxForce'], errors='coerce')
    if unit == 'kN': f_vals = f_vals / 1000.0
    set_stat_cell(table.cell(r_idx, 3), f_vals, 1, 1.0, custom_color)
    
    current_col = 4
    if include_disp:
        set_stat_cell(table.cell(r_idx, current_col), data['Displacement'], 1, 1.0, custom_color)
        current_col += 1
        
    if current_col < len(table.columns):
        set_stat_cell(table.cell(r_idx, current_col), data['Angle'], 1, 1.0, custom_color)

def format_cell_text(table, r, c, text):
    try:
        cell = table.cell(r, c)
        cell.text = text
        format_cell(cell, 12)
    except: pass

def set_stat_cell(cell, series, decimals=1, factor=1.0, color_rgb=None):
    series = pd.to_numeric(series, errors='coerce').dropna() * factor
    if len(series) == 0:
        txt = "-"
    else:
        mean = series.mean()
        std = series.std(ddof=1) if len(series) > 1 else 0.0
        txt = f"{mean:.{decimals}f}±{std:.{decimals}f}"
    
    cell.text = txt
    format_cell(cell, 12, bold=True, color=color_rgb)

def format_cell(cell, size, bold=False, color=None):
    if not cell.text_frame.paragraphs:
        cell.text_frame.text = ""
    p = cell.text_frame.paragraphs[0]
    p.font.size = Pt(size)
    p.font.name = '微软雅黑'
    p.font.bold = bold
    if color:
        p.font.color.rgb = RGBColor(*color)
    p.alignment = 2

def replace_text_in_slide(slide, old_txt, new_txt):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                if old_txt in p.text:
                    p.text = p.text.replace(old_txt, new_txt)

def add_table_row(table, clone_idx):
    import copy
    from pptx.oxml.ns import qn
    tr = table.rows[clone_idx]._tr
    new_tr = copy.deepcopy(tr)
    for tc in new_tr.tc_lst:
        if tc.tcPr is not None:
            for tag in ["a:vMerge", "a:gridSpan"]:
                elem = tc.tcPr.find(qn(tag))
                if elem is not None:
                    tc.tcPr.remove(elem)
    tr.addnext(new_tr)
