#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.7 - 硬度数据处理模块

软件名称：育材堂报告助手
版本号：V3.7
开发单位：育材堂
开发完成日期：2024年

模块功能：
    提供PDF格式硬度报告的数据提取功能。

主要功能：
    - 从PDF文件中提取硬度测量数据表格
    - 自动识别统计表（包含Mean和SD列）
    - 提取平均值和标准差数据
    - 自动编号分组

Copyright (c) 2024 育材堂. All rights reserved.
"""

import pdfplumber

def parse_hardness_report(file_path):
    extracted_data = []
    print(f"--- 开始处理文件: {file_path} ---")
    
    # 1. 初始化自动计数器，从1开始
    auto_id_counter = 1
    
    try:
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if not tables: continue
                
                for table in tables:
                    # 数据清洗
                    clean_table = [[str(cell).strip() if cell else "" for cell in row] for row in table]
                    if len(clean_table) < 2: continue

                    # 识别表头
                    header = clean_table[0]
                    header_str = " ".join(header).replace("\n", "")
                    
                    # 检查是否是统计表 (必须包含 SD 和 平均值)
                    # 这里的匹配逻辑还是必须的，为了防止读到非统计数据的表格
                    if ("SD" in header_str or "Std" in header_str) and \
                       ("平均" in header_str or "Average" in header_str or "Mean" in header_str):
                        
                        mean_idx = -1
                        sd_idx = -1
                        
                        # 只需要找数据列，不再关心 ID 列在哪里
                        for idx, col_name in enumerate(header):
                            col = col_name.replace("\n", "").strip()
                            if "平均" in col or "Average" in col or "Mean" in col:
                                mean_idx = idx
                            if "SD" in col or "Std" in col:
                                sd_idx = idx
                        
                        # 如果找到了关键的数据列
                        if mean_idx != -1 and sd_idx != -1:
                            # 跳过表头，读取数据行
                            for row in clean_table[1:]:
                                # 确保行长度足够
                                if len(row) <= max(mean_idx, sd_idx):
                                    continue
                                
                                raw_mean = row[mean_idx]
                                raw_sd = row[sd_idx]
                                
                                # 只有当均值和SD都有数据时才处理
                                if raw_mean and raw_sd:
                                    # 尝试转数字
                                    try:
                                        num_mean = float(raw_mean)
                                        num_sd = float(raw_sd)
                                    except:
                                        num_mean = raw_mean
                                        num_sd = raw_sd

                                    # --- 核心修改：直接使用计数器作为 ID ---
                                    current_id = str(auto_id_counter)
                                    auto_id_counter += 1 
                                    # ------------------------------------

                                    extracted_data.append({
                                        "id": current_id,
                                        "mean": num_mean,
                                        "sd": num_sd
                                    })
                                    print(f"提取第 {current_id} 组: {num_mean} ± {num_sd}")

    except Exception as e:
        print(f"Error: {e}")
        return [{"error": str(e)}]
        
    return extracted_data
