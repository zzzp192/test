import originpro as op
import os
import sys
import pandas as pd
import re
from pptx import Presentation
from pptx.util import Inches
import time

def find_data_sheet(file_path):
    """
    自动识别数据所在的工作表
    - 包含"曲线数据"或"应变"/"应力"的为绘图表
    - 包含"试样编号"的为汇总表（提取试样编号列表）
    返回: (sheet_name, sample_ids) 
    """
    if file_path.endswith('.csv'):
        return None, None
    
    xls = pd.ExcelFile(file_path)
    data_sheet = None
    sample_ids = []
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # 优先识别曲线数据表（名称包含"曲线"或列名包含"应变"/"应力"）
        if '曲线' in sheet_name:
            data_sheet = sheet_name
            continue
        
        # 检查列名是否包含应变/应力
        col_str = ' '.join(str(c) for c in df.columns)
        if '应变' in col_str or '应力' in col_str:
            data_sheet = sheet_name
            continue
        
        # 检查是否包含"试样编号"列（汇总表，提取试样编号列表）
        for col in df.columns:
            if '试样编号' in str(col):
                sample_ids = [str(v) for v in df[col] if pd.notna(v)]
                break
    
    return data_sheet, sample_ids

def plot_in_origin(file_path, template_path=None, lines_per_graph=1, swap_xy=False):
    """
    在 Origin 中绘图的核心函数
    """
    print(f"--- 正在尝试连接 Origin... ---")

    # --- 1. 连接 Origin ---
    try:
        if op.oext:
            op.set_show(True)
    except Exception as e:
        print(f"连接警告 (可忽略): {e}")

    # --- 2. 创建工作簿 ---
    try:
        wb = op.new_book() 
        if not wb: raise Exception("创建工作簿失败")
        wks = wb[0] 
    except Exception as e:
        return f"Origin 初始化失败: {str(e)}\n请确保 Origin 已管理员启动且执行过 doc -s"

    print(f"导入数据: {os.path.basename(file_path)}")
    
    # --- 3. 自动识别工作表 ---
    data_sheet, sample_ids = find_data_sheet(file_path)
    
    # --- 4. 导入数据 ---
    try:
        if data_sheet:
            print(f"识别到数据表: {data_sheet}")
            df = pd.read_excel(file_path, sheet_name=data_sheet)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        # 如果需要交换XY列（在内存中，不修改原文件）
        if swap_xy:
            cols = list(df.columns)
            new_cols = []
            for i in range(0, len(cols) - 1, 2):
                new_cols.append(cols[i + 1])
                new_cols.append(cols[i])
            if len(cols) % 2 == 1:
                new_cols.append(cols[-1])
            df = df[new_cols]
            print(f"已交换XY列顺序")
        
        # 将试样编号设置为Y列的列名
        if sample_ids:
            cols = list(df.columns)
            for i, sid in enumerate(sample_ids):
                y_col_idx = i * 2 + 1  # Y列索引: 1, 3, 5, 7...
                if y_col_idx < len(cols):
                    cols[y_col_idx] = sid
            df.columns = cols
            print(f"已将试样编号设置为Y列名: {sample_ids[:3]}...")
        
        wks.from_df(df)
        print(f"数据导入成功，列数: {wks.cols}")
    except Exception as e:
        return f"导入数据失败: {str(e)}"

    # --- 4. 获取列数 ---
    num_cols = wks.cols
    if num_cols < 2:
        return "错误：数据列不足，至少需要 2 列 (1组 XY)"

    # --- 5. 分组绘图逻辑 (适配 XYXY) ---
    # 我们只关注 Y 列的索引：1, 3, 5, 7...
    y_cols = list(range(1, num_cols, 2))
    
    if not y_cols:
        return "警告：未找到有效的 Y 列，请检查数据列数是否为偶数。"

    chunks = [y_cols[i:i + lines_per_graph] for i in range(0, len(y_cols), lines_per_graph)]
    
    created_graphs = []

    for i, chunk in enumerate(chunks):
        template = template_path if template_path and os.path.exists(template_path) else None
        
        try:
            # 创建图形
            if template:
                graph = op.new_graph(template=template)
            else:
                graph = op.new_graph()
            layer = graph[0] 
            
            for y_idx in chunk:
                # 对应的 X 轴就在它前一列
                x_idx = y_idx - 1 
                
                # add_plot 使用整数索引
                layer.add_plot(wks, coly=y_idx, colx=x_idx, type='line')
            
            layer.rescale()
            
            fname = os.path.splitext(os.path.basename(file_path))[0]
            graph.name = f"{fname}_G{i+1}"
            created_graphs.append(graph.name)
            
        except Exception as e:
            print(f"绘图组 {i+1} 失败: {e}")

    return f"成功！\n数据模式: XYXY\n已创建 {len(created_graphs)} 张图表。\n请切换到 Origin 查看。"


def get_sample_ids_from_excel(file_path, data_type='tensile'):
    """从Excel提取试样编号列表"""
    xls = pd.ExcelFile(file_path)
    sample_ids = []
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        for col in df.columns:
            if '试样编号' in str(col):
                sample_ids = [str(v) for v in df[col] if pd.notna(v) and str(v).strip()]
                if sample_ids:
                    return sample_ids
    return sample_ids


def export_graph_to_image(graph, output_path):
    """导出Origin图形为EMF矢量图"""
    try:
        # 优先导出EMF格式（矢量图，可在PPT中编辑）
        emf_path = output_path.replace('.png', '.emf')
        graph.save_fig(emf_path)
        return emf_path if os.path.exists(emf_path) else None
    except:
        # 备选PNG
        try:
            graph.save_fig(output_path, width=800)
            return output_path if os.path.exists(output_path) else None
        except:
            return None


def create_ppt_with_origin_graphs(graph_names, output_ppt_path, folder=None):
    """导出Origin图形为PNG并创建PPT"""
    if folder is None:
        folder = os.path.dirname(output_ppt_path)
    
    image_paths = []
    for i, gname in enumerate(graph_names):
        img_path = os.path.join(folder, f"temp_graph_{i}.png")
        emf_path = os.path.join(folder, f"temp_graph_{i}.emf")
        try:
            graph = op.find_graph(gname)
            if graph:
                print(f"正在导出图形 {i+1}/{len(graph_names)}: {gname}")
                graph.activate()
                time.sleep(0.3)
                
                # 方法1: 直接使用save_fig导出EMF
                try:
                    graph.save_fig(emf_path)
                    time.sleep(0.3)
                    if os.path.exists(emf_path):
                        image_paths.append(emf_path)
                        print(f"EMF导出成功: {emf_path}")
                        continue
                except Exception as e1:
                    print(f"EMF导出失败: {e1}")
                
                # 方法2: 使用save_fig导出PNG
                try:
                    graph.save_fig(img_path, width=1200)
                    time.sleep(0.3)
                    if os.path.exists(img_path):
                        image_paths.append(img_path)
                        print(f"PNG导出成功: {img_path}")
                        continue
                except Exception as e2:
                    print(f"PNG导出失败: {e2}")
                
                # 方法3: LabTalk expGraph
                try:
                    lt_path = img_path.replace('\\', '/')
                    op.lt_exec(f'expGraph type:=png tr.unit:=2 tr.width:=1200 path:="{lt_path}";')
                    time.sleep(0.5)
                    if os.path.exists(img_path):
                        image_paths.append(img_path)
                        print(f"LabTalk导出成功: {img_path}")
                except Exception as e3:
                    print(f"LabTalk导出失败: {e3}")
                    
        except Exception as e:
            print(f"导出图形 {gname} 失败: {e}")
    
    print(f"共导出 {len(image_paths)} 张图片")
    
    if image_paths:
        create_ppt_from_images(image_paths, output_ppt_path)
        # 清理临时文件
        for p in image_paths:
            try: os.remove(p)
            except: pass
        return output_ppt_path
    else:
        print("警告: 没有成功导出任何图片，PPT将为空")
    return None


def create_ppt_from_images(image_paths, output_ppt_path, origin_project_path=None):
    """将图片列表创建为PPT，每页一张图（备用方案）"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # 使用最后一个布局（通常是空白）
    blank_layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
    
    for img_path in image_paths:
        if os.path.exists(img_path):
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(img_path, Inches(0.5), Inches(0.5), 
                                    width=Inches(12.333), height=Inches(6.5))
    
    prs.save(output_ppt_path)
    return output_ppt_path


def save_origin_project(output_path):
    """保存当前Origin项目"""
    try:
        op.save(output_path)
        return True
    except:
        return False


def append_origin_graphs_to_ppt(graph_names, ppt_path, folder=None):
    """将Origin图形作为EMF图片添加到已有PPT文件末尾"""
    if folder is None:
        folder = os.path.dirname(ppt_path)
    
    try:
        prs = Presentation(ppt_path)
        blank_layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
        
        for i, gname in enumerate(graph_names):
            img_path = os.path.join(folder, f"temp_append_{i}.emf")
            try:
                graph = op.find_graph(gname)
                if graph:
                    graph.save_fig(img_path)
                    if os.path.exists(img_path):
                        slide = prs.slides.add_slide(blank_layout)
                        slide.shapes.add_picture(img_path, Inches(0.5), Inches(0.5), 
                                                width=Inches(12.333), height=Inches(6.5))
                        os.remove(img_path)
            except Exception as e:
                print(f"添加图形 {gname} 失败: {e}")
        
        prs.save(ppt_path)
        return ppt_path
    except Exception as e:
        print(f"添加图形到PPT失败: {e}")
        return None


def get_tensile_sample_ids(file_path):
    """从拉伸报告Excel提取试样编号列表（处理特殊格式）"""
    xls = pd.ExcelFile(file_path)
    sample_ids = []
    
    for sheet_name in xls.sheet_names:
        # 读取整个sheet，不设置header
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # 遍历每一行，找到包含"试样编号"的行
        for row_idx in range(min(10, len(df_raw))):  # 只检查前10行
            row_values = [str(v) for v in df_raw.iloc[row_idx] if pd.notna(v)]
            if any('试样编号' in v for v in row_values):
                # 找到试样编号所在的列
                for col_idx, val in enumerate(df_raw.iloc[row_idx]):
                    if pd.notna(val) and '试样编号' in str(val):
                        # 提取该列下面的所有非空值
                        sample_ids = [str(v) for v in df_raw.iloc[row_idx+1:, col_idx] 
                                     if pd.notna(v) and str(v).strip()]
                        if sample_ids:
                            print(f"从第{row_idx+1}行找到试样编号列，提取到{len(sample_ids)}个编号")
                            return sample_ids
    return sample_ids


def plot_tensile_to_ppt(file_path, template_path=None, lines_per_graph=12, swap_xy=True, append_to_ppt=None):
    """拉伸报告Origin绘图并导出PPT
    原始数据格式: 应力01(Y), 应变01(X), 应力02(Y), 应变02(X)...
    目标XYXY格式: 应变01(X), 试样编号1(Y), 应变02(X), 试样编号2(Y)...
    """
    try:
        if op.oext:
            op.set_show(True)
    except: pass
    
    # 读取数据 - 使用专门的拉伸报告试样编号提取函数
    sample_ids = get_tensile_sample_ids(file_path)
    print(f"提取到试样编号: {sample_ids}")
    
    # 找曲线数据表
    xls = pd.ExcelFile(file_path)
    df = None
    for sheet in xls.sheet_names:
        if '曲线' in sheet:
            df = pd.read_excel(xls, sheet_name=sheet)
            break
    if df is None:
        df = pd.read_excel(file_path)
    
    # 原始列: 应力01, 应变01, 应力02, 应变02...
    # 目标: 应变01(X), 试样编号1(Y), 应变02(X), 试样编号2(Y)...
    cols = list(df.columns)
    new_cols = []
    sample_idx = 0
    
    # 每两列一组：应力(Y), 应变(X) -> 应变(X), 试样编号(Y)
    for i in range(0, len(cols) - 1, 2):
        y_col = cols[i]      # 应力列 -> Y
        x_col = cols[i + 1]  # 应变列 -> X
        
        # 交换顺序：X在前，Y在后
        new_cols.append(x_col)  # X列保持原名（应变）
        if sample_idx < len(sample_ids):
            new_cols.append(sample_ids[sample_idx])  # Y列用试样编号
            print(f"替换列名: {y_col} -> {sample_ids[sample_idx]}")
            sample_idx += 1
        else:
            new_cols.append(y_col)
    
    if len(cols) % 2 == 1:
        new_cols.append(cols[-1])
    
    # 重新排列数据列
    reordered_cols = []
    for i in range(0, len(cols) - 1, 2):
        reordered_cols.append(cols[i + 1])  # 应变(X)
        reordered_cols.append(cols[i])      # 应力(Y)
    if len(cols) % 2 == 1:
        reordered_cols.append(cols[-1])
    
    df = df[reordered_cols]
    df.columns = new_cols
    
    # 创建工作簿导入数据
    wb = op.new_book()
    wks = wb[0]
    wks.from_df(df)
    
    # 设置列类型：偶数列(0,2,4...)为X，奇数列(1,3,5...)为Y
    num_cols = wks.cols
    wks_name = wks.get_book().name
    for col_idx in range(num_cols):
        col_type = 4 if col_idx % 2 == 0 else 1  # 4=X, 1=Y
        op.lt_exec(f'wks.col{col_idx + 1}.type = {col_type};')
    
    print(f"已设置XYXY列类型，共{num_cols}列")
    
    # 分组绘图
    y_cols = list(range(1, num_cols, 2))
    chunks = [y_cols[i:i + lines_per_graph] for i in range(0, len(y_cols), lines_per_graph)]
    
    graph_names = []
    folder = os.path.dirname(file_path)
    fname = os.path.splitext(os.path.basename(file_path))[0]
    
    for i, chunk in enumerate(chunks):
        graph = op.new_graph(template=template_path) if template_path else op.new_graph()
        layer = graph[0]
        for y_idx in chunk:
            layer.add_plot(wks, coly=y_idx, colx=y_idx - 1, type='line')
        layer.rescale()
        gname = f"{fname}_T{i+1}"
        graph.name = gname
        graph_names.append(gname)
    
    # 保存Origin项目文件
    opju_path = os.path.join(folder, f"拉伸曲线_{fname}.opju")
    save_origin_project(opju_path)
    
    # 如果指定了append_to_ppt，则添加到已有PPT
    if append_to_ppt and os.path.exists(append_to_ppt):
        result = append_origin_graphs_to_ppt(graph_names, append_to_ppt)
        if result:
            return f"成功！已添加 {len(chunks)} 张图表到报告\nOrigin项目: {opju_path}"
    
    # 否则创建新PPT
    ppt_path = os.path.join(folder, f"拉伸曲线_{fname}.pptx")
    result = create_ppt_with_origin_graphs(graph_names, ppt_path)
    
    if not result:
        image_paths = []
        for i, gname in enumerate(graph_names):
            img_path = os.path.join(folder, f"temp_{i}.emf")
            try:
                op.find_graph(gname).save_fig(img_path)
                image_paths.append(img_path)
            except: pass
        create_ppt_from_images(image_paths, ppt_path)
        for p in image_paths:
            try: os.remove(p)
            except: pass
    
    return f"成功！已创建 {len(chunks)} 张图表\nPPT: {ppt_path}\nOrigin项目: {opju_path}"


def plot_vda_to_ppt(file_path, template_path=None, lines_per_graph=12, swap_xy=True):
    """VDA报告Origin绘图并导出PPT
    原始数据格式: 力_1(Y), 位移_1(X), 力_2(Y), 位移_2(X)...
    目标XYXY格式: 位移_1(X), 试样编号1(Y), 位移_2(X), 试样编号2(Y)...
    """
    try:
        if op.oext:
            op.set_show(True)
    except: pass
    
    # 读取数据
    sample_ids = get_sample_ids_from_excel(file_path)
    print(f"VDA提取到试样编号: {sample_ids}")
    
    # 找VDA原始数据表
    xls = pd.ExcelFile(file_path)
    df = None
    for sheet in xls.sheet_names:
        if '原始数据' in sheet or 'VDA' in sheet:
            df = pd.read_excel(xls, sheet_name=sheet)
            if '力_1' in str(df.columns) or '力_' in ' '.join(str(c) for c in df.columns):
                break
    if df is None:
        df = pd.read_excel(file_path)
    
    # 原始列: 力_1(Y), 位移_1(X), 力_2(Y), 位移_2(X)...
    # 目标: 位移_1(X), 试样编号1(Y), 位移_2(X), 试样编号2(Y)...
    cols = list(df.columns)
    new_cols = []
    sample_idx = 0
    
    # 每两列一组：力(Y), 位移(X) -> 位移(X), 试样编号(Y)
    for i in range(0, len(cols) - 1, 2):
        y_col = cols[i]      # 力列 -> Y
        x_col = cols[i + 1]  # 位移列 -> X
        
        new_cols.append(x_col)  # X列保持原名（位移）
        if sample_idx < len(sample_ids):
            new_cols.append(sample_ids[sample_idx])  # Y列用试样编号
            print(f"VDA替换列名: {y_col} -> {sample_ids[sample_idx]}")
            sample_idx += 1
        else:
            new_cols.append(y_col)
    
    if len(cols) % 2 == 1:
        new_cols.append(cols[-1])
    
    # 重新排列数据列
    reordered_cols = []
    for i in range(0, len(cols) - 1, 2):
        reordered_cols.append(cols[i + 1])  # 位移(X)
        reordered_cols.append(cols[i])      # 力(Y)
    if len(cols) % 2 == 1:
        reordered_cols.append(cols[-1])
    
    df = df[reordered_cols]
    df.columns = new_cols
    
    # 创建工作簿
    wb = op.new_book()
    wks = wb[0]
    wks.from_df(df)
    
    # 设置列类型：偶数列(0,2,4...)为X，奇数列(1,3,5...)为Y
    num_cols = wks.cols
    for col_idx in range(num_cols):
        col_type = 4 if col_idx % 2 == 0 else 1  # 4=X, 1=Y
        op.lt_exec(f'wks.col{col_idx + 1}.type = {col_type};')
    
    print(f"VDA已设置XYXY列类型，共{num_cols}列")
    
    # 分组绘图
    y_cols = list(range(1, num_cols, 2))
    chunks = [y_cols[i:i + lines_per_graph] for i in range(0, len(y_cols), lines_per_graph)]
    
    graph_names = []
    folder = os.path.dirname(file_path)
    fname = os.path.splitext(os.path.basename(file_path))[0]
    
    for i, chunk in enumerate(chunks):
        graph = op.new_graph(template=template_path) if template_path else op.new_graph()
        layer = graph[0]
        
        for y_idx in chunk:
            x_idx = y_idx - 1
            layer.add_plot(wks, coly=y_idx, colx=x_idx, type='line')
        
        layer.rescale()
        gname = f"{fname}_V{i+1}"
        graph.name = gname
        graph_names.append(gname)
    
    # 保存Origin项目文件
    opju_path = os.path.join(folder, f"VDA曲线_{fname}.opju")
    save_origin_project(opju_path)
    
    # 创建PPT
    ppt_path = os.path.join(folder, f"VDA曲线_{fname}.pptx")
    result = create_ppt_with_origin_graphs(graph_names, ppt_path)
    
    if not result:
        image_paths = []
        for i, gname in enumerate(graph_names):
            img_path = os.path.join(folder, f"temp_{i}.emf")
            try:
                op.find_graph(gname).save_fig(img_path)
                image_paths.append(img_path)
            except: pass
        create_ppt_from_images(image_paths, ppt_path)
        for p in image_paths:
            try: os.remove(p)
            except: pass
    
    return f"成功！已创建 {len(chunks)} 张图表\nPPT: {ppt_path}\nOrigin项目: {opju_path}\n(PPT中的图形可双击编辑)"


def plot_phase_change(file_paths, template_path=None):
    """相变点绘图：CSV文件，第3列X，第4列Y，每个文件一张图"""
    try:
        if op.oext:
            op.set_show(True)
    except: pass
    
    if isinstance(file_paths, str):
        file_paths = [file_paths]
    
    graph_names = []
    folder = os.path.dirname(file_paths[0]) if file_paths else "."
    
    for file_path in file_paths:
        # 读取CSV，分号分隔
        try:
            df = pd.read_csv(file_path, sep=';', encoding='utf-8', skiprows=5)
        except:
            try:
                df = pd.read_csv(file_path, sep=';', encoding='gbk', skiprows=5)
            except:
                df = pd.read_csv(file_path, skiprows=5)
        
        if df.shape[1] < 4:
            continue
        
        # 第3列X(index 2)，第4列Y(index 3)
        x_data = pd.to_numeric(df.iloc[:, 2].astype(str).str.replace(',', '.'), errors='coerce')
        y_data = pd.to_numeric(df.iloc[:, 3].astype(str).str.replace(',', '.'), errors='coerce')
        
        plot_df = pd.DataFrame({'Temperature': x_data, 'Change': y_data}).dropna()
        
        wb = op.new_book()
        wks = wb[0]
        wks.from_df(plot_df)
        
        graph = op.new_graph(template=template_path) if template_path else op.new_graph()
        layer = graph[0]
        layer.add_plot(wks, coly=1, colx=0, type='line')
        layer.rescale()
        
        fname = os.path.splitext(os.path.basename(file_path))[0]
        graph.name = fname
        graph_names.append(fname)
    
    # 保存Origin项目文件
    opju_path = os.path.join(folder, f"相变点曲线.opju")
    save_origin_project(opju_path)
    
    # 创建PPT
    ppt_path = os.path.join(folder, f"相变点曲线.pptx")
    result = create_ppt_with_origin_graphs(graph_names, ppt_path)
    
    if not result:
        image_paths = []
        for gname in graph_names:
            img_path = os.path.join(folder, f"temp_{gname}.emf")
            try:
                op.find_graph(gname).save_fig(img_path)
                image_paths.append(img_path)
            except: pass
        create_ppt_from_images(image_paths, ppt_path)
        for p in image_paths:
            try: os.remove(p)
            except: pass
    
    return f"成功！已处理 {len(file_paths)} 个文件\nPPT: {ppt_path}\nOrigin项目: {opju_path}\n(PPT中的图形可双击编辑)"
