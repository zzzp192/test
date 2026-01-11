# 试验报告助手 V2.8

一款用于材料试验数据处理和报告生成的桌面工具。

## 功能模块

### 1. 拉伸报告
- 支持 Word/Excel 数据导入
- 自动提取试样编号、厚度、Rp、Rm、Ag、A 等参数
- 自动计算平均值±标准差
- 一键生成 PPT 报告
- **Origin绘图**：支持模板选择、XY列调换、每图曲线数设置，自动将"应力01"等替换为试样编号

### 2. VDA弯曲
- 支持 Excel/CSV 数据导入
- 自动提取试样编号、厚度、最大力、压头位移、角度等参数
- 单位自动转换（N→kN）
- 一键生成 PPT 报告
- **Origin绘图**：支持模板选择、XY列调换、每图曲线数设置，自动将"力_1"等替换为试样编号

### 3. 硬度提取
- 批量处理硬度数据

### 4. 相变点绘图
- 支持拖拽多个 CSV 文件
- 使用第3列(Temperature)作为X轴，第4列(Change)作为Y轴
- 每个文件生成一张图
- 支持 Origin 模板选择
- 自动导出为 PPT（每页一张图）

## 界面特性

- 默认亮色主题，支持深色/亮色切换
- 拖拽文件支持
- 现代化 UI 设计

## 依赖

- Python 3.x
- tkinter / tkinterdnd2
- python-pptx
- pandas
- openpyxl
- python-docx
- originpro (Origin 2024b+)

## 使用方法

```bash
python main.py
```

## 文件结构

```
├── main.py              # 主程序入口
├── gui_shared.py        # 共享UI组件和主题
├── gui_tensile.py       # 拉伸报告界面
├── gui_vda.py           # VDA弯曲界面
├── gui_hardness.py      # 硬度提取界面
├── gui_origin.py        # 相变点绘图界面
├── tensile_processor.py # 拉伸数据处理
├── vda_processor.py     # VDA数据处理
├── origin_processor.py  # Origin绘图处理
├── ppt_utils.py         # PPT工具函数
├── 拉伸模板.pptx        # 拉伸报告模板
└── VDA弯曲角模板.pptx   # VDA报告模板
