#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.15 - 主程序入口模块

软件名称：育材堂报告助手
版本号：V3.15
开发单位：育材堂
开发者：张桢
开发完成日期：2026年1月

功能描述：
    本软件是一款用于材料试验数据处理和报告生成的桌面工具，
    集成Origin绘图功能，支持拉伸、VDA弯曲、硬度等多种试验数据的处理。

主要功能模块：
    1. 拉伸报告处理 - 自动提取试样参数并生成PPT报告
    2. VDA弯曲报告处理 - 处理VDA弯曲试验数据
    3. 硬度数据提取 - 从PDF中提取显微硬度数据
    4. 相变点绘图 - 批量处理相变点CSV数据并绘图

技术特点：
    - 基于Python 3.11开发，使用Tkinter构建图形界面
    - 集成Origin绘图引擎，支持OLE对象嵌入
    - 支持文件拖拽操作
    - 支持深色/亮色主题切换

运行环境：
    - Windows 10/11
    - Python 3.11+
    - Origin 2019+

Copyright (c) 2026 育材堂. All rights reserved.
"""

# ============================================================
# 标准库导入
# ============================================================
import tkinter as tk
from tkinter import ttk
from typing import Optional

# ============================================================
# 第三方库导入
# ============================================================
from tkinterdnd2 import TkinterDnD

# ============================================================
# 本地模块导入
# ============================================================
from gui_tensile import TensileFrame
from gui_vda import VDAFrame
from gui_hardness import HardnessFrame
from gui_origin import OriginFrame
from gui_shared import COLORS, FONTS, create_button, update_theme_colors

# ============================================================
# 版本信息
# ============================================================
__version__ = "3.15"
__author__ = "张桢"
__copyright__ = "Copyright (c) 2026 育材堂"
__license__ = "Proprietary"


class MainApp:
    """
    主应用程序类
    
    负责创建和管理应用程序的主窗口、标签页和主题切换功能。
    
    Attributes:
        root: TkinterDnD根窗口实例
        current_theme: 当前主题模式 ('light' 或 'dark')
        notebook: 标签页容器
        tab_tensile: 拉伸报告标签页
        tab_vda: VDA弯曲报告标签页
        tab_hard: 硬度提取标签页
        tab_origin: 相变点绘图标签页
    """
    
    def __init__(self, root: TkinterDnD.Tk) -> None:
        """
        初始化主应用程序
        
        Args:
            root: TkinterDnD根窗口实例
        """
        self.root = root
        self.root.title("育材堂报告助手 V3.15")
        self.root.geometry("900x750")
        self.root.minsize(860, 700)
        
        self.current_theme: str = 'light'
        self.notebook: Optional[ttk.Notebook] = None
        self.tab_tensile: Optional[TensileFrame] = None
        self.tab_vda: Optional[VDAFrame] = None
        self.tab_hard: Optional[HardnessFrame] = None
        self.tab_origin: Optional[OriginFrame] = None
        
        self.setup_ui()

    def setup_ui(self) -> None:
        """
        设置用户界面
        
        创建主窗口的所有UI组件，包括标题栏、标签页和状态栏。
        """
        self.root.configure(bg=COLORS['bg_dark'])
        
        # 清除现有控件（用于主题切换时重建UI）
        for widget in self.root.winfo_children():
            widget.destroy()

        self.configure_styles()
        self.create_header()
        
        # 创建标签页容器
        self.notebook = ttk.Notebook(self.root, style='Tech.TNotebook')
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        
        # 添加功能标签页
        self.tab_tensile = TensileFrame(self.notebook)
        self.notebook.add(self.tab_tensile, text="  拉伸报告  ")
        
        self.tab_vda = VDAFrame(self.notebook)
        self.notebook.add(self.tab_vda, text="  VDA 弯曲  ")

        self.tab_hard = HardnessFrame(self.notebook)
        self.notebook.add(self.tab_hard, text="  硬度提取  ")

        self.tab_origin = OriginFrame(self.notebook)
        self.notebook.add(self.tab_origin, text="  相变点绘图  ")
        
        # 数据源同步：拉伸报告数据变化时同步到Origin
        self.tab_tensile.v_tensile_src.trace_add('write', self.sync_data_source)
        
        self.create_status_bar()

    def configure_styles(self) -> None:
        """
        配置ttk样式
        
        设置标签页的外观样式，包括背景色、前景色和字体。
        """
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Tech.TNotebook', background=COLORS['bg_dark'], borderwidth=0, tabmargins=[0, 0, 0, 0])
        style.configure('Tech.TNotebook.Tab',
                       background=COLORS['bg_medium'],
                       foreground=COLORS['text'],
                       borderwidth=0,
                       padding=[22, 12],
                       font=FONTS['body_bold'])
        style.map('Tech.TNotebook.Tab',
                 background=[('selected', COLORS['accent_soft']), ('active', COLORS['bg_light'])],
                 foreground=[('selected', COLORS['accent']), ('active', COLORS['text'])])
        style.configure('Vertical.TScrollbar',
                        background=COLORS['bg_light'],
                        troughcolor=COLORS['bg_medium'],
                        bordercolor=COLORS['border'],
                        arrowcolor=COLORS['text_dim'])

    def create_header(self) -> None:
        """
        创建标题栏
        
        包含应用程序图标、标题和主题切换按钮。
        """
        header = tk.Frame(self.root, bg=COLORS['bg_medium'], height=76)
        header.pack(fill='x', padx=20, pady=18)
        header.pack_propagate(False)
        
        # 左侧标题区域
        title_frame = tk.Frame(header, bg=COLORS['bg_medium'])
        title_frame.pack(side='left', padx=18, pady=12)
        
        text_frame = tk.Frame(title_frame, bg=COLORS['bg_medium'])
        text_frame.pack(side='left')
        
        tk.Label(text_frame, text="育材堂报告助手", font=FONTS['display'],
                bg=COLORS['bg_medium'], fg=COLORS['text']).pack(anchor='w')
        tk.Label(text_frame, text="材料试验报告处理与 Origin 绘图工具  V3.15", font=FONTS['small'],
                bg=COLORS['bg_medium'], fg=COLORS['text_dim']).pack(anchor='w', pady=(3, 0))
        
        # 右侧控制区域
        right_frame = tk.Frame(header, bg=COLORS['bg_medium'])
        right_frame.pack(side='right', padx=18)

        text = "切换亮色" if self.current_theme == 'dark' else "切换深色"
        btn_theme = create_button(right_frame, text, self.toggle_theme, "secondary")
        btn_theme.pack(side='left')
        
    def create_status_bar(self) -> None:
        """
        创建状态栏
        
        显示系统状态和Origin连接状态。
        """
        status = tk.Frame(self.root, bg=COLORS['bg_medium'], height=40)
        status.pack(fill='x', side='bottom', padx=20, pady=(0, 16))
        status.pack_propagate(False)
        tk.Label(status, text="● 系统就绪", font=FONTS['small'],
                bg=COLORS['bg_medium'], fg=COLORS['success']).pack(side='left', padx=16)
        tk.Label(status, text="Origin Link: ON", font=FONTS['small'],
                bg=COLORS['bg_medium'], fg=COLORS['text_dim']).pack(side='left', padx=8)

    def toggle_theme(self) -> None:
        """
        切换主题
        
        在亮色和暗色主题之间切换，并重建UI以应用新主题。
        """
        self.current_theme = 'light' if self.current_theme == 'dark' else 'dark'
        update_theme_colors(self.current_theme)
        self.setup_ui()
    
    def sync_data_source(self, *args) -> None:
        """
        同步数据源
        
        当拉伸报告数据源变化时，自动同步到Origin绘图模块。
        
        Args:
            *args: trace_add回调参数（未使用）
        """
        src = self.tab_tensile.v_tensile_src.get()
        if src and (src.endswith('.xlsx') or src.endswith('.xls') or src.endswith('.csv')):
            self.tab_origin.set_data_source(src)


def main() -> None:
    """
    程序入口函数
    
    创建主窗口并启动事件循环。
    """
    root = TkinterDnD.Tk()
    app = MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
