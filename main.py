#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.16 - 主程序入口模块

软件名称：育材堂报告助手
版本号：V3.16
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
from typing import Any, Optional, TYPE_CHECKING

# ============================================================
# 第三方库导入
# ============================================================
from tkinterdnd2 import TkinterDnD

from gui_shared import COLORS, FONTS, create_button, update_theme_colors

if TYPE_CHECKING:
    from gui_tensile import TensileFrame
    from gui_vda import VDAFrame
    from gui_hardness import HardnessFrame
    from gui_origin import OriginFrame

# ============================================================
# 版本信息
# ============================================================
__version__ = "3.16"
__author__ = "张桢"
__copyright__ = "Copyright (c) 2026 育材堂"
__license__ = "Proprietary"


def get_bootloader_splash() -> Optional[Any]:
    try:
        import pyi_splash
    except ImportError:
        return None
    try:
        return pyi_splash if pyi_splash.is_alive() else None
    except RuntimeError:
        return None


class StartupSplash:
    """启动期提示层，尽早给用户可见反馈。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self._step = 0
        self._job: Optional[str] = None
        self._destroyed = False

        root.title(f"育材堂报告助手 V{__version__}")
        root.geometry("440x260")
        root.minsize(440, 260)
        root.resizable(False, False)
        root.configure(bg=COLORS['bg_dark'])

        root.update_idletasks()
        x = max((root.winfo_screenwidth() - 440) // 2, 0)
        y = max((root.winfo_screenheight() - 260) // 2, 0)
        root.geometry(f"440x260+{x}+{y}")

        self.window = tk.Frame(root, bg=COLORS['bg_medium'], padx=28, pady=24)
        self.window.place(relx=0.5, rely=0.5, anchor='center', width=380, height=190)

        tk.Label(
            self.window,
            text="育材堂报告助手",
            font=FONTS['title'],
            bg=COLORS['bg_medium'],
            fg=COLORS['text'],
        ).pack(anchor='w')

        self.status = tk.Label(
            self.window,
            text="正在启动，请稍候",
            font=FONTS['body'],
            bg=COLORS['bg_medium'],
            fg=COLORS['text_dim'],
        )
        self.status.pack(anchor='w', pady=(10, 12))

        self.progress = tk.Canvas(
            self.window,
            height=8,
            bg=COLORS['bg_light'],
            highlightthickness=0,
        )
        self.progress.pack(fill='x', pady=(0, 12))
        self.bar = self.progress.create_rectangle(0, 0, 90, 8, fill=COLORS['accent'], outline='')

        self.spinner = tk.Label(
            self.window,
            text=".",
            font=FONTS['small'],
            bg=COLORS['bg_medium'],
            fg=COLORS['accent'],
        )
        self.spinner.pack(anchor='w')
        self.animate()

    def set_status(self, text: str) -> None:
        if self._destroyed:
            return
        self.status.configure(text=text)
        self.root.update_idletasks()

    def animate(self) -> None:
        if self._destroyed:
            return
        width = max(self.progress.winfo_width(), 260)
        start = (self._step * 14) % max(width, 1)
        end = min(start + 90, width)
        self.progress.coords(self.bar, start, 0, end, 8)
        self.spinner.configure(text="正在加载" + "." * ((self._step % 3) + 1))
        self._step += 1
        self._job = self.root.after(120, self.animate)

    def destroy(self) -> None:
        self._destroyed = True
        if self._job:
            try:
                self.root.after_cancel(self._job)
            except tk.TclError:
                pass
        if self.window.winfo_exists():
            self.window.destroy()


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
    
    def __init__(self, root: tk.Tk, startup_splash: Optional[StartupSplash] = None) -> None:
        """
        初始化主应用程序
        
        Args:
            root: TkinterDnD根窗口实例
        """
        self.root = root
        self.startup_splash = startup_splash
        self.root.title(f"育材堂报告助手 V{__version__}")
        self.root.geometry("900x750")
        self.root.minsize(860, 700)
        self.root.resizable(True, True)
        
        self.current_theme: str = 'light'
        self.notebook: Optional[ttk.Notebook] = None
        self.tab_tensile: Optional["TensileFrame"] = None
        self.tab_vda: Optional["VDAFrame"] = None
        self.tab_hard: Optional["HardnessFrame"] = None
        self.tab_origin: Optional["OriginFrame"] = None
        self._frame_classes: Optional[tuple[Any, Any, Any, Any]] = None
        
        self.setup_ui()

    def setup_ui(self) -> None:
        """
        设置用户界面
        
        创建主窗口的所有UI组件，包括标题栏、标签页和状态栏。
        """
        if self.startup_splash:
            self.startup_splash.set_status("正在加载功能模块")
        TensileFrame, VDAFrame, HardnessFrame, OriginFrame = self.load_frame_classes()
        if self.startup_splash:
            self.startup_splash.set_status("正在准备界面")
            self.root.update()
            self.root.withdraw()
            self.startup_splash.destroy()
            self.startup_splash = None

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

    def load_frame_classes(self) -> tuple[Any, Any, Any, Any]:
        if self._frame_classes is None:
            from gui_tensile import TensileFrame
            from gui_vda import VDAFrame
            from gui_hardness import HardnessFrame
            from gui_origin import OriginFrame

            self._frame_classes = (TensileFrame, VDAFrame, HardnessFrame, OriginFrame)
        return self._frame_classes

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
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)
        header.rowconfigure(0, weight=1)
        
        # 左侧标题区域
        title_frame = tk.Frame(header, bg=COLORS['bg_medium'])
        title_frame.grid(row=0, column=0, sticky='nsew', padx=18, pady=10)
        title_frame.columnconfigure(0, weight=1)
        title_frame.rowconfigure(0, weight=1)
        
        text_frame = tk.Frame(title_frame, bg=COLORS['bg_medium'])
        text_frame.grid(row=0, column=0, sticky='nsew')
        text_frame.columnconfigure(0, weight=1)
        
        tk.Label(text_frame, text="育材堂报告助手", font=FONTS['display'],
                bg=COLORS['bg_medium'], fg=COLORS['text']).grid(row=0, column=0, sticky='w')
        subtitle = tk.Label(text_frame, text=f"材料试验报告处理与Origin绘图工具 V{__version__}", font=FONTS['small'],
                bg=COLORS['bg_medium'], fg=COLORS['text_dim'], anchor='w', justify='left', wraplength=520)
        subtitle.grid(row=1, column=0, sticky='w', pady=(3, 0))
        text_frame.bind(
            '<Configure>',
            lambda event, label=subtitle: label.configure(wraplength=max(event.width - 2, 120))
        )
        
        # 右侧控制区域
        right_frame = tk.Frame(header, bg=COLORS['bg_medium'])
        right_frame.grid(row=0, column=1, sticky='e', padx=18)

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
    bootloader_splash = get_bootloader_splash()
    # TkinterDnD.Tk() creates a visible native window before it loads the
    # tkdnd package.  On slower computers that package load leaves users
    # staring at an empty window titled "tk".  Create and hide the standard
    # Tk root first, then enable drag-and-drop behind a real loading screen.
    root = tk.Tk()
    root.withdraw()

    if bootloader_splash:
        bootloader_splash.update_text("正在打开加载界面")

    splash = StartupSplash(root)
    root.deiconify()
    root.update()

    splash.set_status("正在启用文件拖放功能")
    root.TkdndVersion = TkinterDnD._require(root)

    if bootloader_splash:
        bootloader_splash.update_text("正在加载功能模块")

    app = MainApp(root, startup_splash=splash)
    root.update_idletasks()

    if bootloader_splash:
        bootloader_splash.close()
    root.deiconify()

    root.mainloop()


if __name__ == "__main__":
    main()
