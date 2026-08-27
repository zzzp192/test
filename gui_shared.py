#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.16 - 共享UI组件模块

软件名称：育材堂报告助手
版本号：V3.16
开发单位：育材堂
开发者：张桢
开发完成日期：2026年1月

模块功能：
    提供应用程序共享的UI组件、主题配置和工具函数。

主要组件：
    - THEMES: 深色/亮色主题配色方案
    - COLORS: 当前活动的颜色配置
    - ScrollableFrame: 可滚动的Frame容器组件
    - 文件拖拽处理函数
    - 路径处理工具函数

Copyright (c) 2026 育材堂. All rights reserved.
"""

# ============================================================
# 标准库导入
# ============================================================
import os
import sys
import re
from typing import Dict, List, Optional, Callable, Any

# ============================================================
# 第三方库导入
# ============================================================
import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import DND_FILES

# ============================================================
# 版本信息
# ============================================================
__version__ = "3.16"
__author__ = "张桢"
__copyright__ = "Copyright (c) 2026 育材堂"

# ============================================================
# 主题配置
# ============================================================
THEMES: Dict[str, Dict[str, str]] = {
    'dark': {
        'bg_dark': '#111827',
        'bg_medium': '#1F2937',
        'bg_light': '#374151',
        'accent': '#14B8A6',
        'accent_hover': '#2DD4BF',
        'accent_soft': '#134E4A',
        'cta': '#F97316',
        'cta_hover': '#FB923C',
        'text': '#F9FAFB',
        'text_dim': '#CBD5E1',
        'success': '#22C55E',
        'warning': '#F59E0B',
        'danger': '#EF4444',
        'border': '#334155',
        'input_bg': '#0F172A',
        'button_bg': '#14B8A6',
        'button_fg': '#FFFFFF',
        'row_even': '#172033',
        'row_odd': '#111827',
        'shadow': '#0B1120'
    },
    'light': {
        'bg_dark': '#F6F8FB',
        'bg_medium': '#FFFFFF',
        'bg_light': '#E8F5F3',
        'accent': '#0D9488',
        'accent_hover': '#0F766E',
        'accent_soft': '#CCFBF1',
        'cta': '#F97316',
        'cta_hover': '#EA580C',
        'text': '#134E4A',
        'text_dim': '#475569',
        'success': '#16A34A',
        'warning': '#D97706',
        'danger': '#DC2626',
        'border': '#DCE7E5',
        'input_bg': '#FFFFFF',
        'button_bg': '#0D9488',
        'button_fg': '#FFFFFF',
        'row_even': '#FFFFFF',
        'row_odd': '#F8FAFC',
        'shadow': '#E2E8F0'
    }
}

FONTS: Dict[str, tuple] = {
    'display': ('Microsoft YaHei UI', 19, 'bold'),
    'title': ('Microsoft YaHei UI', 13, 'bold'),
    'subtitle': ('Microsoft YaHei UI', 9),
    'body': ('Microsoft YaHei UI', 10),
    'body_bold': ('Microsoft YaHei UI', 10, 'bold'),
    'small': ('Microsoft YaHei UI', 9),
    'mono': ('Consolas', 10),
}

SPACING: Dict[str, int] = {
    'page_x': 24,
    'page_y': 20,
    'section_gap': 14,
    'control_gap': 10,
    'card_pad': 16,
}

# 当前活动的颜色配置（默认使用亮色主题）
COLORS: Dict[str, str] = THEMES['light'].copy()


def update_theme_colors(mode: str = 'dark') -> None:
    """
    更新主题颜色
    
    切换应用程序的主题配色方案。
    
    Args:
        mode: 主题模式，可选 'dark' 或 'light'
    """
    new_theme = THEMES.get(mode, THEMES['dark'])
    COLORS.clear()
    COLORS.update(new_theme)


# ============================================================
# Win11 风格共享控件
# ============================================================
def apply_hover(widget: tk.Widget, normal_bg: str, hover_bg: str) -> tk.Widget:
    """Add a lightweight hover state to clickable Tk widgets."""
    widget.bind("<Enter>", lambda _e: widget.configure(bg=hover_bg))
    widget.bind("<Leave>", lambda _e: widget.configure(bg=normal_bg))
    return widget


def create_page(parent: tk.Widget, scrollable: bool = False) -> tk.Frame:
    if scrollable:
        container = ScrollableFrame(parent, style_bg=COLORS['bg_dark'])
        container.pack(fill="both", expand=True, padx=SPACING['page_x'], pady=SPACING['page_y'])
        page = container.scrollable_frame
        page._scrollable_page = container
        return page

    page = tk.Frame(parent, bg=COLORS['bg_dark'])
    page.pack(fill="both", expand=True, padx=SPACING['page_x'], pady=SPACING['page_y'])
    return page


def create_section(parent: tk.Widget, title: str, subtitle: str = "") -> tk.Frame:
    outer = tk.Frame(parent, bg=COLORS['border'])
    outer.pack(fill="x", pady=(0, SPACING['section_gap']))

    inner = tk.Frame(outer, bg=COLORS['bg_medium'], padx=SPACING['card_pad'], pady=SPACING['card_pad'])
    inner.pack(fill="both", expand=True, padx=1, pady=1)

    tk.Label(inner, text=title, bg=COLORS['bg_medium'], fg=COLORS['text'], font=FONTS['title']).pack(anchor="w")
    if subtitle:
        tk.Label(inner, text=subtitle, bg=COLORS['bg_medium'], fg=COLORS['text_dim'], font=FONTS['subtitle']).pack(anchor="w", pady=(2, 12))
    else:
        tk.Frame(inner, bg=COLORS['bg_medium'], height=8).pack(fill="x")
    content = tk.Frame(inner, bg=COLORS['bg_medium'])
    content.pack(fill="both", expand=True)
    return content


def create_field_label(parent: tk.Widget, text: str) -> tk.Label:
    return tk.Label(parent, text=text, bg=COLORS['bg_medium'], fg=COLORS['text_dim'], font=FONTS['small'])


def create_entry(parent: tk.Widget, textvariable: tk.Variable, width: int = 40, mono: bool = False) -> tk.Entry:
    return tk.Entry(
        parent,
        textvariable=textvariable,
        width=width,
        font=FONTS['mono'] if mono else FONTS['body'],
        bg=COLORS['input_bg'],
        fg=COLORS['text'],
        insertbackground=COLORS['accent'],
        relief='flat',
        bd=0,
        highlightthickness=1,
        highlightbackground=COLORS['border'],
        highlightcolor=COLORS['accent'],
    )


def create_button(parent: tk.Widget, text: str, command: Callable, variant: str = "secondary") -> tk.Button:
    palette = {
        "primary": (COLORS['accent'], COLORS['accent_hover'], COLORS['button_fg']),
        "cta": (COLORS['cta'], COLORS['cta_hover'], COLORS['button_fg']),
        "success": (COLORS['success'], COLORS['accent_hover'], COLORS['button_fg']),
        "secondary": (COLORS['bg_light'], COLORS['accent_soft'], COLORS['text']),
    }
    bg, hover, fg = palette.get(variant, palette["secondary"])
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg,
        fg=fg,
        activebackground=hover,
        activeforeground=fg,
        relief='flat',
        bd=0,
        cursor='hand2',
        font=FONTS['body_bold'] if variant in ("primary", "cta", "success") else FONTS['body'],
        padx=18,
        pady=9,
    )
    return apply_hover(btn, bg, hover)


def create_spinbox(parent: tk.Widget, from_: float, to: float, textvariable: tk.Variable, width: int = 6, increment: float = 1) -> tk.Spinbox:
    return tk.Spinbox(
        parent,
        from_=from_,
        to=to,
        textvariable=textvariable,
        width=width,
        bg=COLORS['input_bg'],
        fg=COLORS['text'],
        buttonbackground=COLORS['bg_light'],
        insertbackground=COLORS['accent'],
        relief='flat',
        highlightthickness=1,
        highlightbackground=COLORS['border'],
        highlightcolor=COLORS['accent'],
        font=FONTS['body'],
        increment=increment,
    )


def create_checkbutton(parent: tk.Widget, text: str, variable: tk.Variable, command: Optional[Callable] = None) -> tk.Checkbutton:
    return tk.Checkbutton(
        parent,
        text=text,
        variable=variable,
        command=command,
        bg=COLORS['bg_medium'],
        fg=COLORS['text'],
        activebackground=COLORS['bg_medium'],
        activeforeground=COLORS['accent'],
        selectcolor=COLORS['input_bg'],
        font=FONTS['body'],
    )


def create_radiobutton(parent: tk.Widget, text: str, variable: tk.Variable, value: Any, command: Optional[Callable] = None) -> tk.Radiobutton:
    return tk.Radiobutton(
        parent,
        text=text,
        variable=variable,
        value=value,
        command=command,
        bg=COLORS['bg_medium'],
        fg=COLORS['text'],
        activebackground=COLORS['bg_medium'],
        activeforeground=COLORS['accent'],
        selectcolor=COLORS['input_bg'],
        font=FONTS['body'],
    )


def create_status_text(parent: tk.Widget, text: str, tone: str = "muted") -> tk.Label:
    color = {
        "muted": COLORS['text_dim'],
        "success": COLORS['success'],
        "warning": COLORS['warning'],
        "danger": COLORS['danger'],
        "accent": COLORS['accent'],
    }.get(tone, COLORS['text_dim'])
    return tk.Label(parent, text=text, bg=COLORS['bg_medium'], fg=color, font=FONTS['body'])


# ============================================================
# 路径工具函数
# ============================================================
def resource_path(relative_path: str) -> str:
    """
    获取资源文件的绝对路径
    
    支持PyInstaller打包后的资源路径解析。
    
    Args:
        relative_path: 相对路径
        
    Returns:
        str: 资源文件的绝对路径
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def get_unique_path(path: str) -> str:
    """
    生成唯一文件路径
    
    如果文件已存在，则在文件名后添加序号。
    
    Args:
        path: 原始文件路径
        
    Returns:
        str: 唯一的文件路径
    """
    if not os.path.exists(path):
        return path
    
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}_{i}{ext}"):
        i += 1
    return f"{base}_{i}{ext}"


# ============================================================
# 文件选择函数
# ============================================================
def browse_file(string_var: tk.StringVar, file_types: List[tuple]) -> None:
    """
    打开文件选择对话框
    
    Args:
        string_var: 用于存储选中文件路径的StringVar
        file_types: 文件类型过滤器列表，如 [("Excel Files", "*.xlsx")]
    """
    path = filedialog.askopenfilename(filetypes=file_types)
    if path:
        string_var.set(path)


# ============================================================
# 拖拽处理函数
# ============================================================
def parse_drop_paths(data: str) -> List[str]:
    """
    解析拖拽数据中的文件路径
    
    处理Windows拖拽事件中的文件路径数据，支持带空格的路径。
    
    Args:
        data: 拖拽事件的原始数据字符串
        
    Returns:
        List[str]: 解析出的文件路径列表
    """
    if '{' in data:
        # 处理带花括号的路径（包含空格的路径）
        paths = re.findall(r'\{([^}]+)\}', data)
        if not paths and data.startswith('{') and data.endswith('}'):
            paths = [data[1:-1]]
    else:
        # 普通空格分隔的路径
        paths = data.split()
    
    return [p for p in paths if p]


def setup_drag_drop(widget: tk.Widget, string_var: tk.StringVar) -> bool:
    """
    设置单文件拖拽功能
    
    为指定控件注册拖拽事件处理。
    
    Args:
        widget: 要注册拖拽的控件
        string_var: 用于存储拖拽文件路径的StringVar
        
    Returns:
        bool: 注册成功返回True，失败返回False
    """
    def _on_drop(event) -> None:
        paths = parse_drop_paths(event.data)
        if paths:
            string_var.set(paths[0])
    
    try:
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind('<<Drop>>', _on_drop)
        return True
    except Exception as e:
        print(f"拖拽注册失败: {e}")
        return False


def setup_drag_drop_listbox(
    listbox: tk.Listbox,
    file_list: List[str],
    callback: Optional[Callable] = None
) -> bool:
    """
    为Listbox设置多文件拖拽功能
    
    Args:
        listbox: 目标Listbox控件
        file_list: 用于存储文件路径的列表
        callback: 拖拽完成后的回调函数（可选）
        
    Returns:
        bool: 注册成功返回True，失败返回False
    """
    def _on_drop(event) -> None:
        paths = parse_drop_paths(event.data)
        for p in paths:
            if p and p not in file_list:
                file_list.append(p)
                listbox.insert('end', os.path.basename(p))
        if callback:
            callback()
    
    try:
        listbox.drop_target_register(DND_FILES)
        listbox.dnd_bind('<<Drop>>', _on_drop)
        return True
    except Exception as e:
        print(f"Listbox拖拽注册失败: {e}")
        return False


# ============================================================
# 可滚动Frame组件
# ============================================================
class ScrollableFrame(ttk.Frame):
    """
    可滚动的Frame容器组件
    
    提供带垂直滚动条的Frame容器，支持鼠标滚轮滚动。
    
    Attributes:
        canvas: 内部Canvas控件
        scrollable_frame: 可滚动的内容Frame
    """
    
    def __init__(self, container: tk.Widget, *args, **kwargs) -> None:
        """
        初始化可滚动Frame
        
        Args:
            container: 父容器
            *args: 传递给ttk.Frame的位置参数
            **kwargs: 传递给ttk.Frame的关键字参数
                style_bg: 背景颜色（可选）
        """
        bg_color = kwargs.pop('style_bg', COLORS['bg_medium'])
        super().__init__(container, *args, **kwargs)
        
        # 创建Canvas和滚动条
        self.canvas = tk.Canvas(self, bg=bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=bg_color)

        # 配置滚动区域
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # 创建窗口
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # 布局
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮事件
        self.scrollable_frame.bind('<Enter>', self._bind_mouse)
        self.scrollable_frame.bind('<Leave>', self._unbind_mouse)

    def _on_canvas_configure(self, event: tk.Event) -> None:
        """
        Canvas大小变化时调整内容宽度
        
        Args:
            event: Configure事件
        """
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mouse(self, event: tk.Event) -> None:
        """
        鼠标进入时绑定滚轮事件
        
        Args:
            event: Enter事件
        """
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _unbind_mouse(self, event: tk.Event) -> None:
        """
        鼠标离开时解绑滚轮事件
        
        Args:
            event: Leave事件
        """
        self.canvas.unbind_all("<MouseWheel>")
        
    def _on_mousewheel(self, event: tk.Event) -> None:
        """
        处理鼠标滚轮事件
        
        Args:
            event: MouseWheel事件
        """
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def update_bg(self, color: str) -> None:
        """
        更新背景颜色
        
        Args:
            color: 新的背景颜色
        """
        self.canvas.configure(bg=color)
        self.scrollable_frame.configure(bg=color)
