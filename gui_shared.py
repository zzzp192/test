import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import DND_FILES
import re

# --- 1. 定义两套主题色 ---
THEMES = {
    'dark': {
        'bg_dark': '#1a1a2e',
        'bg_medium': '#16213e', 
        'bg_light': '#0f3460',
        'accent': '#00d9ff',
        'accent_hover': '#00b8d4',
        'text': '#e8e8e8',
        'text_dim': '#a0a0a0',
        'success': '#00e676',
        'warning': '#ffc107',
        'border': '#2a3f5f',
        'input_bg': '#0d1b2a',
        'button_bg': '#00d9ff',
        'button_fg': '#1a1a2e',
        'row_even': '#1a1a2e',
        'row_odd': '#202a44'
    },
    'light': {
        'bg_dark': '#f0f2f5',
        'bg_medium': '#ffffff',
        'bg_light': '#e1e4e8',
        'accent': '#007bff',
        'accent_hover': '#0056b3',
        'text': '#333333',
        'text_dim': '#666666',
        'success': '#28a745',
        'warning': '#ffc107',
        'border': '#ced4da',
        'input_bg': '#ffffff',
        'button_bg': '#007bff',
        'button_fg': '#ffffff',
        'row_even': '#ffffff',
        'row_odd': '#f8f9fa'
    }
}

COLORS = THEMES['light'].copy()

def update_theme_colors(mode='dark'):
    new_theme = THEMES.get(mode, THEMES['dark'])
    COLORS.clear()
    COLORS.update(new_theme)

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_unique_path(path):
    """生成唯一文件路径，如果存在则添加序号"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}_{i}{ext}"):
        i += 1
    return f"{base}_{i}{ext}"

def browse_file(string_var, file_types):
    p = filedialog.askopenfilename(filetypes=file_types)
    if p:
        string_var.set(p)

def parse_drop_paths(data):
    """解析拖拽数据中的文件路径"""
    if '{' in data:
        paths = re.findall(r'\{([^}]+)\}', data)
        if not paths and data.startswith('{') and data.endswith('}'):
            paths = [data[1:-1]]
    else:
        paths = data.split()
    return [p for p in paths if p]

def setup_drag_drop(widget, string_var):
    """设置单文件拖拽"""
    def _on_drop(event):
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

def setup_drag_drop_listbox(listbox, file_list, callback=None):
    """为Listbox设置多文件拖拽"""
    def _on_drop(event):
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

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        bg_color = kwargs.pop('style_bg', COLORS['bg_medium'])
        super().__init__(container, *args, **kwargs)
        
        self.canvas = tk.Canvas(self, bg=bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=bg_color)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.scrollable_frame.bind('<Enter>', self._bind_mouse)
        self.scrollable_frame.bind('<Leave>', self._unbind_mouse)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mouse(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _unbind_mouse(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def update_bg(self, color):
        self.canvas.configure(bg=color)
        self.scrollable_frame.configure(bg=color)
