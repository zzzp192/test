import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD 

# 导入所有功能模块
from gui_tensile import TensileFrame
from gui_vda import VDAFrame
from gui_hardness import HardnessFrame
from gui_origin import OriginFrame  # <--- [新增] 导入 Origin 模块

from gui_shared import COLORS, update_theme_colors

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔬 试验报告助手 V2.8 Ultimate") # 版本号升级！
        self.root.geometry("900x750")
        
        self.current_theme = 'light' 
        self.setup_ui() 

    def setup_ui(self):
        self.root.configure(bg=COLORS['bg_dark'])
        
        for widget in self.root.winfo_children():
            widget.destroy()

        self.configure_styles()
        self.create_header()
        
        self.notebook = ttk.Notebook(self.root, style='Tech.TNotebook')
        self.notebook.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        # --- 添加标签页 ---
        self.tab_tensile = TensileFrame(self.notebook)
        self.notebook.add(self.tab_tensile, text="  ⚡ 拉伸报告  ")
        
        self.tab_vda = VDAFrame(self.notebook)
        self.notebook.add(self.tab_vda, text="  📐 VDA弯曲  ")

        self.tab_hard = HardnessFrame(self.notebook)
        self.notebook.add(self.tab_hard, text="  💎 硬度提取  ")

        # 相变点绘图 Tab
        self.tab_origin = OriginFrame(self.notebook)
        self.notebook.add(self.tab_origin, text="  相变点绘图  ")
        
        # 数据源同步：拉伸报告数据变化时同步到Origin
        self.tab_tensile.v_tensile_src.trace_add('write', self.sync_data_source)
        
        self.create_status_bar()

    # ... (其余代码如 configure_styles, create_header, toggle_theme 等保持不变) ...
    # 只需要确保上面的 setup_ui 更新了即可
    
    # ... (create_header, create_status_bar, toggle_theme 代码复用之前的即可)
    # 为了完整性，这里不需要重复粘贴 create_header 等未修改的辅助函数
    # 只要保证 setup_ui 里添加了 tab_origin 即可。

    def configure_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Tech.TNotebook', background=COLORS['bg_dark'], borderwidth=0)
        style.configure('Tech.TNotebook.Tab',
                       background=COLORS['bg_medium'],
                       foreground=COLORS['text'],
                       padding=[20, 12],
                       font=('微软雅黑', 11, 'bold'))
        style.map('Tech.TNotebook.Tab',
                 background=[('selected', COLORS['bg_light'])],
                 foreground=[('selected', COLORS['accent'])])

    def create_header(self):
        header = tk.Frame(self.root, bg=COLORS['bg_medium'], height=70)
        header.pack(fill='x', padx=15, pady=15)
        header.pack_propagate(False)
        
        title_frame = tk.Frame(header, bg=COLORS['bg_medium'])
        title_frame.pack(side='left', padx=20, pady=10)
        
        tk.Label(title_frame, text="🔬", font=('Segoe UI Emoji', 28),
                bg=COLORS['bg_medium'], fg=COLORS['accent']).pack(side='left')
        
        text_frame = tk.Frame(title_frame, bg=COLORS['bg_medium'])
        text_frame.pack(side='left', padx=15)
        
        tk.Label(text_frame, text="试验报告助手", font=('微软雅黑', 18, 'bold'),
                bg=COLORS['bg_medium'], fg=COLORS['text']).pack(anchor='w')
        
        # --- 右侧 ---
        right_frame = tk.Frame(header, bg=COLORS['bg_medium'])
        right_frame.pack(side='right', padx=20)

        icon = "🌞" if self.current_theme == 'dark' else "🌙"
        btn_theme = tk.Button(right_frame, text=icon + " 切换主题", 
                             command=self.toggle_theme,
                             bg=COLORS['bg_light'], fg=COLORS['text'],
                             relief='flat', cursor='hand2', font=('微软雅黑', 9))
        btn_theme.pack(side='left', padx=15)
        
    def create_status_bar(self):
        status = tk.Frame(self.root, bg=COLORS['bg_medium'], height=35)
        status.pack(fill='x', side='bottom', padx=15, pady=(0, 15))
        status.pack_propagate(False)
        tk.Label(status, text="● 系统就绪 | Origin Link: ON", font=('微软雅黑', 9),
                bg=COLORS['bg_medium'], fg=COLORS['success']).pack(side='left', padx=15)

    def toggle_theme(self):
        self.current_theme = 'light' if self.current_theme == 'dark' else 'dark'
        update_theme_colors(self.current_theme)
        self.setup_ui()
    
    def sync_data_source(self, *args):
        """同步拉伸报告数据源到Origin绘图"""
        src = self.tab_tensile.v_tensile_src.get()
        if src and (src.endswith('.xlsx') or src.endswith('.xls') or src.endswith('.csv')):
            self.tab_origin.set_data_source(src)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = MainApp(root)
    root.mainloop()
