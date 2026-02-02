import tkinter as tk
from tkinter import ttk, messagebox
import processor
from gui_shared import ScrollableFrame, browse_file, setup_drag_drop, COLORS

class HardnessFrame(tk.Frame):
    def __init__(self, parent):
        # 初始化时设置背景色
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.cached_hardness_data = [] 
        self.setup_ui()

    def setup_ui(self):
        # --- 1. 清理旧控件 (用于主题切换刷新) ---
        for widget in self.winfo_children():
            widget.destroy()
            
        self.configure(bg=COLORS['bg_dark'])

        # --- 2. 主容器 ---
        # 模仿 gui_tensile 的样式
        main_frame = tk.LabelFrame(self, text="💎 显微硬度数据提取", padx=20, pady=20,
                             bg=COLORS['bg_dark'], fg=COLORS['accent'],
                             font=('微软雅黑', 11, 'bold'))
        main_frame.pack(fill="both", expand=True, padx=25, pady=25)

        # 变量初始化
        self.hard_pdf_src = tk.StringVar()
        self.hard_precision = tk.IntVar(value=1)

        # --- 3. 顶部操作区 (Grid布局) ---

        # 提示标签
        label_source = tk.Label(main_frame, text="PDF 数据源 | 💡 可拖拽到整个界面任意位置",
                bg=COLORS['bg_dark'], fg=COLORS['text'],
                font=('微软雅黑', 10))
        label_source.grid(row=0, column=0, sticky='w', pady=(0,5))

        # 输入框 (科技感样式)
        entry = tk.Entry(main_frame, textvariable=self.hard_pdf_src,
                        font=('Consolas', 10), bg=COLORS['input_bg'], fg=COLORS['text'],
                        insertbackground=COLORS['accent'], relief='flat', highlightthickness=1,
                        highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'])
        entry.grid(row=1, column=0, padx=(0,10), sticky='ew', ipady=8)

        # 浏览按钮
        btn_browse = tk.Button(main_frame, text="📂 浏览",
                              command=lambda: browse_file(self.hard_pdf_src, [("PDF Files", "*.pdf")]),
                              bg=COLORS['bg_light'], fg=COLORS['text'], font=('微软雅黑', 9),
                              relief='flat', cursor='hand2')
        btn_browse.grid(row=1, column=1, sticky='ew', ipady=5, padx=5)

        # 注册拖拽 - 扩展到整个界面
        setup_drag_drop(self, self.hard_pdf_src)
        setup_drag_drop(main_frame, self.hard_pdf_src)
        setup_drag_drop(label_source, self.hard_pdf_src)
        setup_drag_drop(entry, self.hard_pdf_src)

        # 让输入框拉伸
        main_frame.columnconfigure(0, weight=1)

        # --- 4. 选项与控制区 ---
        ctrl_frame = tk.Frame(main_frame, bg=COLORS['bg_dark'])
        ctrl_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
        setup_drag_drop(ctrl_frame, self.hard_pdf_src)

        # 精度选择
        label_precision = tk.Label(ctrl_frame, text="显示精度: ", bg=COLORS['bg_dark'], fg=COLORS['text'])
        label_precision.pack(side="left")
        setup_drag_drop(label_precision, self.hard_pdf_src)
        
        style = ttk.Style()
        style.configure('Tech.TRadiobutton', background=COLORS['bg_dark'], foreground=COLORS['text'])
        
        for val, text in [(0, "整数"), (1, "1位小数"), (2, "2位小数")]:
            rb = tk.Radiobutton(ctrl_frame, text=text, variable=self.hard_precision, 
                               value=val, command=self.refresh_hardness_list,
                               bg=COLORS['bg_dark'], fg=COLORS['text'], selectcolor=COLORS['bg_medium'],
                               activebackground=COLORS['bg_dark'], activeforeground=COLORS['accent'])
            rb.pack(side="left", padx=5)

        # 提取按钮
        btn_extract = tk.Button(ctrl_frame, text="🚀 开始提取数据", command=self.start_extract, 
                 bg=COLORS['accent'], fg=COLORS['button_fg'], 
                 font=('微软雅黑', 10, 'bold'), relief='flat', cursor='hand2', padx=20)
        btn_extract.pack(side="right")

        # --- 5. 结果列表区 ---
        # 使用自定义背景色的 ScrollableFrame
        self.list_container = tk.Frame(main_frame, bg=COLORS['bg_medium'], padx=2, pady=2)
        self.list_container.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=10)
        main_frame.rowconfigure(4, weight=1) # 让列表区占用剩余高度
        setup_drag_drop(self.list_container, self.hard_pdf_src)

        self.hard_scroll = ScrollableFrame(self.list_container, style_bg=COLORS['bg_medium'])
        self.hard_scroll.pack(fill="both", expand=True)
        setup_drag_drop(self.hard_scroll, self.hard_pdf_src)
        setup_drag_drop(self.hard_scroll.canvas, self.hard_pdf_src)
        setup_drag_drop(self.hard_scroll.scrollable_frame, self.hard_pdf_src)

        # 初始提示
        initial_label = tk.Label(self.hard_scroll.scrollable_frame, text="暂无数据，请先提取...",
                bg=COLORS['bg_medium'], fg=COLORS['text_dim'], font=('微软雅黑', 10))
        initial_label.pack(pady=40)
        setup_drag_drop(initial_label, self.hard_pdf_src)

    def start_extract(self):
        p = self.hard_pdf_src.get()
        if not p:
            messagebox.showwarning("提示", "请先选择或拖入 PDF 文件")
            return
        
        # 清空旧显示
        self.clear_list()
        tk.Label(self.hard_scroll.scrollable_frame, text="正在处理中...", bg=COLORS['bg_medium'], fg=COLORS['accent']).pack(pady=20)
        self.update() 

        try:
            self.cached_hardness_data = processor.parse_hardness_report(p)
            self.refresh_hardness_list()
        except Exception as e:
            self.clear_list()
            tk.Label(self.hard_scroll.scrollable_frame, text=f"处理出错: {e}", bg=COLORS['bg_medium'], fg=COLORS['warning']).pack(pady=20)

    def clear_list(self):
        for widget in self.hard_scroll.scrollable_frame.winfo_children():
            widget.destroy()

    def refresh_hardness_list(self):
        self.clear_list()
            
        if not self.cached_hardness_data:
            return

        if "error" in self.cached_hardness_data[0]:
             tk.Label(self.hard_scroll.scrollable_frame, text=f"错误: {self.cached_hardness_data[0]['error']}", 
                      bg=COLORS['bg_medium'], fg="red").pack()
             return

        decimals = self.hard_precision.get()
        
        # --- 列表表头 ---
        header_frame = tk.Frame(self.hard_scroll.scrollable_frame, bg=COLORS['bg_light'], height=30)
        header_frame.pack(fill="x", pady=(0, 2))
        
        headers = [("序号", 8), ("Mean ± SD (硬度值)", 30), ("操作", 10)]
        for txt, w in headers:
            tk.Label(header_frame, text=txt, width=w, 
                    bg=COLORS['bg_light'], fg=COLORS['text'], font=('微软雅黑', 9, 'bold')).pack(side="left", padx=5, pady=5)

        # --- 数据行 ---
        for i, item in enumerate(self.cached_hardness_data):
            # 斑马纹交替颜色
            row_bg = COLORS['row_even'] if i % 2 == 0 else COLORS['row_odd']
            
            row_frame = tk.Frame(self.hard_scroll.scrollable_frame, bg=row_bg)
            row_frame.pack(fill="x", pady=1)
            
            try:
                m = float(item['mean'])
                s = float(item['sd'])
                val_str = f"{m:.{decimals}f}±{s:.{decimals}f}"
            except:
                val_str = f"{item['mean']}±{item['sd']}"

            # 序号
            tk.Label(row_frame, text=f"Group {item['id']}", width=8, anchor="w",
                    bg=row_bg, fg=COLORS['text']).pack(side="left", padx=5, pady=8)
            
            # 数值显示 (Entry)
            lbl_val = tk.Entry(row_frame, width=30, justify='center', font=('Arial', 10),
                             bg=COLORS['input_bg'], fg=COLORS['accent'],
                             relief='flat', bd=0)
            lbl_val.insert(0, val_str)
            # lbl_val.configure(state='readonly') # 如果想要完全只读可以取消注释，但这样无法选中复制
            lbl_val.pack(side="left", padx=5)
            
            # 复制按钮
            btn = tk.Button(row_frame, text="复制", width=8, cursor="hand2",
                           bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat', font=('微软雅黑', 8))
            btn.configure(command=lambda t=val_str, b=btn: self.copy_to_clipboard(t, b))
            btn.pack(side="left", padx=5)

    def copy_to_clipboard(self, text, btn_widget):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        
        orig_bg = btn_widget.cget("bg")
        orig_text = btn_widget.cget("text")
        
        btn_widget.configure(text="已复制!", bg=COLORS['success'], fg='white')
        self.after(1000, lambda: btn_widget.configure(text=orig_text, bg=orig_bg, fg=COLORS['text']))