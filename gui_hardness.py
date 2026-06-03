import tkinter as tk
from tkinter import messagebox
import processor
from gui_shared import (
    ScrollableFrame, browse_file, setup_drag_drop, COLORS,
    FONTS, create_button, create_entry, create_field_label, create_page,
    create_radiobutton, create_section, create_status_text
)

class HardnessFrame(tk.Frame):
    def __init__(self, parent):
        # 初始化时设置背景色
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.cached_hardness_data = [] 
        self.setup_ui()

    def setup_ui(self):
        for widget in self.winfo_children():
            widget.destroy()
            
        self.configure(bg=COLORS['bg_dark'])

        page = create_page(self)
        self.hard_pdf_src = tk.StringVar()
        self.hard_precision = tk.IntVar(value=1)

        source_section = create_section(page, "显微硬度数据提取", "导入 PDF 报告，提取硬度均值与标准差。")
        create_field_label(source_section, "PDF 数据源").grid(row=0, column=0, columnspan=2, sticky='w')
        entry = create_entry(source_section, self.hard_pdf_src, width=54, mono=True)
        entry.grid(row=1, column=0, padx=(0, 10), pady=(6, 0), sticky='ew', ipady=9)
        btn_browse = create_button(
            source_section,
            "浏览文件",
            lambda: browse_file(self.hard_pdf_src, [("PDF Files", "*.pdf")]),
            "secondary",
        )
        btn_browse.grid(row=1, column=1, sticky='ew', pady=(6, 0))
        source_section.columnconfigure(0, weight=1)

        setup_drag_drop(self, self.hard_pdf_src)
        setup_drag_drop(page, self.hard_pdf_src)
        setup_drag_drop(source_section, self.hard_pdf_src)
        setup_drag_drop(entry, self.hard_pdf_src)

        ctrl_frame = create_section(page, "显示选项", "调整结果数值的显示精度。")
        setup_drag_drop(ctrl_frame, self.hard_pdf_src)

        create_field_label(ctrl_frame, "显示精度").pack(anchor='w')
        precision_row = tk.Frame(ctrl_frame, bg=COLORS['bg_medium'])
        precision_row.pack(fill='x', pady=(6, 0))
        for val, text in [(0, "整数"), (1, "1位小数"), (2, "2位小数")]:
            create_radiobutton(precision_row, text, self.hard_precision, val, self.refresh_hardness_list).pack(side="left", padx=(0, 18))

        create_button(page, "开始提取数据", self.start_extract, "primary").pack(fill='x', pady=(0, 14))

        result_section = create_section(page, "提取结果", "结果可直接复制到报告或表格。")
        result_section.pack_configure(fill="both", expand=True)
        self.list_container = tk.Frame(result_section, bg=COLORS['border'], padx=1, pady=1)
        self.list_container.pack(fill="both", expand=True, pady=(2, 0))
        setup_drag_drop(self.list_container, self.hard_pdf_src)

        self.hard_scroll = ScrollableFrame(self.list_container, style_bg=COLORS['bg_medium'])
        self.hard_scroll.pack(fill="both", expand=True)
        setup_drag_drop(self.hard_scroll, self.hard_pdf_src)
        setup_drag_drop(self.hard_scroll.canvas, self.hard_pdf_src)
        setup_drag_drop(self.hard_scroll.scrollable_frame, self.hard_pdf_src)

        initial_label = create_status_text(self.hard_scroll.scrollable_frame, "暂无数据，请先提取。", "muted")
        initial_label.pack(pady=40)
        setup_drag_drop(initial_label, self.hard_pdf_src)

    def start_extract(self):
        p = self.hard_pdf_src.get()
        if not p:
            messagebox.showwarning("提示", "请先选择或拖入 PDF 文件")
            return
        
        # 清空旧显示
        self.clear_list()
        create_status_text(self.hard_scroll.scrollable_frame, "正在处理中...", "accent").pack(pady=20)
        self.update() 

        try:
            self.cached_hardness_data = processor.parse_hardness_report(p)
            self.refresh_hardness_list()
        except Exception as e:
            self.clear_list()
            create_status_text(self.hard_scroll.scrollable_frame, f"处理出错: {e}", "warning").pack(pady=20)

    def clear_list(self):
        for widget in self.hard_scroll.scrollable_frame.winfo_children():
            widget.destroy()

    def refresh_hardness_list(self):
        self.clear_list()
            
        if not self.cached_hardness_data:
            return

        if "error" in self.cached_hardness_data[0]:
             create_status_text(self.hard_scroll.scrollable_frame, f"错误: {self.cached_hardness_data[0]['error']}", "danger").pack()
             return

        decimals = self.hard_precision.get()
        
        # --- 列表表头 ---
        header_frame = tk.Frame(self.hard_scroll.scrollable_frame, bg=COLORS['bg_light'], height=34)
        header_frame.pack(fill="x", pady=(0, 2))
        
        headers = [("序号", 8), ("Mean ± SD (硬度值)", 30), ("操作", 10)]
        for txt, w in headers:
            tk.Label(header_frame, text=txt, width=w, 
                    bg=COLORS['bg_light'], fg=COLORS['text'], font=FONTS['small']).pack(side="left", padx=8, pady=7)

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
                    bg=row_bg, fg=COLORS['text'], font=FONTS['body']).pack(side="left", padx=8, pady=9)
            
            # 数值显示 (Entry)
            lbl_val = tk.Entry(row_frame, width=30, justify='center', font=FONTS['body'],
                             bg=COLORS['input_bg'], fg=COLORS['accent'],
                             relief='flat', bd=0, highlightthickness=1,
                             highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'])
            lbl_val.insert(0, val_str)
            # lbl_val.configure(state='readonly') # 如果想要完全只读可以取消注释，但这样无法选中复制
            lbl_val.pack(side="left", padx=5)
            
            # 复制按钮
            btn = create_button(row_frame, "复制", lambda: None, "secondary")
            btn.configure(command=lambda t=val_str, b=btn: self.copy_to_clipboard(t, b))
            btn.pack(side="left", padx=8, pady=4)

    def copy_to_clipboard(self, text, btn_widget):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        
        orig_bg = btn_widget.cget("bg")
        orig_text = btn_widget.cget("text")
        
        btn_widget.configure(text="已复制!", bg=COLORS['success'], fg='white')
        self.after(1000, lambda: btn_widget.configure(text=orig_text, bg=orig_bg, fg=COLORS['text']))
