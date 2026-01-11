import tkinter as tk
from tkinter import filedialog, messagebox
import os
from tkinterdnd2 import DND_FILES
import origin_processor
from gui_shared import COLORS

class OriginFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.file_list = []
        self.setup_ui()

    def setup_ui(self):
        for widget in self.winfo_children():
            widget.destroy()
        self.configure(bg=COLORS['bg_dark'])

        main_frame = tk.LabelFrame(self, text="🔥 相变点绘图", padx=20, pady=20,
                             bg=COLORS['bg_dark'], fg=COLORS['accent'],
                             font=('微软雅黑', 11, 'bold'))
        main_frame.pack(fill="both", expand=True, padx=25, pady=25)

        self.o_template_path = tk.StringVar()

        tk.Label(main_frame, text="拖拽CSV文件到下方区域（支持多文件）:", 
                bg=COLORS['bg_dark'], fg=COLORS['text'],
                font=('微软雅黑', 10)).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0,5))
        
        self.drop_zone = tk.Listbox(main_frame, height=8, bg=COLORS['input_bg'], fg=COLORS['text'],
                                   selectmode=tk.EXTENDED, font=('Consolas', 9))
        self.drop_zone.grid(row=1, column=0, columnspan=3, sticky='nsew', pady=10)
        
        def do_register():
            try:
                self.drop_zone.drop_target_register(DND_FILES)
                self.drop_zone.dnd_bind('<<Drop>>', self.on_drop)
            except Exception as e:
                print(f"拖拽注册失败: {e}")
        
        self.drop_zone.after(100, do_register)

        btn_frame = tk.Frame(main_frame, bg=COLORS['bg_dark'])
        btn_frame.grid(row=2, column=0, columnspan=3, sticky='ew', pady=5)
        
        tk.Button(btn_frame, text="添加文件", command=self.add_files,
                 bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').pack(side='left', padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_files,
                 bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').pack(side='left', padx=5)

        tk.Label(main_frame, text="绘图模板:", bg=COLORS['bg_dark'], fg=COLORS['text']).grid(row=3, column=0, sticky='w', pady=10)
        tk.Entry(main_frame, textvariable=self.o_template_path, width=30, bg=COLORS['input_bg'], fg=COLORS['text']).grid(row=3, column=1, sticky='ew', padx=5)
        tk.Button(main_frame, text="选择", command=self.browse_template, bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').grid(row=3, column=2)

        tk.Button(main_frame, text="🚀 开始绘图", command=self.run_plot,
                 bg=COLORS['success'], fg=COLORS['button_fg'], font=("微软雅黑", 12, "bold"),
                 relief='flat', cursor='hand2').grid(row=4, column=0, columnspan=3, sticky='ew', ipady=10, pady=15)

        main_frame.columnconfigure(1, weight=1)

    def on_drop(self, event):
        files = self.parse_drop_data(event.data)
        for f in files:
            if f.endswith('.csv') and f not in self.file_list:
                self.file_list.append(f)
                self.drop_zone.insert(tk.END, os.path.basename(f))

    def parse_drop_data(self, data):
        files = []
        if '{' in data:
            import re
            files = re.findall(r'\{([^}]+)\}', data)
            remaining = re.sub(r'\{[^}]+\}', '', data).strip()
            if remaining:
                files.extend(remaining.split())
        else:
            files = data.split()
        return [f.strip() for f in files if f.strip()]

    def add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        for p in paths:
            if p not in self.file_list:
                self.file_list.append(p)
                self.drop_zone.insert(tk.END, os.path.basename(p))

    def clear_files(self):
        self.file_list.clear()
        self.drop_zone.delete(0, tk.END)

    def browse_template(self):
        p = filedialog.askopenfilename(
            initialdir="C:/Users/deity/Documents/OriginLab/User Files",
            filetypes=[("Origin Template", "*.otpu *.otp")])
        if p:
            self.o_template_path.set(p)

    def run_plot(self):
        if not self.file_list:
            return messagebox.showwarning("提示", "请先添加CSV文件")
        
        tmpl = self.o_template_path.get() or None
        try:
            result = origin_processor.plot_phase_change(self.file_list, tmpl)
            messagebox.showinfo("完成", result)
        except Exception as e:
            import traceback
            messagebox.showerror("错误", f"{e}\n{traceback.format_exc()}")

    def set_data_source(self, path):
        pass
