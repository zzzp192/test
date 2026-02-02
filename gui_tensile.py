import tkinter as tk
from tkinter import messagebox, filedialog
import os
import tensile_processor
import origin_processor
import config_manager
from gui_shared import resource_path, browse_file, get_unique_path, COLORS
from tkinterdnd2 import DND_FILES

class TensileFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.setup_ui()

    def setup_ui(self):
        for widget in self.winfo_children():
            widget.destroy()
        self.configure(bg=COLORS['bg_dark'])
        
        frame = tk.LabelFrame(self, text="📊 拉伸实验报告生成", padx=20, pady=20,
                             bg=COLORS['bg_dark'], fg=COLORS['accent'], font=('微软雅黑', 11, 'bold'))
        frame.pack(fill="x", padx=25, pady=25)
        
        self.v_tensile_src = tk.StringVar()
        self.v_include_ag = tk.BooleanVar(value=True)
        
        # 文件选择
        label_source = tk.Label(frame, text="选择原始数据文件 (Word/Excel) | 💡 可拖拽到整个界面任意位置", bg=COLORS['bg_dark'], fg=COLORS['text'], font=('微软雅黑', 10))
        label_source.grid(row=0, column=0, columnspan=4, sticky='w', pady=(0,10))

        entry = tk.Entry(frame, textvariable=self.v_tensile_src, width=50, font=('Consolas', 10), bg=COLORS['input_bg'], fg=COLORS['text'],
                        insertbackground=COLORS['accent'], relief='flat', highlightthickness=1, highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'])
        entry.grid(row=1, column=0, columnspan=3, padx=(0,10), pady=5, sticky='ew', ipady=8)

        tk.Button(frame, text="📂 浏览", command=lambda: browse_file(self.v_tensile_src, [("Data Files", "*.xlsx *.xls *.csv *.docx")]),
                 bg=COLORS['bg_light'], fg=COLORS['text'], font=('微软雅黑', 9), relief='flat', cursor='hand2', padx=15).grid(row=1, column=3, padx=5, ipady=5)

        # 注册拖拽 - 扩展到整个界面
        self._setup_dnd(self)
        self._setup_dnd(frame)
        self._setup_dnd(label_source)
        self._setup_dnd(entry)

        # 选项
        opt_frame = tk.Frame(frame, bg=COLORS['bg_dark'])
        opt_frame.grid(row=2, column=0, columnspan=4, sticky='w', pady=5)
        self._setup_dnd(opt_frame)
        tk.Checkbutton(opt_frame, text="包含 Ag (最大力总延伸率)", variable=self.v_include_ag, bg=COLORS['bg_dark'], fg=COLORS['text'], selectcolor=COLORS['bg_medium'], font=('微软雅黑', 9)).pack(side="left")

        # 绘图选项
        self.plot_frame = tk.LabelFrame(frame, text="绘图选项", padx=10, pady=10, bg=COLORS['bg_dark'], fg=COLORS['text_dim'], font=('微软雅黑', 9))
        self.plot_frame.grid(row=3, column=0, columnspan=4, sticky='ew', pady=10)
        self._setup_dnd(self.plot_frame)

        self.o_template = tk.StringVar(value=config_manager.get_template('tensile_template'))
        self.o_lines = tk.IntVar(value=12)
        self.o_swap_xy = tk.BooleanVar(value=True)
        self.o_width = tk.DoubleVar(value=15.0)
        self.o_height = tk.DoubleVar(value=12.0)
        self.o_copy_to_ppt = tk.BooleanVar(value=False)  # 默认不复制到PPT

        tk.Label(self.plot_frame, text="模板:", bg=COLORS['bg_dark'], fg=COLORS['text']).grid(row=0, column=0, sticky='w')
        tk.Entry(self.plot_frame, textvariable=self.o_template, width=25, bg=COLORS['input_bg'], fg=COLORS['text']).grid(row=0, column=1, sticky='ew', padx=5)
        tk.Button(self.plot_frame, text="选择", command=self.browse_template, bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').grid(row=0, column=2)

        tk.Label(self.plot_frame, text="每图曲线数:", bg=COLORS['bg_dark'], fg=COLORS['text']).grid(row=0, column=3, padx=(15,5))
        tk.Spinbox(self.plot_frame, from_=1, to=50, textvariable=self.o_lines, width=5, bg=COLORS['input_bg'], fg=COLORS['text']).grid(row=0, column=4)

        tk.Checkbutton(self.plot_frame, text="调换XY列", variable=self.o_swap_xy, bg=COLORS['bg_dark'], fg=COLORS['text'], selectcolor=COLORS['bg_medium']).grid(row=0, column=5, padx=15)

        # 图片尺寸选项（放在同一行）
        size_frame = tk.Frame(self.plot_frame, bg=COLORS['bg_dark'])
        size_frame.grid(row=1, column=0, columnspan=6, sticky='w', pady=(5,0))
        self._setup_dnd(size_frame)

        tk.Label(size_frame, text="图片宽(cm):", bg=COLORS['bg_dark'], fg=COLORS['text']).pack(side='left')
        tk.Spinbox(size_frame, from_=5, to=30, textvariable=self.o_width, width=5, bg=COLORS['input_bg'], fg=COLORS['text'], increment=0.5).pack(side='left', padx=(5,15))
        tk.Label(size_frame, text="图片高(cm):", bg=COLORS['bg_dark'], fg=COLORS['text']).pack(side='left')
        tk.Spinbox(size_frame, from_=5, to=25, textvariable=self.o_height, width=5, bg=COLORS['input_bg'], fg=COLORS['text'], increment=0.5).pack(side='left', padx=5)

        # 复制到PPT选项
        tk.Checkbutton(size_frame, text="复制到PPT", variable=self.o_copy_to_ppt, bg=COLORS['bg_dark'], fg=COLORS['text'], selectcolor=COLORS['bg_medium']).pack(side='left', padx=(20,0))

        self.plot_frame.columnconfigure(1, weight=1)

        # 按钮放最后一排
        btn_frame = tk.Frame(frame, bg=COLORS['bg_dark'])
        btn_frame.grid(row=4, column=0, columnspan=4, pady=15, sticky='ew')
        self._setup_dnd(btn_frame)

        tk.Button(btn_frame, text="📋 仅提取数据", command=self.run_extract_only, bg=COLORS['accent'], fg=COLORS['button_fg'], font=("微软雅黑", 10, "bold"), relief='flat', cursor='hand2').pack(side='left', expand=True, fill='x', padx=2, ipady=10)
        tk.Button(btn_frame, text="📈 仅绘图", command=self.run_plot_only, bg=COLORS['success'], fg=COLORS['button_fg'], font=("微软雅黑", 10, "bold"), relief='flat', cursor='hand2').pack(side='left', expand=True, fill='x', padx=2, ipady=10)
        
        frame.columnconfigure(1, weight=1)

    def _setup_dnd(self, widget):
        """设置拖拽"""
        def on_drop(event):
            data = event.data
            if '{' in data:
                import re
                paths = re.findall(r'\{([^}]+)\}', data)
                path = paths[0] if paths else data.strip('{}')
            else:
                path = data.split()[0] if data.split() else data
            self.v_tensile_src.set(path)
        
        def do_register():
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', on_drop)
            except Exception as e:
                print(f"拖拽注册失败: {e}")
        
        widget.after(100, do_register)

    def browse_template(self):
        p = filedialog.askopenfilename(initialdir="C:/Users/deity/Documents/OriginLab/User Files", filetypes=[("Origin Template", "*.otpu *.otp")])
        if p:
            self.o_template.set(p)
            config_manager.set_template('tensile_template', p)

    def run_extract_only(self):
        src = self.v_tensile_src.get()
        if not src: return messagebox.showwarning("提示", "请先选择数据文件")

        pptx = resource_path("拉伸模板.pptx")
        if not os.path.exists(pptx): return messagebox.showerror("错误", "未找到模板文件")

        folder = os.path.dirname(src)
        fname = os.path.splitext(os.path.basename(src))[0]
        out = get_unique_path(os.path.join(folder, f"拉伸报告_{fname}.pptx"))

        try:
            msg = tensile_processor.generate_report(src, pptx, out, self.v_include_ag.get())
            if msg and "错误" not in msg:
                messagebox.showinfo("成功", msg)
                os.startfile(out)
            else:
                messagebox.showerror("失败", msg)
        except Exception as e:
            messagebox.showerror("异常", str(e))

    def run_plot_only(self):
        src = self.v_tensile_src.get()
        if not src: return messagebox.showwarning("提示", "请先选择数据文件")

        # 检查Origin连接
        success, err = origin_processor.init_origin()
        if not success:
            return messagebox.showerror("Origin连接失败", err)

        copy_to_ppt = self.o_copy_to_ppt.get()
        if copy_to_ppt:
            messagebox.showwarning("注意", "绘图期间请勿操作键盘鼠标！\n点击确定开始绘图...")

        try:
            msg = origin_processor.plot_tensile_to_ppt(
                src,
                self.o_template.get() or None,
                self.o_lines.get(),
                self.o_swap_xy.get(),
                width_cm=self.o_width.get(),
                height_cm=self.o_height.get(),
                copy_to_ppt=copy_to_ppt
            )
            messagebox.showinfo("完成", msg)
            # 如果复制到PPT，打开生成的PPT
            if copy_to_ppt:
                folder = os.path.dirname(src)
                fname = os.path.splitext(os.path.basename(src))[0]
                ppt_path = os.path.join(folder, f"拉伸曲线_{fname}.pptx")
                if os.path.exists(ppt_path):
                    os.startfile(ppt_path)
        except Exception as e:
            messagebox.showerror("错误", str(e))
