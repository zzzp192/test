import tkinter as tk
from tkinter import messagebox, filedialog
import os
import vda_processor
import origin_processor
from gui_shared import resource_path, browse_file, get_unique_path, COLORS
from tkinterdnd2 import DND_FILES

class VDAFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.setup_ui()

    def setup_ui(self):
        for widget in self.winfo_children():
            widget.destroy()
        self.configure(bg=COLORS['bg_dark'])
        
        frame = tk.LabelFrame(self, text="📐 VDA弯曲报告生成", padx=20, pady=20,
                             bg=COLORS['bg_dark'], fg=COLORS['accent'], font=('微软雅黑', 11, 'bold'))
        frame.pack(fill="x", padx=25, pady=25)
        
        self.v_vda_src = tk.StringVar()
        
        # 文件选择
        tk.Label(frame, text="选择原始数据文件 (Excel):", bg=COLORS['bg_dark'], fg=COLORS['text'], font=('微软雅黑', 10)).grid(row=0, column=0, columnspan=4, sticky='w', pady=(0,10))
        
        entry = tk.Entry(frame, textvariable=self.v_vda_src, width=50, font=('Consolas', 10), bg=COLORS['input_bg'], fg=COLORS['text'],
                        insertbackground=COLORS['accent'], relief='flat', highlightthickness=1, highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'])
        entry.grid(row=1, column=0, columnspan=3, padx=(0,10), pady=5, sticky='ew', ipady=8)
        
        tk.Button(frame, text="📂 浏览", command=lambda: browse_file(self.v_vda_src, [("Excel Files", "*.xlsx *.xls")]),
                 bg=COLORS['bg_light'], fg=COLORS['text'], font=('微软雅黑', 9), relief='flat', cursor='hand2', padx=15).grid(row=1, column=3, padx=5, ipady=5)
        
        # 拖拽区域
        self.drop_zone = tk.Label(frame, text="⬇️  拖拽文件到此处  ⬇️", bg=COLORS['bg_medium'], fg=COLORS['text_dim'], font=('微软雅黑', 10), height=2, relief="flat", cursor="hand2")
        self.drop_zone.grid(row=2, column=0, columnspan=4, sticky="ew", pady=10)
        
        # 注册拖拽
        self._setup_dnd(entry)
        self._setup_dnd(self.drop_zone)

        # 绘图选项
        self.plot_frame = tk.LabelFrame(frame, text="绘图选项", padx=10, pady=10, bg=COLORS['bg_dark'], fg=COLORS['text_dim'], font=('微软雅黑', 9))
        self.plot_frame.grid(row=3, column=0, columnspan=4, sticky='ew', pady=10)
        
        self.o_template = tk.StringVar()
        self.o_lines = tk.IntVar(value=12)
        self.o_swap_xy = tk.BooleanVar(value=True)
        
        tk.Label(self.plot_frame, text="模板:", bg=COLORS['bg_dark'], fg=COLORS['text']).grid(row=0, column=0, sticky='w')
        tk.Entry(self.plot_frame, textvariable=self.o_template, width=25, bg=COLORS['input_bg'], fg=COLORS['text']).grid(row=0, column=1, sticky='ew', padx=5)
        tk.Button(self.plot_frame, text="选择", command=self.browse_template, bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').grid(row=0, column=2)
        
        tk.Label(self.plot_frame, text="每图曲线数:", bg=COLORS['bg_dark'], fg=COLORS['text']).grid(row=0, column=3, padx=(15,5))
        tk.Spinbox(self.plot_frame, from_=1, to=50, textvariable=self.o_lines, width=5, bg=COLORS['input_bg'], fg=COLORS['text']).grid(row=0, column=4)
        
        tk.Checkbutton(self.plot_frame, text="调换XY列", variable=self.o_swap_xy, bg=COLORS['bg_dark'], fg=COLORS['text'], selectcolor=COLORS['bg_medium']).grid(row=0, column=5, padx=15)
        
        self.plot_frame.columnconfigure(1, weight=1)

        # 按钮放最后一排
        btn_frame = tk.Frame(frame, bg=COLORS['bg_dark'])
        btn_frame.grid(row=4, column=0, columnspan=4, pady=15, sticky='ew')
        
        tk.Button(btn_frame, text="📋 仅提取数据", command=self.run_extract_only, bg=COLORS['accent'], fg=COLORS['button_fg'], font=("微软雅黑", 10, "bold"), relief='flat', cursor='hand2').pack(side='left', expand=True, fill='x', padx=2, ipady=8)
        tk.Button(btn_frame, text="📈 仅绘图", command=self.run_plot_only, bg=COLORS['success'], fg=COLORS['button_fg'], font=("微软雅黑", 10, "bold"), relief='flat', cursor='hand2').pack(side='left', expand=True, fill='x', padx=2, ipady=8)
        tk.Button(btn_frame, text="⚡ 提取&绘图", command=self.run_both, bg=COLORS['warning'], fg='#333', font=("微软雅黑", 10, "bold"), relief='flat', cursor='hand2').pack(side='left', expand=True, fill='x', padx=2, ipady=8)
        
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
            self.v_vda_src.set(path)
        
        def do_register():
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', on_drop)
            except Exception as e:
                print(f"拖拽注册失败: {e}")
        
        widget.after(100, do_register)

    def browse_template(self):
        p = filedialog.askopenfilename(initialdir="C:/Users/deity/Documents/OriginLab/User Files", filetypes=[("Origin Template", "*.otpu *.otp")])
        if p: self.o_template.set(p)

    def run_extract_only(self):
        src = self.v_vda_src.get()
        if not src: return messagebox.showwarning("提示", "请先选择数据文件")
        
        pptx = resource_path("VDA弯曲角模板.pptx")
        if not os.path.exists(pptx): return messagebox.showerror("错误", "未找到模板文件")
        
        folder = os.path.dirname(src)
        fname = os.path.splitext(os.path.basename(src))[0]
        out = get_unique_path(os.path.join(folder, f"VDA报告_{fname}.pptx"))
        
        try:
            msg = vda_processor.process_vda_report(src, pptx, out)
            if msg and "错误" not in msg:
                messagebox.showinfo("成功", msg)
                os.startfile(out)
            else:
                messagebox.showerror("失败", msg)
        except Exception as e:
            messagebox.showerror("异常", str(e))

    def run_plot_only(self):
        src = self.v_vda_src.get()
        if not src: return messagebox.showwarning("提示", "请先选择数据文件")
        
        # 显示进度窗口
        progress = tk.Toplevel(self)
        progress.title("处理中")
        progress.geometry("300x100")
        progress.transient(self)
        tk.Label(progress, text="正在处理Origin绘图...\n请稍候", font=('微软雅黑', 11)).pack(expand=True)
        progress.update()
        
        try:
            msg = origin_processor.plot_vda_to_ppt(src, self.o_template.get() or None, self.o_lines.get(), self.o_swap_xy.get())
            progress.destroy()
            messagebox.showinfo("完成", msg)
        except Exception as e:
            progress.destroy()
            messagebox.showerror("错误", str(e))

    def run_both(self):
        src = self.v_vda_src.get()
        if not src: return messagebox.showwarning("提示", "请先选择数据文件")
        
        pptx = resource_path("VDA弯曲角模板.pptx")
        if not os.path.exists(pptx): return messagebox.showerror("错误", "未找到模板文件")
        
        folder = os.path.dirname(src)
        fname = os.path.splitext(os.path.basename(src))[0]
        out = get_unique_path(os.path.join(folder, f"VDA报告_{fname}.pptx"))
        
        try:
            msg1 = vda_processor.process_vda_report(src, pptx, out)
            msg2 = origin_processor.plot_vda_to_ppt(src, self.o_template.get() or None, self.o_lines.get(), self.o_swap_xy.get())
            messagebox.showinfo("完成", f"{msg1}\n\n{msg2}")
            os.startfile(out)
        except Exception as e:
            messagebox.showerror("异常", str(e))
