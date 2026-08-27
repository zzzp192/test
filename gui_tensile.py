import tkinter as tk
from tkinter import messagebox, filedialog
import os
import tensile_processor
import origin_processor
import matplotlib_ppt
import config_manager
from gui_shared import (
    resource_path, browse_file, get_unique_path, COLORS,
    create_button, create_checkbutton, create_entry, create_field_label,
    create_page, create_radiobutton, create_section, create_spinbox
)
from tkinterdnd2 import DND_FILES

class TensileFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.setup_ui()

    def setup_ui(self):
        for widget in self.winfo_children():
            widget.destroy()
        self.configure(bg=COLORS['bg_dark'])

        page = create_page(self, scrollable=True)
        self.v_tensile_src = tk.StringVar()
        self.v_elongation_mode = tk.StringVar(value="ag")

        data_section = create_section(page, "拉伸实验报告")
        create_field_label(data_section, "原始数据文件").grid(row=0, column=0, columnspan=2, sticky='w')
        entry = create_entry(data_section, self.v_tensile_src, width=54, mono=True)
        entry.grid(row=1, column=0, padx=(0, 10), pady=(6, 0), sticky='ew', ipady=9)
        create_button(
            data_section,
            "浏览文件",
            lambda: browse_file(self.v_tensile_src, [("Data Files", "*.xlsx *.xls *.csv *.docx")]),
            "secondary",
        ).grid(row=1, column=1, sticky='ew', pady=(6, 0))
        data_section.columnconfigure(0, weight=1)

        self._setup_dnd(self)
        self._setup_dnd(page)
        self._setup_dnd(data_section)
        self._setup_dnd(entry)

        option_section = create_section(page, "报告字段")
        opt_frame = tk.Frame(option_section, bg=COLORS['bg_medium'])
        opt_frame.pack(fill='x')
        for value, text in [
            ("a_only", "仅 A（断后伸长率）"),
            ("ag", "A + Ag（最大力延伸率）"),
            ("at", "A + At（断裂总延伸率）"),
        ]:
            create_radiobutton(opt_frame, text, self.v_elongation_mode, value).pack(side="left", padx=(0, 18))
        self._setup_dnd(option_section)

        self.plot_frame = create_section(
            page,
            "绘图选项",
            "Origin 与一键 PPT 共用曲线分组和 XY 调换规则；一键 PPT 固定生成 12×16 cm（高×宽）透明图。",
        )
        self._setup_dnd(self.plot_frame)

        self.o_template = tk.StringVar(value=config_manager.get_template('tensile_template'))
        self.o_lines = tk.IntVar(value=12)
        self.o_swap_xy = tk.BooleanVar(value=True)

        create_field_label(self.plot_frame, "模板").grid(row=0, column=0, sticky='w')
        create_entry(self.plot_frame, self.o_template, width=28).grid(row=1, column=0, columnspan=2, sticky='ew', padx=(0, 10), pady=(6, 12), ipady=8)
        create_button(self.plot_frame, "选择", self.browse_template, "secondary").grid(row=1, column=2, sticky='ew', pady=(6, 12))
        create_field_label(self.plot_frame, "每图曲线数").grid(row=0, column=3, sticky='w', padx=(18, 0))
        create_spinbox(self.plot_frame, 1, 50, self.o_lines, width=6).grid(row=1, column=3, sticky='w', padx=(18, 0), pady=(6, 12), ipady=5)
        create_checkbutton(self.plot_frame, "调换 XY 列", self.o_swap_xy).grid(row=1, column=4, sticky='w', padx=(18, 0), pady=(6, 12))

        self.plot_frame.columnconfigure(1, weight=1)

        btn_frame = tk.Frame(page, bg=COLORS['bg_dark'])
        btn_frame.pack(fill='x', pady=(2, 0))
        self._setup_dnd(btn_frame)

        create_button(btn_frame, "仅提取数据", self.run_extract_only, "primary").pack(side='left', expand=True, fill='x', padx=(0, 4))
        create_button(btn_frame, "仅origin绘图", self.run_plot_only, "secondary").pack(side='left', expand=True, fill='x', padx=4)
        create_button(btn_frame, "一键PPT（非origin出图）", self.run_one_click_ppt, "cta").pack(side='left', expand=True, fill='x', padx=(4, 0))

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
            msg = tensile_processor.generate_report(src, pptx, out, elongation_mode=self.v_elongation_mode.get())
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

        try:
            msg = origin_processor.plot_tensile_in_origin(
                src,
                self.o_template.get() or None,
                self.o_lines.get(),
                self.o_swap_xy.get(),
            )
            messagebox.showinfo("完成", msg)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def run_one_click_ppt(self):
        src = self.v_tensile_src.get()
        if not src:
            return messagebox.showwarning("提示", "请先选择数据文件")

        template = resource_path("拉伸模板.pptx")
        if not os.path.exists(template):
            return messagebox.showerror("错误", "未找到模板文件")

        folder = os.path.dirname(src)
        fname = os.path.splitext(os.path.basename(src))[0]
        output = get_unique_path(os.path.join(folder, f"拉伸报告_{fname}.pptx"))
        try:
            msg = matplotlib_ppt.create_tensile_one_click_ppt(
                src,
                template,
                output,
                lines_per_graph=self.o_lines.get(),
                swap_xy=self.o_swap_xy.get(),
                elongation_mode=self.v_elongation_mode.get(),
            )
            messagebox.showinfo("完成", msg)
            os.startfile(output)
        except Exception as e:
            messagebox.showerror("错误", str(e))
