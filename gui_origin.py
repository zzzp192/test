#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
育材堂报告助手 V3.7 - 相变点绘图模块

软件名称：育材堂报告助手
版本号：V3.7
开发单位：育材堂
开发完成日期：2024年

模块功能：
    提供相变点数据的批量绘图功能。

主要功能：
    - 支持拖拽多个CSV文件
    - 使用Temperature作为X轴，Change作为Y轴
    - 每个文件生成一张图
    - 支持Origin模板选择
    - 自动导出为PPT（每页一张OLE图形）

Copyright (c) 2024 育材堂. All rights reserved.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
from tkinterdnd2 import DND_FILES
import origin_processor
import config_manager
from gui_shared import (
    COLORS, create_button, create_checkbutton, create_entry, create_field_label,
    create_page, create_section, create_spinbox
)

class OriginFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=COLORS['bg_dark'])
        self.file_list = []
        self.setup_ui()

    def setup_ui(self):
        for widget in self.winfo_children():
            widget.destroy()
        self.configure(bg=COLORS['bg_dark'])

        page = create_page(self)
        self.o_template_path = tk.StringVar(value=config_manager.get_template('phase_template'))

        files_section = create_section(page, "相变点绘图", "添加一个或多个 CSV 文件，批量生成相变点曲线。")

        self.drop_zone = tk.Listbox(
            files_section,
            height=8,
            bg=COLORS['input_bg'],
            fg=COLORS['text'],
            selectbackground=COLORS['accent'],
            selectforeground=COLORS['button_fg'],
            highlightthickness=1,
            highlightbackground=COLORS['border'],
            highlightcolor=COLORS['accent'],
            relief='flat',
            selectmode=tk.EXTENDED,
            font=('Consolas', 9),
        )
        self.drop_zone.grid(row=0, column=0, columnspan=2, sticky='nsew')

        def do_register():
            try:
                self.drop_zone.drop_target_register(DND_FILES)
                self.drop_zone.dnd_bind('<<Drop>>', self.on_drop)
            except Exception as e:
                print(f"拖拽注册失败: {e}")

        self.drop_zone.after(100, do_register)

        self._setup_dnd(self)
        self._setup_dnd(page)
        self._setup_dnd(files_section)

        btn_frame = tk.Frame(files_section, bg=COLORS['bg_medium'])
        btn_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(12, 0))
        self._setup_dnd(btn_frame)

        create_button(btn_frame, "添加文件", self.add_files, "secondary").pack(side='left', padx=(0, 8))
        create_button(btn_frame, "清空列表", self.clear_files, "secondary").pack(side='left')
        files_section.columnconfigure(0, weight=1)

        option_section = create_section(page, "绘图选项", "设置 Origin 模板、图片尺寸和 PPT 输出方式。")
        create_field_label(option_section, "绘图模板").grid(row=0, column=0, sticky='w')
        create_entry(option_section, self.o_template_path, width=34).grid(row=1, column=0, sticky='ew', padx=(0, 10), pady=(6, 12), ipady=8)
        create_button(option_section, "选择", self.browse_template, "secondary").grid(row=1, column=1, sticky='ew', pady=(6, 12))

        self.o_width = tk.DoubleVar(value=11.0)
        self.o_height = tk.DoubleVar(value=8.8)
        self.o_copy_to_ppt = tk.BooleanVar(value=False)  # 默认不复制到PPT

        size_frame = tk.Frame(option_section, bg=COLORS['bg_medium'])
        size_frame.grid(row=2, column=0, columnspan=2, sticky='w')
        self._setup_dnd(size_frame)

        create_field_label(size_frame, "图片宽(cm)").pack(side='left')
        create_spinbox(size_frame, 5, 30, self.o_width, width=6, increment=0.5).pack(side='left', padx=(8, 18), ipady=4)
        create_field_label(size_frame, "图片高(cm)").pack(side='left')
        create_spinbox(size_frame, 5, 25, self.o_height, width=6, increment=0.5).pack(side='left', padx=(8, 18), ipady=4)
        create_checkbutton(size_frame, "复制到 PPT", self.o_copy_to_ppt).pack(side='left', padx=(8, 0))
        option_section.columnconfigure(0, weight=1)

        btn_plot = create_button(page, "开始绘图", self.run_plot, "cta")
        btn_plot.pack(fill='x', pady=(2, 0))
        self._setup_dnd(btn_plot)

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
            config_manager.set_template('phase_template', p)

    def run_plot(self):
        if not self.file_list:
            return messagebox.showwarning("提示", "请先添加CSV文件")
        
        # 检查Origin连接
        success, err = origin_processor.init_origin()
        if not success:
            return messagebox.showerror("Origin连接失败", err)
        
        copy_to_ppt = self.o_copy_to_ppt.get()
        if copy_to_ppt:
            messagebox.showwarning("注意", "绘图期间请勿操作键盘鼠标！\n点击确定开始绘图...")
        
        tmpl = self.o_template_path.get() or None
        try:
            result = origin_processor.plot_phase_change(
                self.file_list, tmpl, 
                width_cm=self.o_width.get(), 
                height_cm=self.o_height.get(),
                copy_to_ppt=copy_to_ppt
            )
            
            if copy_to_ppt:
                ppt_path, opju_path, count = result
                messagebox.showinfo("完成", f"成功！已处理 {count} 个文件\nPPT: {ppt_path}\nOrigin项目: {opju_path}")
                os.startfile(ppt_path)
            else:
                opju_path, count = result
                messagebox.showinfo("完成", f"成功！已在Origin中创建 {count} 张图表\nOrigin项目: {opju_path}")
        except Exception as e:
            import traceback
            messagebox.showerror("错误", f"{e}\n{traceback.format_exc()}")

    def set_data_source(self, path):
        pass

    def _setup_dnd(self, widget):
        """设置拖拽"""
        def on_drop(event):
            files = self.parse_drop_data(event.data)
            for f in files:
                if f.endswith('.csv') and f not in self.file_list:
                    self.file_list.append(f)
                    self.drop_zone.insert(tk.END, os.path.basename(f))

        def do_register():
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', on_drop)
            except Exception as e:
                print(f"拖拽注册失败: {e}")

        widget.after(100, do_register)
