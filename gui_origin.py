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

        self.o_template_path = tk.StringVar(value=config_manager.get_template('phase_template'))

        label_hint = tk.Label(main_frame, text="拖拽CSV文件到下方区域（支持多文件）| 💡 可拖拽到整个界面任意位置",
                bg=COLORS['bg_dark'], fg=COLORS['text'],
                font=('微软雅黑', 10))
        label_hint.grid(row=0, column=0, columnspan=3, sticky='w', pady=(0,5))

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

        # 注册拖拽 - 扩展到整个界面
        self._setup_dnd(self)
        self._setup_dnd(main_frame)
        self._setup_dnd(label_hint)

        btn_frame = tk.Frame(main_frame, bg=COLORS['bg_dark'])
        btn_frame.grid(row=2, column=0, columnspan=3, sticky='ew', pady=5)
        self._setup_dnd(btn_frame)

        tk.Button(btn_frame, text="添加文件", command=self.add_files,
                 bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').pack(side='left', padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_files,
                 bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').pack(side='left', padx=5)

        label_template = tk.Label(main_frame, text="绘图模板:", bg=COLORS['bg_dark'], fg=COLORS['text'])
        label_template.grid(row=3, column=0, sticky='w', pady=10)
        tk.Entry(main_frame, textvariable=self.o_template_path, width=30, bg=COLORS['input_bg'], fg=COLORS['text']).grid(row=3, column=1, sticky='ew', padx=5)
        tk.Button(main_frame, text="选择", command=self.browse_template, bg=COLORS['bg_light'], fg=COLORS['text'], relief='flat').grid(row=3, column=2)

        # 图片尺寸选项
        self.o_width = tk.DoubleVar(value=11.0)
        self.o_height = tk.DoubleVar(value=8.8)
        self.o_copy_to_ppt = tk.BooleanVar(value=False)  # 默认不复制到PPT

        size_frame = tk.Frame(main_frame, bg=COLORS['bg_dark'])
        size_frame.grid(row=4, column=0, columnspan=3, sticky='w', pady=5)
        self._setup_dnd(size_frame)

        label_width = tk.Label(size_frame, text="图片宽(cm):", bg=COLORS['bg_dark'], fg=COLORS['text'])
        label_width.pack(side='left')
        self._setup_dnd(label_width)

        tk.Spinbox(size_frame, from_=5, to=30, textvariable=self.o_width, width=5, bg=COLORS['input_bg'], fg=COLORS['text'], increment=0.5).pack(side='left', padx=(5,15))

        label_height = tk.Label(size_frame, text="图片高(cm):", bg=COLORS['bg_dark'], fg=COLORS['text'])
        label_height.pack(side='left')
        self._setup_dnd(label_height)

        tk.Spinbox(size_frame, from_=5, to=25, textvariable=self.o_height, width=5, bg=COLORS['input_bg'], fg=COLORS['text'], increment=0.5).pack(side='left', padx=5)

        # 复制到PPT选项
        tk.Checkbutton(size_frame, text="复制到PPT", variable=self.o_copy_to_ppt, bg=COLORS['bg_dark'], fg=COLORS['text'], selectcolor=COLORS['bg_medium']).pack(side='left', padx=(20,0))

        btn_plot = tk.Button(main_frame, text="🚀 开始绘图", command=self.run_plot,
                 bg=COLORS['success'], fg=COLORS['button_fg'], font=("微软雅黑", 12, "bold"),
                 relief='flat', cursor='hand2')
        btn_plot.grid(row=5, column=0, columnspan=3, sticky='ew', ipady=10, pady=15)
        self._setup_dnd(btn_plot)

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
