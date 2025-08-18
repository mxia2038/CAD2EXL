#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P&ID管道数据提取工具 - GUI版本 (响应式设计)
从P&ID图纸中提取管道号并生成Excel报告
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import re
import pandas as pd
import logging
import os
import sys
import json
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageTk

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PIDExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("P&ID管道数据提取工具")
        
        # 设置固定窗口尺寸适配笔记本
        self.root.geometry("850x650")
        self.root.minsize(850, 650)
        
        # 文件路径变量
        self.dwg_file = tk.StringVar()
        self.code_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        # 支持的项目类型列表（便于未来扩展）
        self.SUPPORTED_PROJECT_TYPES = ["巨化项目", "乌兹项目"]
        
        # 项目类型变量
        self.project_type = tk.StringVar()
        self.project_type.set(self.SUPPORTED_PROJECT_TYPES[0])  # 默认选择第一个项目类型
        
        # 设置默认值
        self.code_file.set("test/code.xlsx")
        self.output_file.set("pipeline_data.xlsx")
        
        # 配置文件路径
        self.config_file = Path.home() / ".pid_extractor_config.json"
        
        # 加载最近使用的文件
        self.load_recent_files()
        
        self.create_widgets()
        # 延迟设置拖拽，等待窗口完全初始化
        self.root.after(100, self.setup_drag_drop)
        
    def create_widgets(self):
        # 配置根窗口
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 创建垂直PanedWindow：上部可滚动设置区 + 下部固定结果区
        # 改用tk.PanedWindow以支持minsize参数
        self.paned_window = tk.PanedWindow(self.root, orient=tk.VERTICAL, sashrelief=tk.RAISED)
        self.paned_window.pack(fill=tk.BOTH, expand=True)
        
        # 上部：可滚动的设置区域
        self.setup_scroll_container = ttk.Frame(self.paned_window)
        # 配置滚动容器的列权重
        self.setup_scroll_container.columnconfigure(0, weight=1)
        self.setup_scroll_container.rowconfigure(0, weight=1)
        
        self.setup_canvas = tk.Canvas(self.setup_scroll_container, highlightthickness=0)
        self.setup_scrollbar = ttk.Scrollbar(self.setup_scroll_container, orient=tk.VERTICAL, command=self.setup_canvas.yview)
        self.setup_inner_frame = ttk.Frame(self.setup_canvas)
        
        # 配置滚动 - 修复右侧空白问题
        # 把 inner frame 嵌进 canvas，拿到 window_id 方便后面调整宽度
        self.inner_window_id = self.setup_canvas.create_window((0, 0), window=self.setup_inner_frame, anchor="nw")
        self.setup_canvas.configure(yscrollcommand=self.setup_scrollbar.set)
        
        # 消除右侧空白：让 inner_frame 宽度跟随 canvas
        def _sync_scroll_region(event):
            # 更新可滚动范围
            self.setup_canvas.configure(scrollregion=self.setup_canvas.bbox("all"))
        
        def _expand_inner_width(event):
            # 让 inner frame 始终跟 canvas 同宽
            self.setup_canvas.itemconfigure(self.inner_window_id, width=event.width)
        
        self.setup_inner_frame.bind("<Configure>", _sync_scroll_region)
        self.setup_canvas.bind("<Configure>", _expand_inner_width)
        
        # 使用grid布局替代pack来更好控制
        self.setup_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.setup_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 下部：固定的结果区域
        self.results_container = ttk.Frame(self.paned_window)
        
        # 添加到PanedWindow - 使用tk.PanedWindow的minsize参数
        self.paned_window.add(self.setup_scroll_container, minsize=450)
        self.paned_window.add(self.results_container, minsize=60)
        
        # 修复sash位置设置时机 - 按o3建议使用after_idle
        def _place_sash():
            # 使用paned_window自身的高度更可靠
            total = self.paned_window.winfo_height()
            if total:  # 防0
                # 设置区域占90%，大幅减少结果区域大小
                desired = int(total * 0.90)
                self.paned_window.sash_place(0, desired, 0)
        
        # 用after_idle确保窗口完全初始化后再调用
        self.root.after_idle(_place_sash)
        
        # 绑定鼠标滚轮事件
        self.bind_mousewheel()
        
        # 创建设置区域内容
        self.create_setup_widgets()
        # 创建结果区域内容  
        self.create_results_widgets()
        
        # 初始滚动提示（显示有内容在下方）
        self.root.after(200, lambda: self.setup_canvas.yview_moveto(0.001))
    
    def on_project_type_changed(self, *args):
        """项目类型变化时的处理"""
        self.update_format_example()
    
    def update_format_example(self):
        """更新格式示例显示"""
        project_type = self.project_type.get()
        
        # 项目格式示例映射
        format_examples = {
            "巨化项目": "示例: 4101BRR-02457-200-03CBMB1-H",
            "乌兹项目": "示例: PA-2001002A-100-C1C-N"
        }
        
        example_text = format_examples.get(project_type, "")
        self.format_example_label.config(text=example_text)
    
    
    def bind_mousewheel(self):
        """绑定鼠标滚轮事件"""
        def on_mousewheel(event):
            self.setup_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def bind_to_mousewheel(event):
            self.setup_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def unbind_from_mousewheel(event):
            self.setup_canvas.unbind_all("<MouseWheel>")
        
        self.setup_canvas.bind('<Enter>', bind_to_mousewheel)
        self.setup_canvas.bind('<Leave>', unbind_from_mousewheel)
    
    
    def create_setup_widgets(self):
        """创建设置区域的所有widgets"""
        # 配置主容器
        self.setup_inner_frame.columnconfigure(0, weight=1)
        
        # 添加公司Logo
        self.setup_logo()
        
        # 标题
        title_label = ttk.Label(self.setup_inner_frame, text="P&ID管道数据提取工具", 
                               font=("Microsoft YaHei", 14, "bold"))
        title_label.grid(row=1, column=0, columnspan=1, pady=(0, 10))
        
        # ========== 项目类型选择区域 ==========
        project_section = ttk.LabelFrame(self.setup_inner_frame, text="🏗️ 项目类型", padding="8")
        project_section.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        project_section.columnconfigure(0, weight=1)
        
        # 项目类型选择器
        project_frame = ttk.Frame(project_section)
        project_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        project_frame.columnconfigure(1, weight=1)
        
        ttk.Label(project_frame, text="选择项目编号标准:", font=("Microsoft YaHei", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # 下拉选择器
        project_combobox = ttk.Combobox(project_frame, textvariable=self.project_type, 
                                       values=self.SUPPORTED_PROJECT_TYPES,
                                       state="readonly", width=15)
        project_combobox.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        
        # 格式示例标签（动态显示）
        self.format_example_label = ttk.Label(project_frame, text="", 
                                             font=("Microsoft YaHei", 9), foreground="gray")
        self.format_example_label.grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        
        # 绑定项目类型变化事件
        project_combobox.bind('<<ComboboxSelected>>', self.on_project_type_changed)
        self.project_type.trace('w', self.on_project_type_changed)
        
        # 初始化显示格式示例
        self.update_format_example()
        
        # ========== 输入文件区域（两列布局）==========
        input_section = ttk.LabelFrame(self.setup_inner_frame, text="📁 输入文件", padding="8")
        input_section.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        
        # 配置两列权重
        input_section.columnconfigure(0, weight=1)
        input_section.columnconfigure(1, weight=1)
        
        # ========== 左列：DWG文件拖放区域 ==========
        dwg_frame = ttk.Frame(input_section)
        dwg_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 4))
        dwg_frame.columnconfigure(0, weight=1)
        
        dwg_label = ttk.Label(dwg_frame, text="DWG源文件", font=("Microsoft YaHei", 10, "bold"))
        dwg_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 3))
        
        dwg_hint = ttk.Label(dwg_frame, text="拖拽 .dwg 文件到下方", 
                            font=("Microsoft YaHei", 8), foreground="gray")
        dwg_hint.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        # DWG拖放框 - 增加高度
        self.dwg_drop_frame = tk.Frame(dwg_frame, relief="solid", borderwidth=2, 
                                      bg="#f8f9fa", height=80)
        self.dwg_drop_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        self.dwg_drop_frame.pack_propagate(False)
        
        dwg_icon_label = tk.Label(self.dwg_drop_frame, text="📋", font=("Microsoft YaHei", 14), 
                                 bg="#f8f9fa", fg="#6c757d")
        dwg_icon_label.place(relx=0.5, rely=0.3, anchor="center")
        
        dwg_text_label = tk.Label(self.dwg_drop_frame, text="拖拽 DWG 文件", 
                                 font=("Microsoft YaHei", 8), bg="#f8f9fa", fg="#6c757d")
        dwg_text_label.place(relx=0.5, rely=0.7, anchor="center")
        
        # DWG文件显示和按钮
        dwg_control_frame = ttk.Frame(dwg_frame)
        dwg_control_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        dwg_control_frame.columnconfigure(0, weight=1)
        
        self.dwg_entry = ttk.Entry(dwg_control_frame, textvariable=self.dwg_file, width=30)
        self.dwg_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(dwg_control_frame, text="浏览", command=self.select_dwg_file).grid(row=0, column=1)
        
        # ========== 右列：介质代码文件拖放区域 ==========
        code_frame = ttk.Frame(input_section)
        code_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(4, 0))
        code_frame.columnconfigure(0, weight=1)
        
        code_label = ttk.Label(code_frame, text="介质代码数据文件", font=("Microsoft YaHei", 10, "bold"))
        code_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 3))
        
        code_hint = ttk.Label(code_frame, text="拖拽 .xlsx 文件到下方", 
                             font=("Microsoft YaHei", 8), foreground="gray")
        code_hint.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        # 介质代码拖放框 - 增加高度
        self.code_drop_frame = tk.Frame(code_frame, relief="solid", borderwidth=2, 
                                       bg="#f8f9fa", height=80)
        self.code_drop_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 6))
        self.code_drop_frame.pack_propagate(False)
        
        code_icon_label = tk.Label(self.code_drop_frame, text="📊", font=("Microsoft YaHei", 14), 
                                  bg="#f8f9fa", fg="#6c757d")
        code_icon_label.place(relx=0.5, rely=0.3, anchor="center")
        
        code_text_label = tk.Label(self.code_drop_frame, text="拖拽 Excel 文件", 
                                  font=("Microsoft YaHei", 8), bg="#f8f9fa", fg="#6c757d")
        code_text_label.place(relx=0.5, rely=0.7, anchor="center")
        
        # 介质代码文件显示和按钮
        code_control_frame = ttk.Frame(code_frame)
        code_control_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        code_control_frame.columnconfigure(0, weight=1)
        
        self.code_entry = ttk.Entry(code_control_frame, textvariable=self.code_file, width=30)
        self.code_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(code_control_frame, text="浏览", command=self.select_code_file).grid(row=0, column=1)
        
        # ========== 输出设置区域（完整宽度）==========
        output_section = ttk.LabelFrame(self.setup_inner_frame, text="💾 输出设置", padding="8")
        output_section.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(8, 0))
        output_section.columnconfigure(0, weight=1)
        
        output_label = ttk.Label(output_section, text="输出文件路径", font=("Microsoft YaHei", 10, "bold"))
        output_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 3))
        
        # 输出文件控制
        output_control_frame = ttk.Frame(output_section)
        output_control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 4))
        output_control_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_control_frame, textvariable=self.output_file, width=60)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 8))
        
        ttk.Button(output_control_frame, text="选择路径", command=self.select_output_file).grid(row=0, column=1)
        
        # ========== 操作按钮区域 ==========
        action_frame = ttk.Frame(self.setup_inner_frame)
        action_frame.grid(row=5, column=0, pady=(10, 0))
        
        # 提取按钮
        extract_button = ttk.Button(action_frame, text="🚀 开始提取数据", command=self.start_extraction,
                                   style="Accent.TButton")
        extract_button.grid(row=0, column=0, ipadx=20, ipady=5)
        
        # ========== 进度条和状态区域 ==========
        status_frame = ttk.Frame(self.setup_inner_frame)
        status_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(8, 8))
        status_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 4))
        
        # 状态标签
        self.status_label = ttk.Label(status_frame, text="请选择DWG文件开始提取")
        self.status_label.grid(row=1, column=0)
    
    def create_results_widgets(self):
        """创建结果区域的widgets"""
        # 配置结果容器的列权重 - 修复右侧空白问题
        self.results_container.columnconfigure(0, weight=1)
        
        # 结果显示区域标题
        result_title = ttk.Label(self.results_container, text="📋 提取结果", 
                                font=("Microsoft YaHei", 12, "bold"))
        result_title.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        # 结果文本框框架
        result_frame = ttk.Frame(self.results_container)
        result_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 结果文本框（减小高度，给设置区域更多空间）
        self.result_text = tk.Text(result_frame, height=3, width=70, wrap="word",
                                  font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置结果区域的网格权重
        self.results_container.rowconfigure(0, weight=0)  # 标题
        self.results_container.rowconfigure(1, weight=1)  # 结果文本框 - 扩展
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
    def setup_logo(self):
        """设置公司Logo"""
        try:
            # 获取logo路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe文件
                base_path = Path(sys._MEIPASS)
            else:
                # 如果是源代码运行
                base_path = Path(__file__).parent
            
            logo_path = base_path / "fig" / "logo.jpg"
            if logo_path.exists():
                # 加载和调整logo大小
                logo_image = Image.open(logo_path)
                # 获取原始尺寸
                original_width, original_height = logo_image.size
                # 计算合适的宽高比，保持原始比例
                target_height = 50  # 减小logo高度
                aspect_ratio = original_width / original_height
                target_width = int(target_height * aspect_ratio)
                logo_image = logo_image.resize((target_width, target_height), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                
                # 显示logo (在标题前)
                logo_label = tk.Label(self.setup_inner_frame, image=self.logo_photo)
                logo_label.grid(row=0, column=0, columnspan=1, pady=(0, 6))
                
        except Exception as e:
            print(f"无法加载logo: {e}")
    
    def load_recent_files(self):
        """加载最近使用的文件"""
        self.recent_files = {
            'dwg': [],
            'code': [],
            'output': []
        }
        
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.recent_files = config.get('recent_files', self.recent_files)
        except Exception as e:
            print(f"无法加载配置文件: {e}")
    
    def save_recent_files(self):
        """保存最近使用的文件"""
        try:
            config = {'recent_files': self.recent_files}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"无法保存配置文件: {e}")
    
    def add_recent_file(self, file_type, file_path):
        """添加到最近使用的文件列表"""
        if file_path and file_path not in self.recent_files[file_type]:
            self.recent_files[file_type].insert(0, file_path)
            # 只保留最近5个文件
            self.recent_files[file_type] = self.recent_files[file_type][:5]
            self.save_recent_files()
    
    def setup_drag_drop(self):
        """设置拖拽功能"""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # 为每个拖拽区域设置独立的处理函数
            def create_drop_handler(target_var, file_type, valid_extensions):
                def on_drop(event):
                    # 处理拖拽的文件路径 - 支持含空格的文件名
                    file_data = event.data.strip()
                    
                    # 移除外层大括号（如果存在）
                    if file_data.startswith('{') and file_data.endswith('}'):
                        file_path = file_data[1:-1]
                    else:
                        file_path = file_data
                    
                    # 如果还是有多个文件，取第一个
                    if '\n' in file_path:
                        file_path = file_path.split('\n')[0]
                    
                    file_path = file_path.strip()
                    
                    if file_path:
                        # 检查文件扩展名
                        if any(file_path.lower().endswith(ext) for ext in valid_extensions):
                            target_var.set(file_path)
                            self.add_recent_file(file_type, file_path)
                            self.show_drop_feedback(event.widget, "success")
                        else:
                            self.show_drop_feedback(event.widget, "error")
                return on_drop
            
            def on_drag_enter(event):
                self.show_drop_feedback(event.widget, "hover")
            
            def on_drag_leave(event):
                self.show_drop_feedback(event.widget, "normal")
            
            # 为DWG文件区域设置拖拽和点击
            self.dwg_drop_frame.drop_target_register(DND_FILES)
            self.dwg_drop_frame.dnd_bind('<<Drop>>', create_drop_handler(self.dwg_file, 'dwg', ['.dwg']))
            self.dwg_drop_frame.dnd_bind('<<DragEnter>>', on_drag_enter)
            self.dwg_drop_frame.dnd_bind('<<DragLeave>>', on_drag_leave)
            # 点击拖拽区域也能选择文件
            self.dwg_drop_frame.bind('<Button-1>', lambda e: self.select_dwg_file())
            
            # 为介质代码文件区域设置拖拽和点击
            self.code_drop_frame.drop_target_register(DND_FILES)
            self.code_drop_frame.dnd_bind('<<Drop>>', create_drop_handler(self.code_file, 'code', ['.xlsx', '.xls']))
            self.code_drop_frame.dnd_bind('<<DragEnter>>', on_drag_enter)
            self.code_drop_frame.dnd_bind('<<DragLeave>>', on_drag_leave)
            # 点击拖拽区域也能选择文件
            self.code_drop_frame.bind('<Button-1>', lambda e: self.select_code_file())
            
            print("拖拽功能已启用")
            
        except ImportError:
            print("拖拽功能需要安装 tkinterdnd2 库")
            print("使用命令: pip install tkinterdnd2")
    
    def show_drop_feedback(self, widget, state):
        """显示拖拽反馈"""
        try:
            if state == "hover":
                widget.configure(bg="#e3f2fd", relief="solid", borderwidth=3)
            elif state == "success":
                widget.configure(bg="#e8f5e8", relief="solid", borderwidth=3)
                # 1秒后恢复正常样式
                self.root.after(1000, lambda: widget.configure(bg="#f8f9fa", relief="solid", borderwidth=2))
            elif state == "error":
                widget.configure(bg="#ffebee", relief="solid", borderwidth=3)
                # 1秒后恢复正常样式
                self.root.after(1000, lambda: widget.configure(bg="#f8f9fa", relief="solid", borderwidth=2))
            else:  # normal
                widget.configure(bg="#f8f9fa", relief="solid", borderwidth=2)
        except Exception as e:
            print(f"拖拽反馈设置失败: {e}")
        
    def select_dwg_file(self):
        # 设置初始目录为最近使用的文件目录
        initialdir = None
        if self.recent_files['dwg']:
            initialdir = os.path.dirname(self.recent_files['dwg'][0])
        
        filename = filedialog.askopenfilename(
            title="选择DWG文件",
            filetypes=[("DWG files", "*.dwg"), ("All files", "*.*")],
            initialdir=initialdir
        )
        if filename:
            self.dwg_file.set(filename)
            self.add_recent_file('dwg', filename)
            
    def select_code_file(self):
        # 设置初始目录为最近使用的文件目录
        initialdir = None
        if self.recent_files['code']:
            initialdir = os.path.dirname(self.recent_files['code'][0])
        
        filename = filedialog.askopenfilename(
            title="选择介质代码文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=initialdir
        )
        if filename:
            self.code_file.set(filename)
            self.add_recent_file('code', filename)
            
    def select_output_file(self):
        # 设置初始目录为最近使用的文件目录
        initialdir = None
        if self.recent_files['output']:
            initialdir = os.path.dirname(self.recent_files['output'][0])
        
        # 生成默认文件名
        from datetime import datetime
        default_name = f"pipeline_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        filename = filedialog.asksaveasfilename(
            title="选择数据保存位置和文件名",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=initialdir
        )
        if filename:
            self.output_file.set(filename)
            self.add_recent_file('output', filename)
            
    def start_extraction(self):
        # 验证输入
        if not self.dwg_file.get():
            messagebox.showerror("错误", "请选择DWG文件")
            return
            
        if not self.code_file.get():
            messagebox.showerror("错误", "请选择介质代码文件")
            return
            
        if not self.output_file.get():
            messagebox.showerror("错误", "请选择输出文件")
            return
            
        # 在新线程中运行提取
        self.progress.start()
        self.status_label.config(text="正在提取数据...")
        self.result_text.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self.extract_data)
        thread.daemon = True
        thread.start()
        
    def log_message(self, message):
        """线程安全的日志记录"""
        self.root.after(0, lambda: self.result_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n"))
        self.root.after(0, lambda: self.result_text.see(tk.END))
        
    def extract_data(self):
        try:
            self.log_message("开始提取P&ID管道数据...")
            
            # 提取文本
            text_entities = self.extract_text_from_dwg(self.dwg_file.get())
            
            if not text_entities:
                self.log_message("未能提取到任何文本")
                self.extraction_complete(False)
                return
                
            self.log_message(f"提取了 {len(text_entities)} 个文本实体")
            
            # 查找管道号
            pipeline_numbers = self.find_pipeline_numbers(text_entities)
            self.log_message(f"找到 {len(pipeline_numbers)} 个管道号")
            
            # 加载介质代码
            medium_codes = self.load_medium_codes(self.code_file.get())
            self.log_message(f"加载了 {len(medium_codes)} 个介质代码")
            
            # 解析管道号
            pipeline_data = []
            for pipeline_number in pipeline_numbers:
                parsed_data = self.parse_pipeline_number(pipeline_number, medium_codes)
                if parsed_data:
                    pipeline_data.append(parsed_data)
                    
            self.log_message(f"成功解析 {len(pipeline_data)} 个管道号")
            
            # 创建Excel输出
            df = self.create_excel_output(pipeline_data, self.output_file.get())
            
            # 统计相态
            phase_counts = df['相态'].value_counts()
            self.log_message("相态统计:")
            for phase, count in phase_counts.items():
                self.log_message(f"  {phase}: {count}个")
            
            self.log_message(f"提取完成！结果已保存到: {self.output_file.get()}")
            self.extraction_complete(True)
            
        except Exception as e:
            self.log_message(f"提取过程中发生错误: {str(e)}")
            self.extraction_complete(False)
            
    def extraction_complete(self, success):
        """提取完成后的处理"""
        self.root.after(0, lambda: self.progress.stop())
        if success:
            self.root.after(0, lambda: self.status_label.config(text="提取完成！"))
            self.root.after(0, lambda: messagebox.showinfo("成功", "数据提取完成！"))
        else:
            self.root.after(0, lambda: self.status_label.config(text="提取失败"))
            self.root.after(0, lambda: messagebox.showerror("错误", "数据提取失败，请查看日志"))

    # Include all the extraction methods from the original file
    def extract_text_from_dwg(self, dwg_path):
        """从DWG文件中提取文本"""
        try:
            from pyautocad import Autocad
            
            # 连接到AutoCAD
            acad = Autocad(create_if_not_exists=True)
            self.log_message("成功连接到AutoCAD")
            
            # 打开文件
            abs_path = os.path.abspath(dwg_path)
            self.log_message(f"打开文件: {abs_path}")
            doc = acad.app.Documents.Open(abs_path)
            self.log_message(f"成功打开文件: {doc.Name}")
            
            # 获取模型空间
            model_space = doc.ModelSpace
            self.log_message(f"模型空间实体数量: {model_space.Count}")
            
            # 提取文本实体
            text_entities = []
            
            # 遍历实体
            total_entities = model_space.Count
            for i in range(total_entities):
                try:
                    # 显示进度
                    if i % 10000 == 0:
                        self.log_message(f"处理进度: {i}/{total_entities} ({i/total_entities*100:.1f}%)")
                    
                    entity = model_space.Item(i)
                    entity_type = entity.ObjectName
                    
                    # 只处理文本相关的实体类型，提高效率
                    if entity_type in ["AcDbText", "AcDbMText", "AcDbBlockReference"]:
                        # 提取文本
                        text_content = None
                        if entity_type == "AcDbText":
                            text_content = entity.TextString
                        elif entity_type == "AcDbMText":
                            text_content = entity.TextString
                        elif entity_type == "AcDbBlockReference":
                            # 处理块参照中的属性
                            try:
                                if hasattr(entity, 'GetAttributes'):
                                    attributes = entity.GetAttributes()
                                    for attr in attributes:
                                        if hasattr(attr, 'TextString'):
                                            text_entities.append(attr.TextString)
                            except:
                                pass
                        
                        if text_content:
                            text_entities.append(text_content)
                        
                except Exception:
                    continue
            
            # 关闭文档
            doc.Close(False)
            self.log_message("已关闭文档")
            
            return text_entities
            
        except Exception as e:
            self.log_message(f"提取文本失败: {e}")
            return []
            
    def normalize_text(self, s):
        """文本标准化，清理不可见字符"""
        import unicodedata
        s = str(s).strip()
        s = unicodedata.normalize('NFKC', s)  # Unicode标准化
        s = s.replace('\x00', '')  # 清理NULL字符
        s = re.sub(r'[\u2010-\u2015]', '-', s)  # Unicode连字符改为ASCII连字符
        s = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', s)  # 清理控制字符
        return s

    def find_pipeline_numbers(self, text_entities):
        """查找管道号 - 根据项目类型使用不同的格式"""
        project_type = self.project_type.get()
        
        if project_type == "乌兹项目":
            # 乌兹项目模式: PA-2001002A-100-C1C-N
            patterns = [
                # 完整格式: 介质代码-管道号-管径-管道等级-保温类型
                r'([A-Z0-9]{1,4})-([A-Z0-9]{4,8})-(\d{2,4})-([A-Z0-9]{1,4})-([A-Z0-9]{1,2})',
                # 简化格式: 介质代码-管道号-管径
                r'([A-Z0-9]{1,4})-([A-Z0-9]{4,8})-(\d{2,4})$',
            ]
            # 乌兹项目测试字符串
            test_strings = [
                'PA-2001002A-100-C1C-N',        # 标准格式
                'BW-2001003B-150-D2D-H',        # 其他介质
                'PA-2001004C-200-E3E-C',        # 大管径
                'PA-2001005D-50',               # 简化格式
            ]
        else:
            # 巨化项目模式 (原有模式)
            patterns = [
                # 标准完整格式: 4位装置号+1-4位介质代码-4-6位管道号-2-4位管径-标准等级格式-1-2位保温
                r'(\d{4}[A-Z0-9]{1,4})-([A-Z0-9]{4,6})-(\d{2,4})-(\d{2}[A-Z0-9]{3,6})-([A-Z]{1,2})',
                # 简化格式: 仅当缺少等级信息时
                r'(\d{4}[A-Z0-9]{1,4})-([A-Z0-9]{4,6})-(\d{2,4})$',
            ]
            # 巨化项目测试字符串
            test_strings = [
                '4101BRR-02457-200-03CBMB1-H',  # 原有测试
                '4101BRR-02457-1000-03CBMB1-H', # DN1000测试
                '4101CSM-01234-1200-02ABCD-H',  # DN1200测试
                '4101D-05678-50-01XYZ-C'        # 小管径测试
            ]
        
        for i, pattern in enumerate(patterns):
            self.log_message(f"测试模式 {i+1}: {pattern}")
            for test_string in test_strings:
                self_check = bool(re.search(pattern, test_string))
                self.log_message(f"  测试字符串 '{test_string}': {self_check}")
        
        pipeline_numbers = []
        pattern_stats = {i: 0 for i in range(len(patterns))}
        
        # 调试：打印前10个文本的详细信息
        self.log_message("开始分析前10个文本实体...")
        for idx, text in enumerate(text_entities[:10]):
            self.log_message(f"文本{idx}: {repr(text)} | 十六进制: {[hex(ord(c)) for c in str(text)[:20]]}")
        
        for text in text_entities:
            # 标准化文本
            normalized_text = self.normalize_text(text)
            
            # 尝试所有模式
            found_match = False
            for pattern_idx, pattern in enumerate(patterns):
                matches = re.findall(pattern, normalized_text)
                for match in matches:
                    if isinstance(match, tuple):
                        if len(match) == 5:  # 完整格式
                            pipeline_number = '-'.join(match)
                        elif len(match) == 3:  # 简化格式
                            pipeline_number = '-'.join(match)
                        else:
                            continue
                    else:
                        pipeline_number = match
                    
                    if pipeline_number not in pipeline_numbers:
                        pipeline_numbers.append(pipeline_number)
                        pattern_stats[pattern_idx] += 1
                        self.log_message(f"找到管道号: {pipeline_number} (模式{pattern_idx+1}, 原文本: {repr(text[:50])})")
                        found_match = True
                        break
                if found_match:
                    break
        
        # 输出统计信息
        self.log_message("各模式匹配统计:")
        for i, count in pattern_stats.items():
            self.log_message(f"  模式{i+1}: {count}个匹配")
        
        return pipeline_numbers
        
    def load_medium_codes(self, code_file_path):
        """从Excel文件加载介质代码映射"""
        try:
            df = pd.read_excel(code_file_path, header=None)
            medium_codes = {}
            
            for i, row in df.iterrows():
                code = row.iloc[0]
                name = row.iloc[1]
                
                # 处理代码列
                if pd.isna(code):
                    # 特殊处理氢氧化钠溶液
                    if not pd.isna(name) and "氢氧化钠溶液" in str(name):
                        code = "NA"
                    else:
                        continue
                else:
                    code = str(code).strip()
                
                # 处理名称列
                if pd.isna(name):
                    continue
                name = str(name).strip()
                
                if code and name and code != 'nan' and name != 'nan':
                    medium_codes[code] = name
                    
            return medium_codes
            
        except Exception as e:
            self.log_message(f"无法加载介质代码文件: {e}")
            return {}
            
    def determine_phase(self, medium_name):
        """根据介质名称判断相态"""
        # 气相关键词
        gas_keywords = ['蒸汽', '气','汽', '空气', '氢气', '氮气', '氧气', '二氧化碳', '天然气', '废气']
        
        # 液相关键词
        liquid_keywords = ['水', '油', '液', '溶液', '酸', '碱', '汽油', '柴油', '凝结']
        
        # 检查是否包含气相关键词
        for keyword in gas_keywords:
            if keyword in medium_name:
                return '气相'
        
        # 检查是否包含液相关键词
        for keyword in liquid_keywords:
            if keyword in medium_name:
                return '液相'
        
        # 默认返回未知相态
        return '未知相态'
        
    def parse_pipeline_number(self, pipeline_number, medium_codes):
        """解析管道号 - 根据项目类型支持不同格式"""
        project_type = self.project_type.get()
        parts = pipeline_number.split('-')
        
        if project_type == "乌兹项目":
            # 乌兹项目格式: PA-2001002A-100-C1C-N
            if len(parts) >= 5:
                # 完整格式: 介质代码-管道号-管径-管道等级-保温类型
                medium_code = parts[0]        # PA
                pipe_number = parts[1]        # 2001002A
                pipe_size = parts[2]          # 100
                pipe_grade = parts[3]         # C1C
                insulation_grade = parts[4]   # N
                
                # 乌兹项目的最终管道号是：介质代码-管道号
                simplified_pipeline_number = f"{medium_code}-{pipe_number}"
                
                # 对于乌兹项目，装置号为空或根据管道号推导
                unit_number = ""
                
            elif len(parts) >= 3:
                # 简化格式: 介质代码-管道号-管径
                medium_code = parts[0]        # PA
                pipe_number = parts[1]        # 2001002A
                pipe_size = parts[2]          # 100
                pipe_grade = "未知等级"
                insulation_grade = "未知"
                
                # 乌兹项目的最终管道号是：介质代码-管道号
                simplified_pipeline_number = f"{medium_code}-{pipe_number}"
                unit_number = ""
            else:
                return None
                
            medium_name = medium_codes.get(medium_code, f"未知介质({medium_code})")
            phase = self.determine_phase(medium_name)
            
            return {
                'pipeline_number': simplified_pipeline_number,
                'unit_number': unit_number,
                'pipe_number': pipe_number,
                'nominal_diameter': pipe_size,
                'pipe_grade': pipe_grade,
                'insulation_grade': insulation_grade,
                'medium_code': medium_code,
                'medium_name': medium_name,
                'phase': phase
            }
        
        else:
            # 巨化项目格式 (原有逻辑)
            if len(parts) >= 5:
                # 完整格式: 装置号和介质代码-管道号-管道尺寸-管道等级-保温等级
                unit_and_medium = parts[0]  # 4101BRR
                pipe_number = parts[1]      # 02457
                pipe_size = parts[2]        # 200 或 1000
                pipe_grade = parts[3]       # 03CBMB1
                insulation_grade = parts[4] # H
                
                # 智能提取装置号和介质代码
                # 假设装置号为前3-5位数字，介质代码为剩余部分
                match = re.match(r'(\d{3,5})([A-Z0-9]+)', unit_and_medium)
                if match:
                    unit_number = match.group(1)
                    medium_code = match.group(2)
                else:
                    # 退化处理：假设前4位为装置号
                    unit_number = unit_and_medium[:4] if len(unit_and_medium) >= 4 else unit_and_medium
                    medium_code = unit_and_medium[4:] if len(unit_and_medium) > 4 else ""
                
                medium_name = medium_codes.get(medium_code, f"未知介质({medium_code})")
                phase = self.determine_phase(medium_name)
                
                # 简化的管道号：装置号和介质代码-管道编号
                simplified_pipeline_number = f"{unit_number}{medium_code}-{pipe_number}"
                
                return {
                    'pipeline_number': simplified_pipeline_number,
                    'unit_number': unit_number,
                    'pipe_number': pipe_number,
                    'nominal_diameter': pipe_size,
                    'pipe_grade': pipe_grade,
                    'insulation_grade': insulation_grade,
                    'medium_code': medium_code,
                    'medium_name': medium_name,
                    'phase': phase
                }
            elif len(parts) >= 3:
                # 简化格式: 装置号和介质代码-管道号-管道尺寸
                unit_and_medium = parts[0]
                pipe_number = parts[1]
                pipe_size = parts[2]
                pipe_grade = "未知等级"
                insulation_grade = "未知"
                
                # 智能提取装置号和介质代码
                match = re.match(r'(\d{3,5})([A-Z0-9]+)', unit_and_medium)
                if match:
                    unit_number = match.group(1)
                    medium_code = match.group(2)
                else:
                    unit_number = unit_and_medium[:4] if len(unit_and_medium) >= 4 else unit_and_medium
                    medium_code = unit_and_medium[4:] if len(unit_and_medium) > 4 else ""
                
                medium_name = medium_codes.get(medium_code, f"未知介质({medium_code})")
                phase = self.determine_phase(medium_name)
                
                # 简化的管道号：装置号和介质代码-管道编号
                simplified_pipeline_number = f"{unit_number}{medium_code}-{pipe_number}"
                
                return {
                    'pipeline_number': simplified_pipeline_number,
                    'unit_number': unit_number,
                    'pipe_number': pipe_number,
                    'nominal_diameter': pipe_size,
                    'pipe_grade': pipe_grade,
                    'insulation_grade': insulation_grade,
                    'medium_code': medium_code,
                    'medium_name': medium_name,
                    'phase': phase
                }
        
        return None
        
    def create_excel_output(self, pipeline_data, output_path):
        """创建Excel输出"""
        # 创建DataFrame
        df_data = []
        for data in pipeline_data:
            if data:
                df_data.append([
                    data['pipeline_number'],
                    data['nominal_diameter'],
                    data['pipe_grade'],
                    data['insulation_grade'],
                    data['medium_name'],
                    data['phase']
                ])
        
        columns = ['管道号', '管径', '管道等级', '保温等级', '介质名称', '相态']
        df = pd.DataFrame(df_data, columns=columns)
        
        # 按管道号排序
        df = df.sort_values('管道号').reset_index(drop=True)
        
        # 保存为Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='管道数据表', index=False)
            
            # 设置列宽
            worksheet = writer.sheets['管道数据表']
            column_widths = {'A': 20, 'B': 8, 'C': 15, 'D': 10, 'E': 15, 'F': 8}
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
            
            # 设置表头样式
            from openpyxl.styles import Font, PatternFill, Alignment
            header_font = Font(bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
            header_alignment = Alignment(horizontal='center', vertical='center')
            
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
        
        return df

def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
        print("tkinterdnd2 不可用，拖拽功能将被禁用")
    
    app = PIDExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()