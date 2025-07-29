#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P&ID管道数据提取工具 - GUI版本
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
        self.root.geometry("600x500")
        
        # 文件路径变量
        self.dwg_file = tk.StringVar()
        self.code_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        # 设置默认值
        self.code_file.set("test/code.xlsx")
        self.output_file.set("pipeline_data.xlsx")
        
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame = main_frame  # 保存引用以便logo使用
        
        # 添加公司Logo
        self.setup_logo()
        
        # 标题
        title_label = ttk.Label(main_frame, text="P&ID管道数据提取工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # DWG文件选择
        ttk.Label(main_frame, text="DWG文件:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.dwg_file, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="浏览", command=self.select_dwg_file).grid(row=2, column=2, pady=5)
        
        # 介质代码文件选择
        ttk.Label(main_frame, text="介质代码文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.code_file, width=50).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="浏览", command=self.select_code_file).grid(row=3, column=2, pady=5)
        
        # 输出文件选择
        ttk.Label(main_frame, text="输出文件:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_file, width=50).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="浏览", command=self.select_output_file).grid(row=4, column=2, pady=5)
        
        # 提取按钮
        extract_button = ttk.Button(main_frame, text="开始提取", command=self.start_extraction)
        extract_button.grid(row=5, column=0, columnspan=3, pady=20)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="请选择DWG文件开始提取")
        self.status_label.grid(row=7, column=0, columnspan=3, pady=10)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="提取结果", padding="10")
        result_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 结果文本框
        self.result_text = tk.Text(result_frame, height=10, width=70)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(8, weight=1)
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
                target_height = 60
                aspect_ratio = original_width / original_height
                target_width = int(target_height * aspect_ratio)
                logo_image = logo_image.resize((target_width, target_height), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                
                # 显示logo (在标题前)
                logo_label = tk.Label(self.main_frame, image=self.logo_photo)
                logo_label.grid(row=0, column=0, columnspan=3, pady=10)
                
        except Exception as e:
            print(f"无法加载logo: {e}")
        
    def select_dwg_file(self):
        filename = filedialog.askopenfilename(
            title="选择DWG文件",
            filetypes=[("DWG files", "*.dwg"), ("All files", "*.*")]
        )
        if filename:
            self.dwg_file.set(filename)
            
    def select_code_file(self):
        filename = filedialog.askopenfilename(
            title="选择介质代码文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.code_file.set(filename)
            
    def select_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            
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
            for i in range(model_space.Count):
                try:
                    entity = model_space.Item(i)
                    entity_type = entity.ObjectName
                    
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
                    elif hasattr(entity, 'TextString'):
                        text_content = entity.TextString
                    
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
            
    def find_pipeline_numbers(self, text_entities):
        """查找管道号"""
        # 新的管道号格式: 装置号和介质代码-管道号-管道尺寸-管道等级-保温等级
        # 示例: 4101BRR-02457-200-03CBMB1-H
        pipeline_pattern = r'(\d{4}[A-Z]{2,3})-(\d{5})-(\d{2,3})-(\d{2}[A-Z0-9]{3,6})-([A-Z]{1,2})'
        pipeline_numbers = []
        
        for text in text_entities:
            matches = re.findall(pipeline_pattern, text)
            for match in matches:
                pipeline_number = '-'.join(match)
                if pipeline_number not in pipeline_numbers:
                    pipeline_numbers.append(pipeline_number)
        
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
        gas_keywords = ['蒸汽', '气', '空气', '氢气', '氮气', '氧气', '二氧化碳', '天然气', '废气']
        
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
        """解析管道号"""
        parts = pipeline_number.split('-')
        if len(parts) >= 5:
            # 新格式: 装置号和介质代码-管道号-管道尺寸-管道等级-保温等级
            unit_and_medium = parts[0]  # 4101BRR
            pipe_number = parts[1]      # 02457
            pipe_size = parts[2]        # 200
            pipe_grade = parts[3]       # 03CBMB1
            insulation_grade = parts[4] # H
            
            # 从装置号和介质代码中提取介质代码（后2-3位字母）
            unit_number = unit_and_medium[:4]  # 4101
            medium_code = unit_and_medium[4:]  # BRR
            
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
    root = tk.Tk()
    app = PIDExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()