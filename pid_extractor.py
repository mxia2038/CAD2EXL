#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P&ID管道数据提取工具
从P&ID图纸中提取管道号并生成Excel报告
"""

import re
import pandas as pd
import logging
import os
import sys

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def extract_text_from_dwg(dwg_path):
    """从DWG文件中提取文本"""
    try:
        from pyautocad import Autocad
        
        # 连接到AutoCAD
        acad = Autocad(create_if_not_exists=True)
        logger.info("成功连接到AutoCAD")
        
        # 打开文件
        abs_path = os.path.abspath(dwg_path)
        logger.info(f"打开文件: {abs_path}")
        doc = acad.app.Documents.Open(abs_path)
        logger.info(f"成功打开文件: {doc.Name}")
        
        # 获取模型空间
        model_space = doc.ModelSpace
        logger.info(f"模型空间实体数量: {model_space.Count}")
        
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
        
        logger.info(f"提取了 {len(text_entities)} 个文本")
        
        # 关闭文档
        doc.Close(False)
        logger.info("已关闭文档")
        
        return text_entities
        
    except Exception as e:
        logger.error(f"提取文本失败: {e}")
        return []

def find_pipeline_numbers(text_entities):
    """查找管道号"""
    pipeline_pattern = r'(\d{4})-(\d{1,3})-([A-Z]{1,3})-(\d{6}[A-Z]?)-([A-Z0-9]{1,4}[A-Z]?)-([A-Z]{1,3})'
    pipeline_numbers = []
    
    for text in text_entities:
        matches = re.findall(pipeline_pattern, text)
        for match in matches:
            pipeline_number = '-'.join(match)
            if pipeline_number not in pipeline_numbers:
                pipeline_numbers.append(pipeline_number)
    
    return pipeline_numbers

def load_medium_codes(code_file_path):
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
                
        logger.info(f"成功加载 {len(medium_codes)} 个介质代码")
        return medium_codes
        
    except Exception as e:
        logger.error(f"无法加载介质代码文件: {e}")
        return {}

def determine_phase(medium_name):
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

def parse_pipeline_number(pipeline_number, medium_codes):
    """解析管道号"""
    parts = pipeline_number.split('-')
    if len(parts) >= 6:
        medium_code = parts[2]
        medium_name = medium_codes.get(medium_code, f"未知介质({medium_code})")
        phase = determine_phase(medium_name)
        
        return {
            'pipeline_number': pipeline_number,
            'nominal_diameter': parts[1],
            'pipe_grade': parts[4],
            'insulation_type': parts[5],
            'medium_name': medium_name,
            'phase': phase
        }
    return None

def create_excel_output(pipeline_data, output_path):
    """创建Excel输出"""
    # 创建DataFrame
    df_data = []
    for data in pipeline_data:
        if data:
            df_data.append([
                data['pipeline_number'],
                data['nominal_diameter'],
                data['pipe_grade'],
                data['insulation_type'],
                data['medium_name'],
                data['phase']
            ])
    
    columns = ['管道号', '管径', '管道等级', '保温型式', '介质', '相态']
    df = pd.DataFrame(df_data, columns=columns)
    
    # 按管道号排序
    df = df.sort_values('管道号').reset_index(drop=True)
    
    # 保存为Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='管道数据表', index=False)
        
        # 设置列宽
        worksheet = writer.sheets['管道数据表']
        column_widths = {'A': 25, 'B': 10, 'C': 15, 'D': 15, 'E': 15, 'F': 10}
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
    
    logger.info(f"成功保存Excel文件: {output_path}")
    return df

def get_resource_path(relative_path):
    """获取资源文件路径（支持打包后的exe）"""
    try:
        # PyInstaller临时文件夹
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def main():
    """主函数"""
    logger.info("开始提取P&ID管道数据...")
    
    # 配置文件路径
    dwg_file = get_resource_path("test/test.dwg")
    code_file = get_resource_path("test/code.xlsx")
    output_file = "pipeline_data.xlsx"
    
    # 提取文本
    text_entities = extract_text_from_dwg(dwg_file)
    
    if not text_entities:
        logger.error("未能提取到任何文本")
        return
    
    # 查找管道号
    pipeline_numbers = find_pipeline_numbers(text_entities)
    logger.info(f"找到 {len(pipeline_numbers)} 个管道号")
    
    # 加载介质代码
    medium_codes = load_medium_codes(code_file)
    
    # 解析管道号
    pipeline_data = []
    for pipeline_number in pipeline_numbers:
        parsed_data = parse_pipeline_number(pipeline_number, medium_codes)
        if parsed_data:
            pipeline_data.append(parsed_data)
    
    logger.info(f"成功解析 {len(pipeline_data)} 个管道号")
    
    # 创建Excel输出
    df = create_excel_output(pipeline_data, output_file)
    
    print(f"\n处理完成！")
    print(f"提取到 {len(df)} 个管道号")
    print(f"结果已保存到: {output_file}")

if __name__ == "__main__":
    main()