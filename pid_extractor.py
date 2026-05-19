#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P&ID管道数据提取工具 - CLI版本
从P&ID图纸中提取管道号并生成Excel报告
"""

import logging
import argparse
from extractor_core import (
    SUPPORTED_PROJECT_TYPES,
    get_resource_path,
    extract_text_from_dwg,
    find_pipeline_numbers,
    load_medium_codes,
    parse_pipeline_number,
    create_excel_output,
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser(description='P&ID管道数据提取工具')
    parser.add_argument(
        '--project-type',
        choices=SUPPORTED_PROJECT_TYPES,
        default='巨化项目',
        help=f'项目类型选择 (可选: {", ".join(SUPPORTED_PROJECT_TYPES)}, 默认: 巨化项目)',
    )
    parser.add_argument('--dwg-file',    default='test/test.dwg',        help='DWG文件路径')
    parser.add_argument('--code-file',   default='test/code.xlsx',        help='介质代码文件路径')
    parser.add_argument('--output-file', default='pipeline_data.xlsx',    help='输出Excel文件路径')
    args = parser.parse_args()

    logger.info(f"开始提取P&ID管道数据... (项目类型: {args.project_type})")

    text_entities = extract_text_from_dwg(get_resource_path(args.dwg_file))
    if not text_entities:
        logger.error("未能提取到任何文本")
        return

    pipeline_numbers = find_pipeline_numbers(text_entities, args.project_type)
    logger.info(f"找到 {len(pipeline_numbers)} 个管道号")

    medium_codes = load_medium_codes(get_resource_path(args.code_file))

    pipeline_data = [
        parse_pipeline_number(pn, medium_codes, args.project_type)
        for pn in pipeline_numbers
    ]
    pipeline_data = [d for d in pipeline_data if d]
    logger.info(f"成功解析 {len(pipeline_data)} 个管道号")

    df = create_excel_output(pipeline_data, args.output_file)
    print(f"\n处理完成！项目类型: {args.project_type}")
    print(f"提取到 {len(df)} 个管道号，结果已保存到: {args.output_file}")


if __name__ == "__main__":
    main()
