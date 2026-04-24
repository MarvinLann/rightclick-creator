#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel格式整理工具
功能：将Excel文件整理为标准格式
- 标题：微软雅黑 20px 加粗 居中
- 副标题：微软雅黑 14px
- 表头：灰色背景 + 微软雅黑 11px 加粗
- 数据：微软雅黑 11px，删除emoji，句号间有空行
- 页面：横向A4
"""

import re
import os
import emoji
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins


def remove_emoji(text):
    """删除emoji并清理空格"""
    if not isinstance(text, str):
        return text
    text = emoji.replace_emoji(text, '')
    return text.strip()


def split_sentences_with_space(text):
    """在句号之间添加空行"""
    if not isinstance(text, str):
        return text
    
    sentences = re.split(r'([。！？])', text)
    result = []
    
    for i in range(0, len(sentences), 2):
        part = sentences[i].strip()
        if i + 1 < len(sentences):
            part += sentences[i + 1]
        if part:
            result.append(part)
    
    return '\n\n'.join(result) if result else text


def is_header_row(ws, row):
    """检查是否是表头行"""
    header_keywords = ['序号', '名称', '项目', '内容', '描述', '证明', '页码', '材料', '来源', '形式', '时间', '目的', '等级', '位置', '问题', '风险', '建议', '修改', '方向']
    keyword_count = 0
    for col in range(1, min(ws.max_column + 1, 10)):
        cell_value = str(ws.cell(row, col).value) if ws.cell(row, col).value else ''
        for keyword in header_keywords:
            if keyword in cell_value and len(cell_value) < 20:
                keyword_count += 1
                break
    return keyword_count >= 2


def identify_row_type(ws, row):
    """智能识别行类型"""
    first_cell = ws.cell(row, 1)
    cell_value = str(first_cell.value) if first_cell.value else ''
    
    # 优先检查是否为表头
    if is_header_row(ws, row):
        return 'header'
    
    # 然后检查是否为标题（合并单元格跨多列）
    if row == 1:
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == 1 and merged_range.min_col == 1:
                if merged_range.max_col - merged_range.min_col >= 2:  # 跨2列以上
                    return 'title'
        # 如果第一行不是明显的标题，且包含表头关键词，则视为表头
        header_keywords = ['序号', '名称', '项目', '内容', '描述', '证明', '页码', '材料', '来源', '形式', '时间', '目的', '等级', '位置', '问题', '风险', '建议', '修改']
        first_row_text = ' '.join(str(ws.cell(row, col).value) for col in range(1, min(ws.max_column + 1, 10)))
        if any(keyword in first_row_text for keyword in header_keywords):
            return 'header'
    
    if row <= 5:
        subtitle_keywords = ['致：', '案号：', '提交人：', '提交日期', '生成日期']
        if any(keyword in cell_value for keyword in subtitle_keywords):
            return 'subtitle'
    
    return 'data'


def apply_format(ws, row, row_type):
    """应用格式"""
    if row_type == 'title':
        font = Font(name='微软雅黑', size=20, bold=True)
        alignment = Alignment(horizontal='center', vertical='center')
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            cell.font = font
            cell.alignment = alignment
    
    elif row_type == 'subtitle':
        font = Font(name='微软雅黑', size=14)
        alignment = Alignment(horizontal='left', vertical='center')
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            cell.font = font
            cell.alignment = alignment
    
    elif row_type == 'header':
        fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        font = Font(name='微软雅黑', size=10, bold=True)
        alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            cell.fill = fill
            cell.font = font
            cell.alignment = alignment
    
    elif row_type == 'data':
        font = Font(name='微软雅黑', size=10)
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            if cell.value and isinstance(cell.value, str):
                cell.value = remove_emoji(cell.value)
                if col >= 3 and '。' in cell.value:
                    cell.value = split_sentences_with_space(cell.value)
                
                # 根据内容长度决定对齐方式
                text_length = len(cell.value.strip())
                if text_length < 20:  # 内容少的单元格，上下左右居中
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                else:  # 内容多的单元格，靠左并保持分行
                    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            else:
                # 非文本或空值单元格，保持居中
                cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = font


def format_excel(input_file, output_file):
    """格式化Excel文件"""
    wb = None
    tmp_path = None
    try:
        # 检查文件格式，如果是XLS，先转换为XLSX
        if input_file.lower().endswith('.xls'):
            import tempfile
            df_dict = pd.read_excel(input_file, sheet_name=None)
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp_path = tmp.name
                with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
                    for sheet_name, df in df_dict.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            wb = load_workbook(tmp_path)
        else:
            wb = load_workbook(input_file)
        
        ws = wb.active
        
        # 处理每一行
        for row in range(1, ws.max_row + 1):
            row_type = identify_row_type(ws, row)
            apply_format(ws, row, row_type)
        
        # 设置合理的列宽
        if ws.max_column >= 5:
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 40
            ws.column_dimensions['D'].width = 50
            ws.column_dimensions['E'].width = 10
        elif ws.max_column == 4:
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 18
            ws.column_dimensions['C'].width = 50
            ws.column_dimensions['D'].width = 50
        else:
            for col in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].width = 20
        
        # 设置页面为横向A4
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_margins = PageMargins(left=1.5, right=1.5, top=1, bottom=1)
        ws.print_options.horizontalCentered = True
        
        # 自动调整行高
        for row in range(1, ws.max_row + 1):
            ws.row_dimensions[row].height = None
        
        wb.save(output_file)
        print(f"已保存到: {output_file}")
    except Exception as e:
        print(f"错误: 处理Excel文件失败: {e}")
        raise
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


if __name__ == "__main__":
    import sys
    import os
    
    if len(sys.argv) < 2:
        print("用法: python excel_format.py <excel文件路径>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    if not os.path.exists(input_file):
        print(f"错误: 文件不存在: {input_file}")
        sys.exit(1)
    
    if not input_file.lower().endswith(('.xlsx', '.xls')):
        print("错误: 只支持Excel文件")
        sys.exit(1)
    
    # 获取输入文件所在目录和文件名
    input_dir = os.path.dirname(input_file)
    input_filename = os.path.basename(input_file)
    output_filename = os.path.splitext(input_filename)[0] + '_格式化.xlsx'
    
    # 在同一目录下生成输出文件
    if input_dir:
        output_file = os.path.join(input_dir, output_filename)
    else:
        output_file = output_filename
    
    format_excel(input_file, output_file)
