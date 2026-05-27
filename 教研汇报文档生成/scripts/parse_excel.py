"""
解析 Excel 看板数据 - 教研汇报技能
用法：python parse_excel.py
"""

import pandas as pd
import sys
import json
import os

def parse_watchboard_excel(file_path, label):
    """
    解析单个看板 Excel 文件，提取四率数据
    """
    print(f"\n{'='*60}")
    print(f"【{label}】文件: {os.path.basename(file_path)}")
    
    try:
        xl = pd.ExcelFile(file_path)
        print(f"工作表列表: {xl.sheet_names}")
        
        # 优先找"实际达成数据"工作表
        target_sheet = None
        for sheet in xl.sheet_names:
            if "实际达成" in sheet or "达成数据" in sheet:
                target_sheet = sheet
                break
        if target_sheet is None:
            target_sheet = xl.sheet_names[0]
        
        print(f"使用工作表: {target_sheet}")
        df = xl.parse(target_sheet, header=None)
        
        print(f"尺寸: {df.shape}")
        print("前60行数据:")
        print(df.to_string(max_rows=60, max_cols=30))
        
        return df
    except Exception as e:
        print(f"读取失败: {e}")
        return None


def extract_indicators(df):
    """
    从 DataFrame 中提取四率指标
    返回: [
        {"indicator": "课中首答正确率", "s6_s7": [...], "target": "..."},
        ...
    ]
    """
    indicators = []
    # 这里根据实际 Excel 结构调整
    # 通常首行是标题，之后每行一个指标
    return indicators


if __name__ == "__main__":
    # 示例文件路径 - 实际使用时由用户上传
    files = {
        "S1-S2": r"C:/Users/yelin01/Downloads/海外教学北极星指标达成监控-看板 (2).xlsx",
        "S3-S5": r"C:/Users/yelin01/Downloads/海外教学北极星指标达成监控-看板 (1).xlsx",
        "S6-S7": r"C:/Users/yelin01/Downloads/海外教学北极星指标达成监控-看板.xlsx",
    }
    
    results = {}
    for label, path in files.items():
        if os.path.exists(path):
            df = parse_watchboard_excel(path, label)
            if df is not None:
                results[label] = df
        else:
            print(f"文件不存在: {path}")
    
    print(f"\n解析完成，共 {len(results)} 个文件")
