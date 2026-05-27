"""
解析 S1-S2 学段 Excel 看板中的愉悦度和镜头感数据
用法：python parse_s1s2.py <excel文件路径>
输出：愉悦度和镜头感各月数据（1-5月），以及定位到的行号
"""

import pandas as pd
import sys
import os

def find_row_by_keyword(df, keyword, max_search=200):
    """在 DataFrame 中按关键词找行号"""
    for i, row in df.iterrows():
        for v in row.values:
            if isinstance(v, str) and keyword in v:
                return i
    return None

def extract_metric_data(df, start_row, num_rows=6):
    """
    从起始行开始，提取各班型数据
    返回：列表，每个元素是 (班型名称, [1月值, 2月值, ...])
    """
    results = []
    for offset in range(num_rows):
        row_idx = start_row + offset
        if row_idx >= len(df):
            break
        row = df.iloc[row_idx]
        # 找班型名称（字符串且在合理位置）
        row_vals = list(row.values)
        # 第一个非空字符串通常是班型名
        class_name = None
        data_vals = []
        for v in row_vals:
            if pd.isna(v):
                continue
            s = str(v).strip()
            # 判断是否是班型名称
            if class_name is None and ('班' in s or '普通' in s or '粤语' in s or '英语' in s):
                class_name = s
            elif class_name is not None:
                # 已经是数据部分
                try:
                    fv = float(v)
                    data_vals.append(fv)
                except:
                    if s and class_name is not None:
                        # 可能还是班型名？跳过
                        pass
        if class_name and data_vals:
            results.append((class_name, data_vals))
    return results

def main():
    if len(sys.argv) < 2:
        print("用法: python parse_s1s2.py <Excel文件路径>")
        sys.exit(1)

    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        sys.exit(1)

    print(f"读取文件: {file_path}")
    print("="*70)

    try:
        xl = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"无法读取Excel: {e}")
        sys.exit(1)

    sheet_name = None
    for name in xl.sheet_names:
        if '实际达成' in name:
            sheet_name = name
            break
    if sheet_name is None:
        sheet_name = xl.sheet_names[0]
    
    print(f"使用工作表: {sheet_name}")
    df = xl.parse(sheet_name, header=None)

    # 定位愉悦度和镜头感
    joy_row = find_row_by_keyword(df, '愉悦度')
    lens_row = find_row_by_keyword(df, '镜头感')

    print(f"\n愉悦度 起始行: {joy_row}")
    print(f"镜头感 起始行: {lens_row}")

    # 打印定位行附近的原始数据（用于调试）
    if joy_row is not None:
        print("\n--- 愉悦度附近数据 ---")
        for i in range(max(0, joy_row-2), min(len(df), joy_row+10)):
            row = df.iloc[i]
            vals = [str(v) for v in row.values[:8] if pd.notna(v)]
            if vals:
                print(f"  行{i}: {vals}")

    if lens_row is not None:
        print("\n--- 镜头感附近数据 ---")
        for i in range(max(0, lens_row-2), min(len(df), lens_row+10)):
            row = df.iloc[i]
            vals = [str(v) for v in row.values[:8] if pd.notna(v)]
            if vals:
                print(f"  行{i}: {vals}")

    # 提取数据（使用更精确的方式）
    print("\n" + "="*70)
    print("提取愉悦度数据（精确模式）:")
    print("="*70)

    if joy_row is not None:
        # 找到标题行后，下一行开始是班型数据
        # 班型行特征：包含"普通班"、"粤语班"、"英语班"
        joy_data = []
        for offset in range(1, 8):  # 最多读8行
            row_idx = joy_row + offset
            if row_idx >= len(df):
                break
            row = df.iloc[row_idx]
            # 找第一个字符串（班型名）
            class_name = None
            data_start_col = None
            for col_idx, v in enumerate(row.values):
                if pd.isna(v):
                    continue
                s = str(v).strip()
                if class_name is None and ('班' in s or '普通' in s or '粤语' in s or '英语' in s):
                    class_name = s
                    data_start_col = col_idx + 1
                elif class_name is not None and data_start_col is not None and col_idx >= data_start_col:
                    break
            if class_name:
                # 提取这一行从班型名后面开始的数字
                data_vals = []
                for v in row.values:
                    if pd.isna(v):
                        continue
                    try:
                        fv = float(v)
                        data_vals.append(fv)
                    except:
                        pass
                if data_vals:
                    joy_data.append((class_name, data_vals))
        
        for name, vals in joy_data:
            # 只取前5个值（1-5月）
            display_vals = vals[:5]
            print(f"  {name}: {display_vals}")

    print("\n" + "="*70)
    print("提取镜头感数据（精确模式）:")
    print("="*70)

    if lens_row is not None:
        lens_data = []
        for offset in range(1, 8):
            row_idx = lens_row + offset
            if row_idx >= len(df):
                break
            row = df.iloc[row_idx]
            class_name = None
            for col_idx, v in enumerate(row.values):
                if pd.isna(v):
                    continue
                s = str(v).strip()
                if class_name is None and ('班' in s or '普通' in s or '粤语' in s or '英语' in s):
                    class_name = s
                    break
            if class_name:
                data_vals = []
                for v in row.values:
                    if pd.isna(v):
                        continue
                    try:
                        fv = float(v)
                        data_vals.append(fv)
                    except:
                        pass
                if data_vals:
                    lens_data.append((class_name, data_vals))
        
        for name, vals in lens_data:
            display_vals = vals[:5]
            print(f"  {name}: {display_vals}")

    # 保存中间结果为 JSON（供生成脚本使用）
    import json
    output = {
        "joy": [{"class": n, "values": v[:5]} for n, v in joy_data] if joy_row is not None else [],
        "lens": [{"class": n, "values": v[:5]} for n, v in lens_data] if lens_row is not None else [],
        "joy_row": int(joy_row) if joy_row else None,
        "lens_row": int(lens_row) if lens_row else None,
    }
    
    out_path = file_path + ".parsed.json"
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n解析结果已保存: {out_path}")

if __name__ == '__main__':
    main()
