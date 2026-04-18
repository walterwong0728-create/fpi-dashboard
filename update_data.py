#!/usr/bin/env python3
"""
FPI数据更新脚本 - 支持数据累积
功能：读取新Excel数据，与现有数据合并（已有日期更新，新日期追加）
用法：python update_data.py <Excel文件路径>
"""

import sys
import json
import re
import openpyxl
from datetime import datetime

def parse_excel(excel_path):
    """解析Excel文件，返回数据列表"""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):  # 跳过表头
        if row[0] is None:
            continue
        
        # 解析日期
        date_val = row[0]
        if isinstance(date_val, datetime):
            date_str = date_val.strftime('%Y-%m-%d')
        elif isinstance(date_val, str):
            # 尝试解析字符串日期
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                try:
                    date_str = datetime.strptime(date_val, fmt).strftime('%Y-%m-%d')
                    break
                except:
                    continue
            else:
                continue
        else:
            continue
        
        record = {
            'dateStr': date_str,
            'visitors': int(row[1]) if row[1] else 0,
            'stay': float(row[2]) if row[2] else 0,
            'cost': float(row[3]) if row[3] else 0,
            'roi': float(row[4]) if row[4] else 0,
            'cart': int(row[5]) if row[5] else 0,
            'buyers': int(row[6]) if row[6] else 0,
            'sales': float(row[7]) if row[7] else 0,
            'kw_people': int(row[8]) if row[8] else 0,
            'kw_cart': int(row[9]) if row[9] else 0,
            'kw_pay': int(row[10]) if row[10] else 0,
            'crowd_people': int(row[11]) if row[11] else 0,
            'crowd_cart': int(row[12]) if row[12] else 0,
            'crowd_pay': int(row[13]) if row[13] else 0,
            'search_people': int(row[14]) if row[14] else 0,
            'search_cart': int(row[15]) if row[15] else 0,
            'search_pay': int(row[16]) if row[16] else 0,
            'rec_people': int(row[17]) if row[17] else 0,
            'rec_cart': int(row[18]) if row[18] else 0,
            'rec_pay': int(row[19]) if row[19] else 0,
        }
        data.append(record)
    
    return data

def merge_data(existing_data, new_data):
    """
    合并数据：已有日期更新，新日期追加
    """
    existing_dict = {r['dateStr']: r for r in existing_data}
    
    added_count = 0
    updated_count = 0
    
    for record in new_data:
        date_str = record['dateStr']
        if date_str in existing_dict:
            existing_dict[date_str] = record  # 更新
            updated_count += 1
        else:
            existing_dict[date_str] = record  # 追加
            added_count += 1
    
    # 按日期排序
    merged = sorted(existing_dict.values(), key=lambda x: x['dateStr'])
    
    return merged, added_count, updated_count

def update_html_file(html_path, merged_data):
    """更新HTML文件中的数据"""
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 生成新的数据数组
    first_date = merged_data[0]['dateStr']
    last_date = merged_data[-1]['dateStr']
    total_days = len(merged_data)
    
    new_data_str = f"// ====== 原始数据 ({first_date} ~ {last_date}, 共{total_days}天) ======\n"
    new_data_str += "// ⚠️ 数据累积：每次更新时追加新数据，已有的日期会更新，不重复不丢失\n"
    new_data_str += "const rawData = [\n"
    
    for r in merged_data:
        new_data_str += f"  {{ dateStr: '{r['dateStr']}', visitors: {r['visitors']}, stay: {r['stay']}, cost: {r['cost']}, roi: {r['roi']}, cart: {r['cart']}, buyers: {r['buyers']}, sales: {r['sales']}, kw_people: {r['kw_people']}, kw_cart: {r['kw_cart']}, kw_pay: {r['kw_pay']}, crowd_people: {r['crowd_people']}, crowd_cart: {r['crowd_cart']}, crowd_pay: {r['crowd_pay']}, search_people: {r['search_people']}, search_cart: {r['search_cart']}, search_pay: {r['search_pay']}, rec_people: {r['rec_people']}, rec_cart: {r['rec_cart']}, rec_pay: {r['rec_pay']} }},\n"
    
    new_data_str += "];\n"
    new_data_str += f"// ⚠️ 更新日志：[{first_date}~{last_date}] 共{total_days}天 | 新增{len(new_data) if 'new_data' in dir() else 0}条，更新{existing_count if 'existing_count' in dir() else 0}条\n"
    
    # 替换旧数据
    pattern = r'// ====== 原始数据 \(.*?\) ======\s*\n(?:// ⚠️ 数据累积.*?\n)?const rawData = \[.*?\];\s*\n(?:// ⚠️ 更新日志.*?\n)?'
    content = re.sub(pattern, new_data_str, content, flags=re.DOTALL)
    
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    return content

def main():
    if len(sys.argv) < 2:
        print("用法: python update_data.py <Excel文件路径>")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    html_path = '/workspace/fpi-dashboard/index.html'
    
    # 1. 解析新Excel
    print(f"📊 读取Excel: {excel_path}")
    new_data = parse_excel(excel_path)
    print(f"   Excel中共有 {len(new_data)} 条记录")
    
    # 2. 读取现有数据
    print(f"📂 读取现有HTML数据...")
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 提取现有rawData
    raw_match = re.search(r'const rawData = \[(.*?)\];', content, re.DOTALL)
    if raw_match:
        # 解析现有数据（简化处理，直接用正则提取）
        existing_data = []
        record_pattern = r"dateStr: '([^']+)'"
        records = re.findall(record_pattern, content)
        for date_str in records:
            # 找到完整记录
            rec_pattern = rf"{{ dateStr: '{date_str}',[^}}]+}}"
            rec_match = re.search(rec_pattern, content)
            if rec_match:
                # 简单解析各字段
                rec_str = rec_match.group(0)
                existing_data.append({'dateStr': date_str})
    else:
        existing_data = []
    
    print(f"   现有数据: {len(existing_data)} 条")
    
    # 3. 合并数据
    merged_data, added, updated = merge_data(existing_data, new_data)
    print(f"🔄 合并完成: 新增{added}条, 更新{updated}条, 共{len(merged_data)}条")
    
    # 4. 更新HTML
    update_html_file(html_path, merged_data)
    print(f"✅ 已更新: {html_path}")

if __name__ == '__main__':
    main()
