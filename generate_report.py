#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
抖音渠道分析看板 - 数据处理脚本
从Excel读取数据，计算KPI指标，生成HTML看板
"""

import os
import re
import json
import glob
from datetime import datetime

import pandas as pd

# ===== 配置 =====
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'docs')
TEMPLATE_FILE = os.path.join(SCRIPT_DIR, 'template.html')

# ===== 渠道映射 =====
# 根据实际数据，渠道名称-回溯家庭单后 列的值映射
CHANNEL_MAP = {
    '2.抖音投放': '抖音投放',
    '3.2抖音自然流': '抖音自然流',
    '1.小红书投放': '小红书基准',
    '3.1小红书自然流': '小红书基准',
}

CHANNEL_LIST = ['抖音投放', '抖音自然流', '小红书基准']


def find_latest_excel():
    """查找data目录下最新的xlsx文件"""
    xlsx_files = glob.glob(os.path.join(DATA_DIR, '*.xlsx'))
    if not xlsx_files:
        raise FileNotFoundError(f"未找到Excel文件，请将xlsx文件放入 {DATA_DIR} 目录")
    
    # 按修改时间排序
    xlsx_files.sort(key=os.path.getmtime, reverse=True)
    return xlsx_files[0]


def safe_float(value):
    """安全转换为浮点数"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, str):
        value = value.strip()
        if value.lower() in ['null', '', '-']:
            return 0.0
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def safe_int(value):
    """安全转换为整数"""
    return int(safe_float(value))


def parse_date(value):
    """解析日期，返回YYYY-MM-DD格式"""
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, str):
        value = value.strip()
        if not value or value.lower() in ['null', '-']:
            return None
        # 尝试解析各种日期格式
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y%m%d', '%m/%d/%Y']:
            try:
                return datetime.strptime(value, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
    return None


def load_data(excel_path):
    """加载Excel数据"""
    print(f"正在读取: {excel_path}")
    
    # 使用pandas读取
    df = pd.read_excel(excel_path, engine='openpyxl')
    print(f"总行数: {len(df)}")
    print(f"总列数: {len(df.columns)}")
    
    # 获取原始列名
    original_columns = list(df.columns)
    
    # 使用列索引（根据任务描述）
    # 0: 日期, 1: 中心, 2: 主管, 3: 顾问, 5: 家庭单回溯后订单进线时间
    # 7: 是否成交, 17: 例子id, 18: 家庭单回溯后订单号, 19: 是否删除好友
    # 20: 例子是否有效, 23: 渠道名称-回溯家庭单后, 27: 总规模保费
    # 52: 对话回合数, 60: 覆盖天次, 61: 回复天次
    # 65: 第一次电话, 66: 信息收集, 112: 是否预约一通, 114: 是否二通
    
    COL = {
        '日期': original_columns[0],
        '中心': original_columns[1],
        '主管': original_columns[2],
        '顾问': original_columns[3],
        '业务日期': original_columns[5],
        '是否成交': original_columns[7],
        '例子id': original_columns[17],
        '订单号': original_columns[18],
        '是否删除好友': original_columns[19],
        '例子是否有效': original_columns[20],
        '渠道来源': original_columns[23],
        '总规模保费': original_columns[27],
        '对话回合数': original_columns[52],
        '覆盖天次': original_columns[60],
        '回复天次': original_columns[61],
        '第一次电话': original_columns[65],
        '信息收集': original_columns[66],
        '是否预约一通': original_columns[112],
        '是否二通': original_columns[114],
        '是否转移成功': original_columns[14],
    }
    
    print(f"使用列映射: {COL}")
    
    # 解析数据
    records = []
    skipped = {'total': 0, 'invalid': 0, 'no_channel': 0, 'no_date': 0}
    
    for idx, row in df.iterrows():
        skipped['total'] += 1
        try:
            # 获取关键字段
            example_id = row[COL['例子id']]
            is_valid = safe_int(row[COL['例子是否有效']])
            channel_raw = row[COL['渠道来源']]
            
            # 去重：跳过无效例子
            if is_valid != 1:
                skipped['invalid'] += 1
                continue
            if pd.isna(example_id) or example_id is None:
                skipped['invalid'] += 1
                continue
            
            # 渠道映射
            channel_str = str(channel_raw).strip() if not pd.isna(channel_raw) else ''
            channel = CHANNEL_MAP.get(channel_str, None)
            if channel is None:
                skipped['no_channel'] += 1
                continue
            
            # 业务日期
            order_time = parse_date(row[COL['业务日期']])
            if order_time is None:
                skipped['no_date'] += 1
                continue
            
            record = {
                'example_id': str(example_id).strip(),
                '业务日期': order_time,
                '中心': str(row[COL['中心']]).strip() if not pd.isna(row[COL['中心']]) else '',
                '主管': str(row[COL['主管']]).strip() if not pd.isna(row[COL['主管']]) else '',
                '顾问': str(row[COL['顾问']]).strip() if not pd.isna(row[COL['顾问']]) else '',
                '渠道分类': channel,
                '订单号': str(row[COL['订单号']]).strip() if not pd.isna(row[COL['订单号']]) else '',
                '是否成交': safe_int(row[COL['是否成交']]),
                '总规模保费': safe_float(row[COL['总规模保费']]),
                '覆盖天次': safe_float(row[COL['覆盖天次']]),
                '回复天次': safe_float(row[COL['回复天次']]),
                '是否删除好友': '是' if str(row[COL['是否删除好友']]).strip() == '是' else '否',
                '对话回合数': safe_float(row[COL['对话回合数']]),
                '是否预约一通': safe_int(row[COL['是否预约一通']]),
                '第一次电话': safe_int(row[COL['第一次电话']]),
                '是否二通': safe_int(row[COL['是否二通']]),
                '是否转移成功': safe_int(row[COL['是否转移成功']]),
                '信息收集': safe_int(row[COL['信息收集']]),
            }
            records.append(record)
        except Exception as e:
            print(f"第{idx+2}行解析错误: {e}")
            continue
    
    print(f"跳过统计: 总计{skipped['total']}, 无效{skipped['invalid']}, 无渠道{skipped['no_channel']}, 无日期{skipped['no_date']}")
    print(f"有效记录数: {len(records)}")
    
    return records


def calculate_kpi_for_records(records, channel_filter=None):
    """计算指定渠道的KPI"""
    if channel_filter:
        filtered = [r for r in records if r['渠道分类'] == channel_filter]
    else:
        filtered = records
    
    if not filtered:
        return {
            'label': channel_filter or '全渠道',
            '进线量': 0, '订单数': 0, '成交订单数': 0, '成交率': 0,
            '平均覆盖天次': 0, '平均回复天次': 0, '微信删除率': 0,
            '平均沟通回合数': 0, '电话预约率': 0, '一通率': 0, '二通率': 0,
            '总规模保费': 0, '人均保费': 0
        }
    
    leads = len(filtered)
    
    # 订单去重
    order_dict = {}
    for r in filtered:
        order_id = r['订单号']
        if order_id and order_id not in order_dict:
            order_dict[order_id] = r
    
    orders = len(order_dict)
    deal_orders = sum(1 for r in order_dict.values() if r['是否成交'] == 1)
    
    # 均值类指标
    avg_cover = sum(r['覆盖天次'] for r in filtered) / leads if leads > 0 else 0
    avg_reply = sum(r['回复天次'] for r in filtered) / leads if leads > 0 else 0
    total_rounds = sum(r['对话回合数'] for r in filtered)
    avg_rounds = total_rounds / leads if leads > 0 else 0
    
    # 比率类指标
    delete_count = sum(1 for r in filtered if r['是否删除好友'] == '是')
    delete_rate = delete_count / leads * 100 if leads > 0 else 0
    
    reserve_count = sum(r['是否预约一通'] for r in filtered)
    reserve_rate = reserve_count / leads * 100 if leads > 0 else 0
    
    first_call_count = sum(r['第一次电话'] for r in filtered)
    first_call_rate = first_call_count / leads * 100 if leads > 0 else 0
    
    second_call_count = sum(r['是否二通'] for r in filtered)
    transfer_count = sum(r['是否转移成功'] for r in filtered)
    second_call_denominator = leads - transfer_count
    second_call_rate = second_call_count / second_call_denominator * 100 if second_call_denominator > 0 else 0
    
    deal_rate = deal_orders / orders * 100 if orders > 0 else 0
    
    # 保费
    total_premium = sum(r['总规模保费'] for r in filtered)
    avg_premium = total_premium / leads if leads > 0 else 0
    
    return {
        'label': channel_filter or '全渠道',
        '进线量': leads,
        '订单数': orders,
        '成交订单数': deal_orders,
        '成交率': round(deal_rate, 2),
        '平均覆盖天次': round(avg_cover, 2),
        '平均回复天次': round(avg_reply, 2),
        '微信删除率': round(delete_rate, 2),
        '平均沟通回合数': round(avg_rounds, 2),
        '电话预约率': round(reserve_rate, 2),
        '一通率': round(first_call_rate, 2),
        '二通率': round(second_call_rate, 2),
        '总规模保费': round(total_premium, 2),
        '人均保费': round(avg_premium, 2)
    }


def build_data_object(records):
    """构建完整的D数据结构"""
    # 计算各渠道KPI
    kpi = {}
    for ch in CHANNEL_LIST:
        kpi[ch] = calculate_kpi_for_records(records, ch)
    kpi['全渠道'] = calculate_kpi_for_records(records)
    
    # 获取数据日期
    data_date = datetime.now().strftime('%Y%m%d')
    if records:
        dates = [r['业务日期'] for r in records if r['业务日期']]
        if dates:
            data_date = dates[-1].replace('-', '')
    
    # 对比数据
    avg_metrics = ['平均覆盖天次', '平均回复天次', '平均沟通回合数']
    rate_metrics = ['微信删除率', '电话预约率', '一通率', '二通率', '成交率']
    all_metrics = avg_metrics + rate_metrics
    
    compare_avg = []
    for metric in avg_metrics:
        item = {'metric': metric}
        for ch in CHANNEL_LIST:
            item[ch] = kpi[ch].get(metric, 0)
        compare_avg.append(item)
    
    compare_rate = []
    for metric in rate_metrics:
        item = {'metric': metric}
        for ch in CHANNEL_LIST:
            item[ch] = kpi[ch].get(metric, 0)
        compare_rate.append(item)
    
    compare_all = compare_avg + compare_rate
    
    # 雷达图数据
    radar_indicators = [{'name': m, 'max': 100} for m in rate_metrics]
    radar = {
        'indicators': radar_indicators,
        '抖音投放': [kpi['抖音投放'].get(m, 0) for m in rate_metrics],
        '抖音自然流': [kpi['抖音自然流'].get(m, 0) for m in rate_metrics],
        '小红书基准': [kpi['小红书基准'].get(m, 0) for m in rate_metrics],
    }
    
    # 月度数据
    monthly_data = {}
    for r in records:
        month = r['业务日期'][:7]  # YYYY-MM
        if month not in monthly_data:
            monthly_data[month] = {'抖音投放': [], '抖音自然流': []}
        if r['渠道分类'] in ['抖音投放', '抖音自然流']:
            monthly_data[month][r['渠道分类']].append(r)
    
    monthly = []
    for month in sorted(monthly_data.keys()):
        item = {'month': month}
        for ch in ['抖音投放', '抖音自然流']:
            item[ch] = calculate_kpi_for_records(monthly_data[month][ch], ch)
        monthly.append(item)
    
    # 中心数据
    center_data = {}
    for r in records:
        center = r['中心']
        if center not in center_data:
            center_data[center] = {'抖音投放': [], '抖音自然流': []}
        if r['渠道分类'] in ['抖音投放', '抖音自然流']:
            center_data[center][r['渠道分类']].append(r)
    
    centers = []
    for center in sorted(center_data.keys()):
        item = {'center': center}
        for ch in ['抖音投放', '抖音自然流']:
            item['投放' if ch == '抖音投放' else '自然流'] = calculate_kpi_for_records(center_data[center][ch], ch)
        centers.append(item)
    
    # 团队数据
    team_data = {}
    for r in records:
        supervisor = r['主管']
        if supervisor not in team_data:
            team_data[supervisor] = {'抖音投放': [], '抖音自然流': []}
        if r['渠道分类'] in ['抖音投放', '抖音自然流']:
            team_data[supervisor][r['渠道分类']].append(r)
    
    teams = []
    for supervisor in sorted(team_data.keys()):
        item = {'supervisor': supervisor}
        for ch in ['抖音投放', '抖音自然流']:
            item['投放' if ch == '抖音投放' else '自然流'] = calculate_kpi_for_records(team_data[supervisor][ch], ch)
        teams.append(item)
    
    # 筛选选项
    centers_list = sorted(list(set(r['中心'] for r in records if r['中心'])))
    supervisors_list = sorted(list(set(r['主管'] for r in records if r['主管'])))
    consultants_list = sorted(list(set(r['顾问'] for r in records if r['顾问'])))
    
    filter_options = {
        'centers': centers_list,
        'supervisors': supervisors_list,
        'consultants': consultants_list,
    }
    
    return {
        'data_date': data_date,
        'kpi': kpi,
        'compare_avg': compare_avg,
        'compare_rate': compare_rate,
        'compare_all': compare_all,
        'radar': radar,
        'avg_metrics': avg_metrics,
        'rate_metrics': rate_metrics,
        'all_metrics': all_metrics,
        'monthly': monthly,
        'centers': centers,
        'teams': teams,
        'detail_count': len(records),
        'filter_options': filter_options,
    }


def build_detail_array(records):
    """构建明细数据数组"""
    detail = []
    for r in records:
        detail.append({
            '业务日期': r['业务日期'],
            '中心': r['中心'],
            '主管': r['主管'],
            '顾问': r['顾问'],
            '渠道分类': r['渠道分类'],
            '家庭单回溯后订单号': r['订单号'],
            '是否成交': r['是否成交'],
            '总规模保费': r['总规模保费'],
            '覆盖天次': r['覆盖天次'],
            '回复天次': r['回复天次'],
            '是否删除好友': r['是否删除好友'],
            '对话回合数': r['对话回合数'],
            '是否预约一通': r['是否预约一通'],
            '第一次电话': r['第一次电话'],
            '是否二通': r['是否二通'],
            '信息收集': r['信息收集'],
        })
    return detail


def replace_html_data(template_path, output_path, data_obj, detail_arr):
    """替换HTML模板中的数据"""
    print(f"正在读取模板: {template_path}")
    
    with open(template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # 序列化数据
    d_json = json.dumps(data_obj, ensure_ascii=False, indent=2)
    detail_json = json.dumps(detail_arr, ensure_ascii=False, indent=2)
    
    # 替换 var D = {...};
    # 不能简单用非贪婪匹配，因为JSON内有嵌套{}。用位置定位法。
    d_start_marker = 'var D = '
    d_start_idx = html_content.find(d_start_marker)
    if d_start_idx == -1:
        raise ValueError("模板中未找到 'var D = '")
    d_start_idx += len(d_start_marker)
    # 从 d_start_idx 开始，找匹配的 };
    # JSON对象的 } 后跟 ; 就是我们的结束
    # 简单方法：用json解析从d_start_idx开始的内容
    import json as _json
    # 找到从d_start_idx开始的完整JSON对象
    decoder = _json.JSONDecoder()
    _, end_idx = decoder.raw_decode(html_content[d_start_idx:])
    d_end_idx = d_start_idx + end_idx
    # 跳过后续的 ; 和换行
    while d_end_idx < len(html_content) and html_content[d_end_idx] in (' ', '\t'):
        d_end_idx += 1
    if d_end_idx < len(html_content) and html_content[d_end_idx] == ';':
        d_end_idx += 1
    
    d_replacement = 'var D = ' + d_json + ';'
    html_content = html_content[:d_start_idx - len(d_start_marker)] + d_replacement + html_content[d_end_idx:]
    
    # 替换 var DETAIL = [...];
    detail_start_marker = 'var DETAIL = '
    # 重新查找位置（因为前面的替换可能改变了偏移）
    detail_start_idx = html_content.find(detail_start_marker)
    if detail_start_idx == -1:
        raise ValueError("模板中未找到 'var DETAIL = '")
    detail_start_idx += len(detail_start_marker)
    # 找匹配的 ];
    # DETAIL是个数组，用 decoder
    _, detail_end_idx = decoder.raw_decode(html_content[detail_start_idx:])
    detail_end_idx = detail_start_idx + detail_end_idx
    while detail_end_idx < len(html_content) and html_content[detail_end_idx] in (' ', '\t'):
        detail_end_idx += 1
    if detail_end_idx < len(html_content) and html_content[detail_end_idx] == ';':
        detail_end_idx += 1
    
    detail_replacement = 'var DETAIL = ' + detail_json + ';'
    html_content = html_content[:detail_start_idx - len(detail_start_marker)] + detail_replacement + html_content[detail_end_idx:]
    
    # 确保docs目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    print(f"正在写入: {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"完成！输出文件大小: {os.path.getsize(output_path) / 1024:.1f} KB")


def main():
    """主函数"""
    print("=" * 50)
    print("抖音渠道分析看板 - 数据处理脚本")
    print("=" * 50)
    
    # 1. 查找最新的Excel文件
    excel_path = find_latest_excel()
    print(f"使用数据文件: {excel_path}")
    
    # 2. 加载数据
    records = load_data(excel_path)
    
    if not records:
        print("错误: 没有找到有效数据")
        return
    
    # 3. 构建数据对象
    print("正在计算KPI指标...")
    data_obj = build_data_object(records)
    
    # 4. 构建明细数组
    print("正在构建明细数据...")
    detail_arr = build_detail_array(records)
    
    # 5. 替换HTML数据
    print("正在生成HTML...")
    replace_html_data(TEMPLATE_FILE, os.path.join(OUTPUT_DIR, 'index.html'), data_obj, detail_arr)
    
    # 6. 输出摘要
    print("\n" + "=" * 50)
    print("数据摘要")
    print("=" * 50)
    print(f"数据日期: {data_obj['data_date']}")
    print(f"总记录数: {data_obj['detail_count']}")
    print("\n各渠道KPI:")
    for ch, kpi_data in data_obj['kpi'].items():
        print(f"  {ch}: 进线{kpi_data['进线量']} | 成交率{kpi_data['成交率']}% | 删除率{kpi_data['微信删除率']}%")
    print("\n完成！")


if __name__ == '__main__':
    main()
