#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
订单自动转换脚本 - 清爽极速版
功能：将客户下单表转换为标准格式，支持淘宝/拼多多自动区分
      支持 .xls 和 .xlsx 格式
      自动补纸功能（每台机器+1卷纸）
"""

import pandas as pd
import re
import os
from datetime import datetime

# ============================================================================
# 配置文件路径
# ============================================================================
SOURCE_FILE = '商颂.xls'
TEMPLATE_FILE = 'config/订单分表模版.xlsx'
EXPRESS_FILE = 'config/快递价格明细表.xls'
OUTPUT_FILE = f'商颂订单_{datetime.now().strftime("%Y-%m-%d")}.xlsx'

# ============================================================================
# 控制台排版工具 (开放式布局，放弃右边框，彻底解决错位)
# ============================================================================
class TxtFormater:
    @staticmethod
    def get_width(text):
        """获取控制台显示宽度 (利用 gbk 编码完美模拟 Windows cmd 的中文字符宽度)"""
        try:
            return len(str(text).encode('gbk', errors='replace'))
        except:
            return len(str(text))

    @staticmethod
    def pad_str(text, width, align='left'):
        """按控制台视觉宽度填充空格"""
        text = str(text)
        w = TxtFormater.get_width(text)
        padding = ' ' * max(0, width - w)
        if align == 'left':
            return text + padding
        else:
            return padding + text

def print_header():
    print("\n" + "=" * 60)
    print(" " * 18 + "订单自动转换脚本 v2.1")
    print("\n")
    print(" " * 21 + "Power by FTH")
    print("=" * 60 + "\n")

def print_section(title, step_num=None):
    prefix = f">>> 步骤 {step_num}：" if step_num else ">>> "
    print(f"\n{prefix}{title}")
    print("-" * 60)

def print_item(label, value):
    print(f"  • {label}：{value}")

def print_summary(stats):
    print("\n" + "=" * 60)
    print(" [ 转换统计报告 ]")
    print("-" * 60)
    
    print("  【订单统计】")
    print(f"    淘宝订单:   {stats['taobao']:>4} 条")
    print(f"    拼多多订单: {stats['pdd']:>4} 条")
    print(f"    需收邮资:   {stats['postage_orders']:>4} 条")
    print(f"    -----------------")
    print(f"    订单总数:   {stats['total']:>4} 条\n")
    
    print("  【金额统计】")
    print(f"    货品金额:   {stats['product_amount']:>8.2f} 元")
    print(f"    邮资金额:   {stats['postage_amount']:>8.2f} 元")
    print(f"    -----------------")
    print(f"    合计金额:   {stats['total_amount']:>8.2f} 元")
    print("=" * 60)

# ============================================================================
# 数据加载函数
# ============================================================================
def load_source_file():
    orders =[]
    source_file = SOURCE_FILE
    xlsx_file = SOURCE_FILE.replace('.xls', '.xlsx')
    if os.path.exists(xlsx_file):
        source_file = xlsx_file
    elif os.path.exists(SOURCE_FILE):
        source_file = SOURCE_FILE
    else:
        print(f'  [错误] 找不到 {SOURCE_FILE} 或 {xlsx_file}')
        return orders

    if source_file.endswith('.xlsx'):
        df = pd.read_excel(source_file, header=None, engine='openpyxl')
    else:
        df = pd.read_excel(source_file, header=None)

    print(f'  • 正在读取 {source_file}...')

    current_platform = None
    for idx in range(len(df)):
        row = df.iloc[idx]
        if pd.isna(row.iloc[0]) and (len(row) <= 1 or pd.isna(row.iloc[1])):
            continue

        col0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        col1 = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
        col2 = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ''
        col3 = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else '' 

        if '订单' in col0:
            if '天猫' in col0 or '淘宝' in col0: current_platform = '淘宝'
            elif '多多' in col0: current_platform = '拼多多'
            continue

        if not col0 or col0 == 'nan': continue

        order_id = col0
        if re.match(r'^\d{6}-\d+$', order_id): platform = '拼多多'
        elif re.match(r'^\d{12,}$', order_id): platform = '淘宝'
        else: continue  

        orders.append({
            'order_id': order_id, 'address': col2, 'interface': col1,
            'platform': platform, 'remark': col3 if col3 and col3 != 'nan' else ''
        })
    return orders

def load_template():
    df_order = pd.read_excel(TEMPLATE_FILE, sheet_name='订单')
    df_product = pd.read_excel(TEMPLATE_FILE, sheet_name='货品')
    product_map = {}
    printer_paper_width_map = {}
    paper_by_width_map = {}

    for idx, row in df_product.iterrows():
        if pd.notna(row['货品名称']) and pd.notna(row['单价']):
            product_info = {
                'name': row['货品名称'], 'code': row.get('货品编号', ''),
                'barcode': row.get('条码', ''), 'spec': row.get('规格', ''),
                'price': float(row['单价'])
            }
            # 添加纸张宽度信息
            if '纸张宽度' in row and pd.notna(row['纸张宽度']):
                product_info['paper_width'] = float(row['纸张宽度'])
            else:
                product_info['paper_width'] = None

            product_map[row['货品名称']] = product_info
            if product_info['spec'] and product_info['spec'] != 'nan':
                product_map[product_info['spec']] = product_info

            # 建立打印机到纸张宽度的映射
            if product_info['paper_width'] and '纸' not in row['货品名称']:
                # 这是打印机，记录其纸张宽度
                printer_paper_width_map[row['货品名称']] = product_info['paper_width']
            elif '纸' in row['货品名称'] and product_info['paper_width']:
                # 这是纸张，记录该宽度对应的纸张货品
                width = product_info['paper_width']
                if width not in paper_by_width_map:
                    paper_by_width_map[width] = product_info

    default_values = {}
    for col in ['业务员', '物流公司', '客户账号', '销售渠道名称', '结算方式']:
        if col in df_order.columns and len(df_order) > 0:
            for idx in range(len(df_order)):
                if pd.notna(df_order.iloc[idx][col]):
                    default_values[col] = df_order.iloc[idx][col]
                    break
    return df_order, product_map, default_values, printer_paper_width_map, paper_by_width_map

def load_express_rules():
    df = pd.read_excel(EXPRESS_FILE, header=None)
    express_rules, postage_rules = {}, {}
    current_platform = None
    
    for idx, row in df.iterrows():
        first_col = str(row.iloc[0]).strip()
        if first_col == '淘宝订单': current_platform = '淘宝'; continue
        elif first_col == '拼多多订单': current_platform = '拼多多'; continue
        if first_col == '省份' or first_col == 'nan': continue
        if current_platform is None: current_platform = '淘宝'

        province = first_col
        postage = str(row.iloc[3]).strip()
        postage = 0.0 if postage in ('nan', '') else float(postage.replace('元', '').strip() or 0)

        key = (province, current_platform)
        express_rules[key] = {
            '1_2kg': str(row.iloc[1]).strip(),
            '2kg_plus': str(row.iloc[2]).strip()
        }
        postage_rules[key] = postage
    return express_rules, postage_rules

# ============================================================================
# 核心解析函数
# ============================================================================
def parse_machine_interfaces(interface_str):
    if pd.isna(interface_str) or interface_str == '': return[]
    lines = str(interface_str).split('\n')
    items =[]
    for line in lines:
        line = line.strip()
        if not line: continue
        ip_addresses = re.findall(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', line)
        line_without_ip = re.sub(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', '', line).strip()
        match = re.search(r'^(.+?)\s*\*\s*(\d+)\s*$', line_without_ip)
        
        if match: items.append({'name': match.group(1).strip(), 'quantity': int(match.group(2)), 'ips': ip_addresses})
        else: items.append({'name': line_without_ip, 'quantity': 1, 'ips': ip_addresses})
    return items

def is_free_item(item_name):
    return any(free_item in item_name for free_item in['送网线', '网线', '开收据', '收据', '测试纸', '顺丰到付'])

def find_best_match(product_name, product_map):
    def normalize(s): return s.replace('（', '(').replace('）', ')').upper().replace(' ', '')
    if product_name in product_map: return product_map[product_name]
    product_normalized = normalize(product_name)
    for key, product in product_map.items():
        if normalize(key) == product_normalized: return product
    for key, product in product_map.items():
        if key in product_name or product_name in key: return product
    for key, product in product_map.items():
        if (normalize(key) in product_normalized) or (product_normalized in normalize(key)): return product
    return None

def extract_province(address):
    address = address.replace('\n', ' ').strip()
    province_map = {
        '北京':'北京', '天津':'天津', '上海':'上海', '重庆':'重庆', '河北':'河北', '山西':'山西',
        '辽宁':'辽宁', '吉林':'吉林', '黑龙江':'黑龙江', '江苏':'江苏', '浙江':'浙江', '安徽':'安徽',
        '福建':'福建', '江西':'江西', '山东':'山东', '河南':'河南', '湖北':'湖北', '湖南':'湖南',
        '广东':'广东', '海南':'海南', '四川':'四川', '贵州':'贵州', '云南':'云南', '陕西':'陕西',
        '甘肃':'甘肃', '青海':'青海', '台湾':'台湾', '内蒙古':'内蒙古', '广西':'广西', '西藏':'西藏',
        '宁夏':'宁夏', '新疆':'新疆', '香港':'香港', '澳门':'澳门'
    }
    for key in province_map:
        if key in address: return province_map[key]
    return None

def calculate_weight_class(products):
    printer_count = sum(p['quantity'] for p in products if p['price'] > 50)
    return '2kg_plus' if printer_count >= 2 else '1_2kg'

def match_express_and_postage(address, platform, products, express_rules, postage_rules):
    province = extract_province(address)
    if not province: return '', 0.0
    weight_class = calculate_weight_class(products)
    key = (province, platform)
    if key in express_rules: return express_rules[key][weight_class], postage_rules.get(key, 0.0)
    else: return '', 0.0

def add_paper_if_needed(interface_items, product_map, printer_paper_width_map, paper_by_width_map, order_id):
    # 统计各种宽度的打印机数量和纸张数量
    printer_by_width = {}
    paper_by_width_count = {}

    for item in interface_items:
        product = find_best_match(item['name'], product_map)
        if product:
            # 如果是打印机（价格>50且名称不含"纸"）
            if product['price'] > 50 and '纸' not in product['name']:
                # 获取打印机的纸张宽度
                paper_width = printer_paper_width_map.get(product['name'])
                if paper_width:
                    printer_by_width[paper_width] = printer_by_width.get(paper_width, 0) + item['quantity']
                else:
                    # 如果没有找到纸张宽度，默认使用80
                    printer_by_width[80.0] = printer_by_width.get(80.0, 0) + item['quantity']

            # 如果是纸张
            elif '纸' in product['name'] and product.get('paper_width'):
                paper_width = product['paper_width']
                paper_by_width_count[paper_width] = paper_by_width_count.get(paper_width, 0) + item['quantity']

    # 计算需要补充的纸张
    papers_to_add_by_width = {}
    for width, printer_count in printer_by_width.items():
        paper_count = paper_by_width_count.get(width, 0)
        papers_to_add = max(0, printer_count - paper_count)
        if papers_to_add > 0:
            papers_to_add_by_width[width] = papers_to_add

    # 如果需要补充纸张，返回第一个宽度的纸张（简化处理）
    if papers_to_add_by_width:
        # 返回需要添加最多的那个宽度
        width = max(papers_to_add_by_width.items(), key=lambda x: x[1])[0]
        papers_to_add = papers_to_add_by_width[width]
        paper_product = paper_by_width_map.get(width)

        if paper_product and papers_to_add > 0:
            width_desc = f"{int(width)}mm" if width == int(width) else f"{width}mm"
            print(f"  [补纸] 订单 {order_id[-8:]} 自动加纸 {papers_to_add} 卷 ({width_desc})")
            return paper_product, papers_to_add

    return None, 0

# ============================================================================
# 主函数
# ============================================================================
def generate_orders():
    print_header()
    
    # === 步骤 1 ===
    print_section("加载客户下单表", 1)
    orders = load_source_file()
    if not orders:
        print("  [错误] 未找到任何订单，请检查下单表文件")
        return
    print_item("订单总数", f"{len(orders)} 条")
    print_item("淘宝订单", f"{sum(1 for o in orders if o['platform'] == '淘宝')} 条")
    print_item("拼多多订单", f"{sum(1 for o in orders if o['platform'] == '拼多多')} 条")

    # === 步骤 2 ===
    print_section("加载模板文件", 2)
    df_order, product_map, default_values, printer_paper_width_map, paper_by_width_map = load_template()
    print_item("货品映射", f"{len(product_map)} 个")
    print_item("打印机纸张宽度映射", f"{len(printer_paper_width_map)} 个")
    print_item("默认业务员", default_values.get("业务员", "未设置"))
    print_item("默认物流", default_values.get("物流公司", "未设置"))

    # === 步骤 3 ===
    print_section("加载快递规则", 3)
    express_rules, postage_rules = load_express_rules()
    print_item("规则总数", f"{len(express_rules)} 条")

    # === 步骤 4 ===
    print_section("转换订单数据 (处理明细)", 4)
    order_rows, product_rows = [], []
    total_amount, total_postage = 0.0, 0.0
    unmatched_products, unmatched_provinces = [], []
    order_stats = {'taobao': 0, 'pdd': 0, 'postage_orders': 0, 'total': len(orders)}

    for idx, order in enumerate(orders, 1):
        interfaces = parse_machine_interfaces(order['interface'])
        matched_products, all_ip_addresses = [], []
        product_total = 0.0

        for item in interfaces:
            if is_free_item(item['name']): continue
            product = find_best_match(item['name'], product_map)
            if product:
                price = product['price'] * item['quantity']
                product_total += price
                total_amount += price
                if item['ips']: all_ip_addresses.extend(item['ips'])

                product_rows.append({
                    '导入编号(关联订单)': order['order_id'], '货品名称': product['name'],
                    '条码': product['barcode'], '货品编号': product['code'],
                    '规格': product['spec'], '数量': item['quantity'], '单价': product['price']
                })
                matched_products.append({'name': product['name'], 'price': product['price'], 'quantity': item['quantity']})
            else:
                unmatched_products.append({'order_id': order['order_id'], 'product': item['name'], 'quantity': item['quantity']})

        # 补纸
        paper_product, papers_to_add = add_paper_if_needed(interfaces, product_map, printer_paper_width_map, paper_by_width_map, order['order_id'])
        if paper_product and papers_to_add > 0:
            price = paper_product['price'] * papers_to_add
            product_total += price
            total_amount += price
            product_rows.append({
                '导入编号(关联订单)': order['order_id'], '货品名称': paper_product['name'],
                '条码': paper_product['barcode'], '货品编号': paper_product['code'],
                '规格': paper_product['spec'], '数量': papers_to_add, '单价': paper_product['price']
            })
            matched_products.append({'name': paper_product['name'], 'price': paper_product['price'], 'quantity': papers_to_add})

        # 快递
        express_company, postage = match_express_and_postage(order['address'], order['platform'], matched_products, express_rules, postage_rules)
        province = extract_province(order['address'])
        if not province: unmatched_provinces.append({'order_id': order['order_id'], 'address': order['address'][:25] + '...'})

        total_postage += postage
        total_with_postage = product_total + postage
        
        if order['platform'] == '淘宝': order_stats['taobao'] += 1
        else: order_stats['pdd'] += 1
        if postage > 0: order_stats['postage_orders'] += 1

        # 备注拼装
        merged_products = {}
        for p in matched_products:
            merged_products[p['name']] = merged_products.get(p['name'], 0) + p.get('quantity', 1)
        cs_remark = '+'.join([f"{n}*{q}" for n, q in merged_products.items()])
        extra_remarks = []
        if all_ip_addresses: extra_remarks.append(f'改IP：{", ".join(set(all_ip_addresses))}')
        if order.get('remark', '').strip(): extra_remarks.append(order['remark'].strip())
        if extra_remarks: cs_remark += ' ' + ' '.join(extra_remarks)

        # 写入订单行
        order_row = {
            '导入编号': order['order_id'], '收货人': '', '手机': '', '收货地址': '',
            '收货人信息(解析)': order['address'], '应收邮资': postage,
            '应收合计': total_with_postage, '客服备注': cs_remark, '物流公司': express_company
        }
        for k, v in default_values.items(): order_row.setdefault(k, v)
        order_rows.append(order_row)

        # 打印单行日志 (使用工具类对齐)
        plat_tag = "[淘宝]" if order['platform'] == '淘宝' else "[多多]"
        id_str = str(order['order_id'][-8:])
        idx_str = f"{idx:>2}/{len(orders)}"
        p_str = TxtFormater.pad_str(province or "未知", 6)
        exp_str = TxtFormater.pad_str(express_company or "无物流", 16)
        post_tag = f"(+{postage}邮)" if postage > 0 else ""
        
        print(f"  {plat_tag} {idx_str} | {id_str} | {p_str} | {exp_str} | ¥{total_with_postage:<6.1f} {post_tag}")

    # === 步骤 5 ===
    print_section("生成最终文件", 5)
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        pd.DataFrame(order_rows).to_excel(writer, sheet_name='订单', index=False)
        pd.DataFrame(product_rows).to_excel(writer, sheet_name='货品', index=False)
        
    print_item("输出文件", OUTPUT_FILE)
    print_item("生成订单", f"{len(order_rows)} 条")
    print_item("生成货品明细", f"{len(product_rows)} 条")
    
    # === 统计汇总 ===
    order_stats.update({
        'total': len(order_rows), 'product_amount': total_amount - total_postage,
        'postage_amount': total_postage, 'total_amount': total_amount
    })
    print_summary(order_stats)
    
    # 异常警告汇总
    if unmatched_products or unmatched_provinces:
        print("\n  [!] 警告信息汇总：")
        for p in unmatched_products:
            print(f"    - 未匹配货品: 订单 {p['order_id'][-8:]} -> {p['product']} x{p['quantity']}")
        for p in unmatched_provinces:
            print(f"    - 未识别省份: 订单 {p['order_id'][-8:]} -> {p['address']}")
    
    print(f"\n  >>> 转换完成！请查看文件: {OUTPUT_FILE}\n")

if __name__ == '__main__':
    try:
        generate_orders()
    except Exception as e:
        print(f'\n  [❌] 严重错误：{str(e)}')
        import traceback
        traceback.print_exc()
    finally:
        input('  按 Enter 键退出...')
