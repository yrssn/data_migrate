"""
从 ba_rlb_legal_information.sql 中解析公司名，
找出因大小写、空格、全角/半角等差异导致的疑似重复数据，
导出到Excel。
"""
import re
import unicodedata
import pandas as pd
from collections import defaultdict


def parse_sql_inserts(sql_file):
    """从SQL文件解析出 (id, company_name, delete_time) 记录"""
    records = []
    with open(sql_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配 INSERT INTO ... VALUES (id, 'company_name', ...);
    # company_name 是第2个字段
    pattern = r"INSERT INTO `ba_rlb_legal_information` VALUES \((\d+),\s*'([^']*?)'"
    for m in re.finditer(pattern, content):
        rid = int(m.group(1))
        company_name = m.group(2)
        records.append({
            'id': rid,
            'company_name': company_name,
        })
    return records


def normalize_name(name):
    """标准化公司名用于比较：
    - 全角转半角
    - 统一为小写
    - 去除所有空白字符（空格、全角空格、\t等）
    - NFKC标准化（统一兼容字符）
    """
    if not name:
        return ''
    # NFKC标准化：全角字母/数字转半角，兼容字符统一
    s = unicodedata.normalize('NFKC', name)
    # 转小写
    s = s.lower()
    # 去除所有空白字符（包括全角空格 \u3000）
    s = re.sub(r'[\s\u3000]+', '', s)
    return s


def find_similar_companies(records):
    """根据标准化后的公司名分组，找出疑似重复"""
    groups = defaultdict(list)
    for rec in records:
        norm = normalize_name(rec['company_name'])
        if norm:  # 跳过空名
            groups[norm].append(rec)

    # 只保留有多条记录的组（即疑似重复）
    duplicates = {k: v for k, v in groups.items() if len(v) > 1}
    return duplicates


def main():
    sql_file = r'C:\Users\Administrator\Desktop\ba_rlb_legal_information.sql'
    output_file = r'C:\Users\Administrator\Desktop\疑似相同公司名分析.xlsx'

    print("正在解析SQL文件...")
    records = parse_sql_inserts(sql_file)
    print(f"共解析出 {len(records)} 条记录")

    print("正在分析疑似相同公司名...")
    duplicates = find_similar_companies(records)

    if not duplicates:
        print("未发现疑似相同的公司名！")
        return

    print(f"发现 {len(duplicates)} 组疑似相同的公司名")

    # 构造输出数据
    rows = []
    group_num = 0
    for norm_name, recs in sorted(duplicates.items(), key=lambda x: x[0]):
        group_num += 1
        for rec in recs:
            rows.append({
                '分组编号': group_num,
                'ID': rec['id'],
                '原始公司名': rec['company_name'],
                '标准化后': norm_name,
                '差异说明': _diff_explain(rec['company_name'], norm_name),
            })

    df = pd.DataFrame(rows)
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\n结果已导出到: {output_file}")
    print(f"共 {group_num} 组, {len(rows)} 条记录")

    # 打印摘要
    print("\n===== 疑似重复摘要 =====")
    group_num = 0
    for norm_name, recs in sorted(duplicates.items(), key=lambda x: x[0]):
        group_num += 1
        print(f"\n--- 第{group_num}组 (标准化: {norm_name}) ---")
        for rec in recs:
            print(f"  ID={rec['id']}: 「{rec['company_name']}」")


def _diff_explain(original, normalized):
    """简单说明原始名和标准化名之间的差异"""
    diffs = []
    if original != original.strip():
        diffs.append("首尾有空格")
    if re.search(r'[\s\u3000]', original):
        if re.search(r'\u3000', original):
            diffs.append("含全角空格")
        if re.search(r'[ \t]', original.strip()):
            diffs.append("含半角空格")
    # 检查全角字母/数字
    if re.search(r'[\uff01-\uff5e]', original):
        diffs.append("含全角字符")
    # 检查大小写
    lower = original.lower()
    if lower != original:
        diffs.append("含大写字母")
    return ', '.join(diffs) if diffs else '完全相同'


if __name__ == '__main__':
    main()
