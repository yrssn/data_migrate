"""
检查 ba_ptzcb_register 表中根据以下6个字段的重复数据:
  shopindex_id, shudi_id, platform_id, customer_id, information_id, information_part2_id
排除 platform_id = 108
重复数据按组分批导出到Excel，每组放在一起便于查看
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from datetime import datetime
from database import DatabaseManager
import pymysql
import pandas as pd


def main():
    # 加载数据源
    db_manager = DatabaseManager()
    datasources = db_manager.get_all_datasources()

    if not datasources:
        print("没有可用的数据源，请先在工具中添加数据源")
        return

    print("可用数据源:")
    for i, ds in enumerate(datasources):
        print(f"  [{i}] {ds.name} ({ds.host}:{ds.port}/{ds.database})")

    choice = input("\n请选择数据源编号: ").strip()
    try:
        ds = datasources[int(choice)]
    except (ValueError, IndexError):
        print("无效的选择")
        return

    print(f"\n连接到: {ds.name} ({ds.host}:{ds.port}/{ds.database})")

    connection = pymysql.connect(
        host=ds.host,
        port=ds.port,
        user=ds.username,
        password=ds.password,
        database=ds.database,
        charset=ds.charset
    )
    cursor = connection.cursor(pymysql.cursors.DictCursor)

    # 查找重复数据：按6个字段分组，组内数量>1的就是重复
    print("\n正在查询重复数据...")

    sql = """
        SELECT 
            r.id,
            r.project_id,
            r.shopindex_id,
            r.shudi_id,
            r.platform_id,
            r.customer_id,
            r.information_id,
            r.information_part2_id,
            r.company_name,
            r.country,
            r.status,
            r.create_time
        FROM ba_ptzcb_register r
        INNER JOIN (
            SELECT shopindex_id, shudi_id, platform_id, customer_id, 
                   information_id, information_part2_id
            FROM ba_ptzcb_register
            WHERE platform_id != 108
              AND delete_time IS NULL
            GROUP BY shopindex_id, shudi_id, platform_id, customer_id, 
                     information_id, information_part2_id
            HAVING COUNT(*) > 1
        ) dup ON r.shopindex_id <=> dup.shopindex_id
            AND r.shudi_id <=> dup.shudi_id
            AND r.platform_id <=> dup.platform_id
            AND r.customer_id <=> dup.customer_id
            AND r.information_id <=> dup.information_id
            AND r.information_part2_id <=> dup.information_part2_id
        WHERE r.platform_id != 108
          AND r.delete_time IS NULL
        ORDER BY r.shopindex_id, r.shudi_id, r.platform_id, r.customer_id, 
                 r.information_id, r.information_part2_id, r.id
    """

    cursor.execute(sql)
    rows = cursor.fetchall()
    cursor.close()
    connection.close()

    if not rows:
        print("没有找到重复数据！")
        return

    print(f"共找到 {len(rows)} 条重复记录")

    # 按组标记
    df = pd.DataFrame(rows)

    group_cols = ['shopindex_id', 'shudi_id', 'platform_id',
                  'customer_id', 'information_id', 'information_part2_id']

    # 给每组编个号
    group_id = 0
    prev_key = None
    group_ids = []
    for _, row in df.iterrows():
        key = tuple(row[c] for c in group_cols)
        if key != prev_key:
            group_id += 1
            prev_key = key
        group_ids.append(group_id)

    df.insert(0, '重复组号', group_ids)

    # 统计
    total_groups = df['重复组号'].nunique()
    print(f"共 {total_groups} 组重复，涉及 {len(df)} 条记录")

    # 导出
    output_file = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        f"ptzcb_duplicates_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\n已导出到: {output_file}")
    print(f"共 {total_groups} 组, {len(df)} 条记录")


if __name__ == '__main__':
    main()
