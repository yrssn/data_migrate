"""
导出 ba_ptzcb_register 表中满足以下条件的数据:
  - status = 6
  - platform_id 关联的 ba_platform.platform_type = 'opt1' (银行)
同时关联查出:
  - ba_platform.platform 作为银行名称
  - customer_id 关联的 ba_rlb_customer.legal_name 法人姓名
结果导出到Excel
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

    print("\n正在查询数据 (status=6 且平台类型为 opt1=银行)...")

    sql = """
        SELECT
            r.*,
            p.platform AS 银行名称,
            c.legal_name AS 法人姓名
        FROM ba_ptzcb_register r
        INNER JOIN ba_platform p ON r.platform_id = p.id
        LEFT JOIN ba_rlb_customer c ON r.customer_id = c.id
        WHERE r.status = 6
          AND p.platform_type = 'opt1'
          AND (r.delete_time IS NULL OR r.delete_time = 0)
        ORDER BY r.id
    """

    cursor.execute(sql)
    rows = cursor.fetchall()
    cursor.close()
    connection.close()

    if not rows:
        print("没有找到符合条件的数据！")
        return

    print(f"共找到 {len(rows)} 条记录")

    df = pd.DataFrame(rows)

    # 导出
    output_file = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        f"ptzcb_bank_status6_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\n已导出到: {output_file}")
    print(f"共 {len(df)} 条记录")


if __name__ == '__main__':
    main()
