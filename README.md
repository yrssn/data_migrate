# 数据库迁移工具

基于PyQt5开发的MySQL数据库迁移工具，使用SQLite存储数据源配置。

## 功能特性

- 数据源管理（添加、编辑、删除MySQL数据库连接）
- 连接测试功能
- Excel表格去重
- Excel导入客户表
- Excel导入法人信息
- 补全主体编号
- 补全平台详情信息
- 修改平台详情状态
- 现有客户银行信息补全
- 补全现有库存关联银行卡信息
- 导入平台注册部Excel初始化订单池
- 根据Excel补充系统平台注册部表信息
- 验证后的非平台注册部数据平台详情补全
- 账户分类关联财务系统（以财务为主，按名称关联/新增并绑定）
- 银行名称关联财务系统（以财务为主，按名称关联/新增并绑定，同步账户分类）
- 银行卡关联财务系统（以财务为主，按法人+卡号关联/新增并绑定，回写 finance_card_id 并同步账户分类）

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行

```bash
python main.py
```

## 项目结构

```
data-migrate/
├── main.py              # 应用入口
├── database.py          # SQLite数据库管理
├── models.py            # 数据模型
├── ui/
│   ├── __init__.py
│   ├── main_window.py   # 主窗口
│   └── datasource_dialog.py  # 数据源管理对话框
└── requirements.txt     # 依赖包
```
# 1. 安装 PyInstaller（如果还没安装）
pip install pyinstaller

# 2. 打包
pyinstaller --onefile --windowed --name="数据库迁移工具" main.py

# 3. 打包完成后，exe文件在 dist 目录下