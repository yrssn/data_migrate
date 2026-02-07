"""
SQLite数据库管理
"""
import sqlite3
from typing import List, Optional
from models import DataSource
from cryptography.fernet import Fernet
import base64
import hashlib
class DatabaseManager:
    """SQLite数据库管理类"""
    
    def __init__(self, db_path: str = "data_migrate.db"):
        self.db_path = db_path
        self.cipher = self._get_cipher()
        self._init_database()
    
    def _get_cipher(self):
        """获取加密密钥"""
        # 使用固定密钥，实际项目中应该使用更安全的方式
        key = base64.urlsafe_b64encode(hashlib.sha256(b"data_migrate_secret_key").digest())
        return Fernet(key)
    
    def _init_database(self):
        """初始化数据库表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建数据源表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS datasources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                host TEXT NOT NULL,
                port INTEGER NOT NULL,
                username TEXT NOT NULL,
                password TEXT NOT NULL,
                database TEXT NOT NULL,
                charset TEXT DEFAULT 'utf8mb4',
                description TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def _encrypt_password(self, password: str) -> str:
        """加密密码"""
        return self.cipher.encrypt(password.encode()).decode()
    
    def _decrypt_password(self, encrypted_password: str) -> str:
        """解密密码"""
        return self.cipher.decrypt(encrypted_password.encode()).decode()
    
    def add_datasource(self, datasource: DataSource) -> int:
        """添加数据源"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        encrypted_password = self._encrypt_password(datasource.password)
        
        cursor.execute('''
            INSERT INTO datasources (name, host, port, username, password, database, charset, description)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (datasource.name, datasource.host, datasource.port, datasource.username,
              encrypted_password, datasource.database, datasource.charset, datasource.description))
        
        datasource_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return datasource_id
    
    def update_datasource(self, datasource: DataSource) -> bool:
        """更新数据源"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        encrypted_password = self._encrypt_password(datasource.password)
        
        cursor.execute('''
            UPDATE datasources
            SET name=?, host=?, port=?, username=?, password=?, database=?, charset=?, description=?, updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        ''', (datasource.name, datasource.host, datasource.port, datasource.username,
              encrypted_password, datasource.database, datasource.charset, datasource.description, datasource.id))
        
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        
        return success
    
    def delete_datasource(self, datasource_id: int) -> bool:
        """删除数据源"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM datasources WHERE id=?', (datasource_id,))
        
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        
        return success
    
    def get_datasource(self, datasource_id: int) -> Optional[DataSource]:
        """获取单个数据源"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM datasources WHERE id=?', (datasource_id,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return self._row_to_datasource(row)
        return None
    
    def get_all_datasources(self) -> List[DataSource]:
        """获取所有数据源"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM datasources ORDER BY created_at DESC')
        rows = cursor.fetchall()
        conn.close()
        
        return [self._row_to_datasource(row) for row in rows]
    
    def _row_to_datasource(self, row) -> DataSource:
        """将数据库行转换为DataSource对象"""
        decrypted_password = self._decrypt_password(row[5])
        
        return DataSource(
            id=row[0],
            name=row[1],
            host=row[2],
            port=row[3],
            username=row[4],
            password=decrypted_password,
            database=row[6],
            charset=row[7],
            description=row[8] or ""
        )
    
    def test_connection(self, datasource: DataSource) -> tuple[bool, str]:
        """测试数据库连接"""
        try:
            import pymysql
            
            connection = pymysql.connect(
                host=datasource.host,
                port=datasource.port,
                user=datasource.username,
                password=datasource.password,
                database=datasource.database,
                charset=datasource.charset,
                connect_timeout=5
            )
            connection.close()
            return True, "连接成功！"
        except Exception as e:
            return False, f"连接失败：{str(e)}"
