"""
数据模型定义
"""
from dataclasses import dataclass
from typing import Optional


@dataclass
class DataSource:
    """MySQL数据源模型"""
    id: Optional[int] = None
    name: str = ""
    host: str = "localhost"
    port: int = 3306
    username: str = ""
    password: str = ""
    database: str = ""
    charset: str = "utf8mb4"
    description: str = ""
    
    def to_dict(self):
        """转换为字典"""
        return {
            'id': self.id,
            'name': self.name,
            'host': self.host,
            'port': self.port,
            'username': self.username,
            'password': self.password,
            'database': self.database,
            'charset': self.charset,
            'description': self.description
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        """从字典创建对象"""
        return cls(**data)
