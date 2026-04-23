# -*- coding: utf-8 -*-
"""
Excel解析与列选择逻辑
负责处理Excel文件的读取、列识别和姓名列表提取
"""
import pandas as pd
from typing import List, Tuple, Dict, Optional
from io import BytesIO
import uuid


class ExcelParser:
    """Excel解析器"""
    
    def __init__(self):
        self.df: Optional[pd.DataFrame] = None
        self.columns: List[str] = []
        self.has_header: bool = True
    
    def read_excel(self, file_bytes: BytesIO) -> Dict:
        """
        读取Excel文件并返回预览数据
        
        Args:
            file_bytes: Excel文件字节流
            
        Returns:
            Dict: 包含列名、预览数据和原始DataFrame的字典
        """
        try:
            # 先尝试读取有表头的情况
            self.df = pd.read_excel(file_bytes, engine='openpyxl')
            self.columns = list(self.df.columns)
            self.has_header = True
            
            # 检查是否可能没有表头（第一行看起来像数据而不是列名）
            # 如果列名看起来像默认的数字列名，可能没有表头
            if self._is_likely_no_header():
                self.has_header = False
                file_bytes.seek(0)
                self.df = pd.read_excel(file_bytes, header=None, engine='openpyxl')
                # 使用A列、B列...作为列名
                self.columns = [self._get_column_letter(i) for i in range(len(self.df.columns))]
                self.df.columns = self.columns
            
            return {
                'success': True,
                'columns': self.columns,
                'preview': self.df.head(10).to_dict('records'),
                'total_rows': len(self.df),
                'has_header': self.has_header
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': f'Excel读取失败: {str(e)}',
                'columns': [],
                'preview': [],
                'total_rows': 0,
                'has_header': True
            }
    
    def _is_likely_no_header(self) -> bool:
        """
        判断Excel是否可能没有表头
        """
        if self.df is None or len(self.df) == 0:
            return False
        
        # 检查列名是否都是字符串类型且看起来不像默认列名
        # 如果列名都是数字或看起来像数据，可能没有表头
        columns = list(self.df.columns)
        
        # 如果只有一行数据，可能没有表头
        if len(self.df) == 1:
            return True
        
        # 检查列名是否都是类似 Unnamed: 0 的格式
        all_unnamed = all(str(col).startswith('Unnamed:') for col in columns)
        if all_unnamed:
            return True
        
        # 检查第一行数据是否与列名类型一致
        # 如果第一行数据看起来和列名很像，可能没有表头
        return False
    
    def _get_column_letter(self, index: int) -> str:
        """
        将数字索引转换为Excel列字母（A, B, ..., Z, AA, AB...）
        """
        letter = ''
        while index >= 0:
            letter = chr(65 + (index % 26)) + letter
            index = index // 26 - 1
        return letter
    
    def extract_names(self, column_name: str) -> Tuple[List[Dict], Dict]:
        """
        从指定列提取姓名列表
        
        Args:
            column_name: 要提取的列名
            
        Returns:
            Tuple[List[Dict], Dict]: (姓名字典列表, 统计信息字典)
        """
        if self.df is None:
            return [], {'error': '未加载Excel文件'}
        
        if column_name not in self.df.columns:
            return [], {'error': f'列名"{column_name}"不存在'}
        
        names = []
        empty_count = 0
        duplicate_names = []
        name_counts = {}
        
        for idx, value in enumerate(self.df[column_name]):
            # 处理空值
            if pd.isna(value) or str(value).strip() == '':
                empty_count += 1
                continue
            
            # 去除前后空格并转成字符串
            name = str(value).strip()
            
            # 生成唯一person_id
            person_id = str(uuid.uuid4())
            
            names.append({
                'person_id': person_id,
                'name': name,
                'order_index': len(names),  # 新的顺序索引
                'current_seat_id': None
            })
            
            # 统计重复姓名
            if name in name_counts:
                name_counts[name] += 1
                if name not in duplicate_names:
                    duplicate_names.append(name)
            else:
                name_counts[name] = 1
        
        stats = {
            'total_rows': len(self.df),
            'empty_count': empty_count,
            'valid_names': len(names),
            'has_duplicates': len(duplicate_names) > 0,
            'duplicate_names': duplicate_names,
            'empty_ratio': empty_count / len(self.df) if len(self.df) > 0 else 0
        }
        
        # 检查空值比例是否过高（超过30%）
        if stats['empty_ratio'] > 0.3:
            stats['warning'] = f'所选列空值比例较高（{int(stats["empty_ratio"]*100)}%），请确认是否选择正确的列'
        
        return names, stats
    
    def get_columns(self) -> List[str]:
        """
        获取所有列名
        """
        return self.columns
