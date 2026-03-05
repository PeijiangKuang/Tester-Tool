"""
Tester Tool - CSV 解析模块
"""

import re
from datetime import datetime
from pathlib import Path


class CSVParser:
    """CSV 文件解析器"""
    
    def __init__(self):
        self.data = {}
        self.start_time = None
        self.end_time = None
    
    def parse(self, csv_path: str) -> dict:
        """
        解析 CSV 文件
        
        Args:
            csv_path: CSV 文件路径
            
        Returns:
            dict: {channel: [(time, value), ...]}
        """
        # 尝试多种编码
        encodings = ['utf-16', 'utf-8', 'gbk', 'gb2312']
        
        for encoding in encodings:
            try:
                with open(csv_path, 'r', encoding=encoding) as f:
                    lines = f.readlines()
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError(f"无法读取 CSV 文件: {csv_path}")
        
        # 找到数据开始行
        data_start_line = 0
        for i, line in enumerate(lines):
            if '扫描' in line:
                data_start_line = i + 1
                break
        
        if data_start_line >= len(lines):
            raise ValueError(f"CSV 文件格式不正确: {csv_path}")
        
        # 解析表头
        header_line = lines[data_start_line]
        headers = header_line.strip().split('\t')
        
        time_col = None
        channel_cols = {}
        
        for i, h in enumerate(headers):
            h = h.strip()
            if '时间' in h:
                time_col = i
            match = re.match(r'^(\d+)\s*\(C\)', h)
            if match:
                channel_cols[int(match.group(1))] = i
        
        # 初始化数据
        data = {ch: [] for ch in channel_cols.keys()}
        all_times = []
        
        # 解析数据行
        for line in lines[data_start_line + 1:]:
            if not line.strip():
                continue
            
            cols = line.strip().split('\t')
            if time_col is None or time_col >= len(cols):
                continue
            
            time_str = cols[time_col].strip()
            time_val = self.parse_time_string(time_str)
            
            if time_val is None:
                continue
            
            all_times.append(time_val)
            
            for ch, col_idx in channel_cols.items():
                if col_idx < len(cols):
                    try:
                        val = float(cols[col_idx].strip())
                        # 过滤异常值
                        if val > -1e10 and val < 1e10:
                            data[ch].append((time_val, val))
                    except (ValueError, IndexError):
                        pass
        
        # 按时间排序
        for ch in data:
            data[ch].sort(key=lambda x: x[0])
        
        self.data = data
        self.start_time = min(all_times) if all_times else None
        self.end_time = max(all_times) if all_times else None
        
        return data
    
    def parse_time_string(self, time_str: str) -> datetime | None:
        """
        解析时间字符串
        
        支持格式: 2024/1/15 10:30:15:0
        """
        # 多种时间格式支持
        patterns = [
            r'(\d{4})/(\d+)/(\d+)\s+(\d+):(\d+):(\d+):(\d+)',  # 2024/1/15 10:30:15:0
            r'(\d{4})-(\d+)-(\d+)\s+(\d+):(\d+):(\d+)',        # 2024-01-15 10:30:15
            r'(\d{4})/(\d+)/(\d+)\s+(\d+):(\d+):(\d+)',        # 2024/1/15 10:30:15
        ]
        
        for pattern in patterns:
            match = re.match(pattern, time_str.strip())
            if match:
                groups = match.groups()
                if len(groups) == 7:
                    year, month, day, hour, minute, second, ms = groups
                    return datetime(
                        int(year), int(month), int(day),
                        int(hour), int(minute), int(second),
                        int(ms) * 1000
                    )
                elif len(groups) == 6:
                    year, month, day, hour, minute, second = groups
                    return datetime(
                        int(year), int(month), int(day),
                        int(hour), int(minute), int(second)
                    )
        
        return None
    
    def get_time_range(self) -> tuple:
        """获取时间范围"""
        return self.start_time, self.end_time
    
    def get_data(self) -> dict:
        """获取解析后的数据"""
        return self.data
