# -*- coding: utf-8 -*-
"""
臨時數據管理模組
用於保存和恢復用戶的數據輸入進度
"""

import json
import os
import tempfile
from datetime import datetime, timedelta

class TempDataManager:
    def __init__(self, base_path):
        """
        初始化臨時數據管理器
        
        Args:
            base_path (str): 基礎路徑
        """
        self.base_path = base_path
        self.temp_dir = os.path.join(tempfile.gettempdir(), 'oir_temp_data')
        
        # 確保臨時目錄存在
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
    
    def save_session_data(self, session_id, data):
        """
        保存session數據到臨時文件
        
        Args:
            session_id (str): session ID
            data (dict): 要保存的數據
        
        Returns:
            bool: 保存成功返回True
        """
        try:
            # 添加時間戳
            data['timestamp'] = datetime.now().isoformat()
            
            # 保存到臨時文件
            temp_file = os.path.join(self.temp_dir, f"session_{session_id}.json")
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2, default=str)
            
            return True
            
        except Exception as e:
            print(f"Error saving temp data: {e}")
            return False
    
    def load_session_data(self, session_id):
        """
        從臨時文件載入session數據
        
        Args:
            session_id (str): session ID
        
        Returns:
            dict: 載入的數據，如果失敗則返回None
        """
        try:
            temp_file = os.path.join(self.temp_dir, f"session_{session_id}.json")
            
            if not os.path.exists(temp_file):
                return None
            
            with open(temp_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 檢查文件是否過期（24小時）
            if 'timestamp' in data:
                timestamp = datetime.fromisoformat(data['timestamp'])
                if datetime.now() - timestamp > timedelta(hours=24):
                    # 刪除過期文件
                    os.remove(temp_file)
                    return None
            
            return data
            
        except Exception as e:
            print(f"Error loading temp data: {e}")
            return None
    
    def delete_session_data(self, session_id):
        """
        刪除session數據文件
        
        Args:
            session_id (str): session ID
        """
        try:
            temp_file = os.path.join(self.temp_dir, f"session_{session_id}.json")
            if os.path.exists(temp_file):
                os.remove(temp_file)
        except Exception as e:
            print(f"Error deleting temp data: {e}")
    
    def cleanup_old_files(self):
        """清理超過24小時的舊文件"""
        try:
            if not os.path.exists(self.temp_dir):
                return
            
            cutoff_time = datetime.now() - timedelta(hours=24)
            
            for filename in os.listdir(self.temp_dir):
                if filename.startswith('session_') and filename.endswith('.json'):
                    file_path = os.path.join(self.temp_dir, filename)
                    
                    # 檢查文件修改時間
                    file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if file_time < cutoff_time:
                        os.remove(file_path)
                        
        except Exception as e:
            print(f"Error cleaning up temp files: {e}")
    
    def get_temp_files_info(self):
        """
        獲取臨時文件資訊
        
        Returns:
            list: 臨時文件資訊列表
        """
        try:
            if not os.path.exists(self.temp_dir):
                return []
            
            files_info = []
            for filename in os.listdir(self.temp_dir):
                if filename.startswith('session_') and filename.endswith('.json'):
                    file_path = os.path.join(self.temp_dir, filename)
                    file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    file_size = os.path.getsize(file_path)
                    
                    files_info.append({
                        'filename': filename,
                        'path': file_path,
                        'modified': file_time.strftime('%Y-%m-%d %H:%M:%S'),
                        'size': file_size
                    })
            
            return files_info
            
        except Exception as e:
            print(f"Error getting temp files info: {e}")
            return []
