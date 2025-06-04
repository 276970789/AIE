#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文处理器
负责处理各种格式的论文文件
"""

import os
import json

class PaperProcessor:
    def __init__(self):
        self.supported_formats = ['.txt', '.md', '.json', '.jsonl']
        
    def process_file(self, file_path):
        """处理文件并返回内容"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
            
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext in ['.txt', '.md']:
            return self._read_text_file(file_path)
        elif file_ext == '.json':
            return self._read_json_file(file_path)
        elif file_ext == '.jsonl':
            return self._read_jsonl_file(file_path)
        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")
    
    def _read_text_file(self, file_path):
        """读取文本文件"""
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
        
        raise UnicodeDecodeError("无法识别文件编码")
    
    def _read_json_file(self, file_path):
        """读取JSON文件"""
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    return json.load(f)
            except UnicodeDecodeError:
                continue
        
        raise UnicodeDecodeError("无法识别文件编码")
    
    def _read_jsonl_file(self, file_path):
        """读取JSONL文件"""
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312']
        
        for encoding in encodings:
            try:
                data_list = []
                with open(file_path, 'r', encoding=encoding) as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            data_list.append(json.loads(line))
                return data_list
            except UnicodeDecodeError:
                continue
        
        raise UnicodeDecodeError("无法识别文件编码")

    def find_files_in_folder(self, folder_path):
        """在文件夹中搜索支持的文件并返回文件名映射"""
        files_map = {}
        
        if not os.path.exists(folder_path):
            print(f"文件夹不存在: {folder_path}")
            return files_map
            
        try:
            # 递归搜索文件夹
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    file_name, file_ext = os.path.splitext(file)
                    
                    # 检查是否为支持的格式
                    if file_ext.lower() in self.supported_formats:
                        # 使用不带扩展名的文件名作为键
                        files_map[file_name] = file_path
                        
        except Exception as e:
            print(f"搜索文件夹失败: {e}")
            
        return files_map

    def batch_load_previews(self, filename_values, folder_path, preview_length=200):
        """批量加载文件预览"""
        previews = {}
        
        # 获取文件映射
        files_map = self.find_files_in_folder(folder_path)
        
        for filename_pattern in filename_values:
            if not filename_pattern:
                continue
                
            file_path = files_map.get(filename_pattern)
            if file_path:
                try:
                    content = self.get_file_content(file_path)
                    if content:
                        # 生成预览内容
                        preview = content[:preview_length]
                        if len(content) > preview_length:
                            preview += "..."
                        previews[filename_pattern] = preview
                    else:
                        previews[filename_pattern] = "[文件内容为空]"
                except Exception as e:
                    previews[filename_pattern] = f"[读取失败: {str(e)}]"
            else:
                previews[filename_pattern] = f"[未找到文件: {filename_pattern}]"
                
        return previews

    def get_file_content(self, file_path):
        """获取文件的完整内容"""
        try:
            return self.process_file(file_path)
        except Exception as e:
            print(f"读取文件内容失败 {file_path}: {e}")
            return None

    def search_and_preview_files(self, folder_path, filename_patterns, preview_length=200):
        """搜索文件并返回匹配结果预览"""
        results = {
            'total_files': 0,
            'matched_files': 0,
            'matches': {}
        }
        
        # 获取文件映射
        files_map = self.find_files_in_folder(folder_path)
        results['total_files'] = len(files_map)
        
        # 检查每个文件名模式的匹配情况
        for pattern in filename_patterns:
            if not pattern.strip():
                continue
                
            pattern = pattern.strip()
            if pattern in files_map:
                file_path = files_map[pattern]
                try:
                    content = self.get_file_content(file_path)
                    if content:
                        preview = content[:preview_length]
                        if len(content) > preview_length:
                            preview += "..."
                        results['matches'][pattern] = {
                            'found': True,
                            'file_path': file_path,
                            'preview': preview,
                            'size': len(content)
                        }
                        results['matched_files'] += 1
                    else:
                        results['matches'][pattern] = {
                            'found': True,
                            'file_path': file_path,
                            'preview': '[文件内容为空]',
                            'size': 0
                        }
                except Exception as e:
                    results['matches'][pattern] = {
                        'found': True,
                        'file_path': file_path,
                        'preview': f'[读取失败: {str(e)}]',
                        'size': 0
                    }
            else:
                results['matches'][pattern] = {
                    'found': False,
                    'file_path': None,
                    'preview': f'[未找到文件: {pattern}]',
                    'size': 0
                }
                
        return results

def get_paper_processor():
    """获取论文处理器实例"""
    return PaperProcessor()