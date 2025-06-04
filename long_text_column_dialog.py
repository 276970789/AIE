#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
长文本列配置对话框
用于设置新建长文本列的名称、文件名字段和搜索文件夹
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

class LongTextColumnDialog:
    def __init__(self, parent, table_manager, on_result_callback):
        print("DEBUG: LongTextColumnDialog __init__ called.") # Debug print
        self.parent = parent
        self.table_manager = table_manager
        self.on_result_callback = on_result_callback
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("新建长文本列")
        self.dialog.geometry("800x700")
        self.dialog.resizable(False, False)
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.center_window()
        self.create_widgets()
        
    def center_window(self):
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"800x700+{x}+{y}")
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 列名输入
        ttk.Label(main_frame, text="长文本列名:").pack(anchor=tk.W, pady=(0, 5))
        self.column_name_var = tk.StringVar()
        self.column_name_entry = ttk.Entry(main_frame, textvariable=self.column_name_var, width=50)
        self.column_name_entry.pack(fill=tk.X, pady=(0, 10))
        
        # 文件名字段选择
        ttk.Label(main_frame, text="文件名字段 (用于匹配文件): \n\n例如：如果你的表格中有一列叫 'arxiv_id'，其内容是 '2301.00001'，\n而你的长文本文件是 '2301.00001.pdf'，那么你就可以选择 'arxiv_id' 作为文件名字段。").pack(anchor=tk.W, pady=(0, 5))

        self.filename_field_var = tk.StringVar()
        self.filename_field_combo = ttk.Combobox(main_frame, textvariable=self.filename_field_var, 
                                                state="readonly", width=40)
        self.filename_field_combo.pack(fill=tk.X, pady=(0, 10))
        self.update_filename_field_list()
        
        # 搜索文件夹
        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(folder_frame, text="长文本文件搜索文件夹 (绝对或相对路径，默认为'论文原文'):").pack(anchor=tk.W)
        folder_input_frame = ttk.Frame(folder_frame)
        folder_input_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.folder_path_var = tk.StringVar(value="./论文原文") # 默认值，相对路径
        self.folder_path_entry = ttk.Entry(folder_input_frame, textvariable=self.folder_path_var)
        self.folder_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def browse_folder():
            folder = filedialog.askdirectory(title="选择长文本文件夹", 
                                           initialdir=self.folder_path_var.get() if self.folder_path_var.get() else ".")
            if folder:
                # 确保路径是相对路径，如果选择的文件夹在工作目录内
                workspace_root = os.getcwd()
                try:
                    relative_path = os.path.relpath(folder, workspace_root)
                    self.folder_path_var.set(relative_path)
                except ValueError:
                    # 如果不在工作目录内，使用绝对路径
                    self.folder_path_var.set(folder)
        
        ttk.Button(folder_input_frame, text="浏览", command=browse_folder, width=8).pack(side=tk.RIGHT, padx=(5, 0))
        
        # 预览长度
        ttk.Label(main_frame, text="预览长度 (在表格中显示的前N个字符): ").pack(anchor=tk.W, pady=(0, 5))
        self.preview_length_var = tk.StringVar(value="200")
        self.preview_length_entry = ttk.Entry(main_frame, textvariable=self.preview_length_var, width=10)
        self.preview_length_entry.pack(anchor=tk.W, pady=(0, 10))
        
        # 文件匹配预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="文件匹配预览", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 预览按钮
        preview_btn_frame = ttk.Frame(preview_frame)
        preview_btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.preview_btn = ttk.Button(preview_btn_frame, text="🔍 预览文件匹配", 
                                     command=self.preview_file_matches, width=15)
        self.preview_btn.pack(side=tk.LEFT)
        
        self.preview_status_label = ttk.Label(preview_btn_frame, text="", foreground="blue")
        self.preview_status_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # 匹配结果显示区域
        result_frame = ttk.Frame(preview_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # 使用Treeview显示匹配结果
        columns = ('filename', 'status', 'preview')
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=8)
        
        self.result_tree.heading('filename', text='文件名模式')
        self.result_tree.heading('status', text='匹配状态')
        self.result_tree.heading('preview', text='内容预览')
        
        self.result_tree.column('filename', width=120)
        self.result_tree.column('status', width=80)
        self.result_tree.column('preview', width=300)
        
        # 添加滚动条
        result_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=result_scrollbar.set)
        
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        cancel_btn = ttk.Button(button_frame, text="取消", command=self.on_cancel, width=10)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        ok_btn = ttk.Button(button_frame, text="确定", command=self.on_ok, width=10)
        ok_btn.pack(side=tk.RIGHT, padx=(0, 5))
        
        self.column_name_entry.focus()
        
        self.dialog.bind('<Return>', lambda e: self.on_ok())
        self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        
    def update_filename_field_list(self):
        """更新文件名字段列表"""
        if self.table_manager.get_dataframe() is not None:
            self.filename_field_combo['values'] = self.table_manager.get_column_names()
            if self.table_manager.get_column_names():
                self.filename_field_combo.set(self.table_manager.get_column_names()[0])
        else:
            self.filename_field_combo['values'] = []
            
    def validate_input(self):
        column_name = self.column_name_var.get().strip()
        filename_field = self.filename_field_var.get().strip()
        folder_path = self.folder_path_var.get().strip()
        preview_length_str = self.preview_length_var.get().strip()
        
        if not column_name:
            messagebox.showwarning("输入错误", "长文本列名不能为空")
            return False
            
        if column_name in self.table_manager.get_column_names():
            messagebox.showwarning("输入错误", f"列名 '{column_name}' 已存在")
            return False
            
        if not filename_field:
            messagebox.showwarning("输入错误", "文件名字段不能为空")
            return False
            
        if not folder_path:
            messagebox.showwarning("输入错误", "搜索文件夹不能为空")
            return False
            
        try:
            preview_length = int(preview_length_str)
            if preview_length <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("输入错误", "预览长度必须是大于0的整数")
            return False
            
        # 检查文件夹路径是否存在
        if not os.path.isdir(folder_path):
            messagebox.showwarning("输入错误", f"文件夹路径不存在或不是一个有效目录: {folder_path}")
            return False

        return True
        
    def preview_file_matches(self):
        """预览文件匹配结果"""
        filename_field = self.filename_field_var.get().strip()
        folder_path = self.folder_path_var.get().strip()
        preview_length_str = self.preview_length_var.get().strip()
        
        # 基本验证
        if not filename_field:
            self.preview_status_label.config(text="请先选择文件名字段", foreground="red")
            return
            
        if not folder_path:
            self.preview_status_label.config(text="请先选择搜索文件夹", foreground="red")
            return
            
        try:
            preview_length = int(preview_length_str)
            if preview_length <= 0:
                raise ValueError
        except ValueError:
            self.preview_status_label.config(text="预览长度必须是大于0的整数", foreground="red")
            return
            
        # 检查文件夹是否存在
        if not os.path.isdir(folder_path):
            self.preview_status_label.config(text=f"文件夹不存在: {folder_path}", foreground="red")
            return
            
        # 获取文件名字段的值
        if self.table_manager.get_dataframe() is None:
            self.preview_status_label.config(text="表格数据为空", foreground="red")
            return
            
        df = self.table_manager.get_dataframe()
        if filename_field not in df.columns:
            self.preview_status_label.config(text=f"字段 '{filename_field}' 不存在", foreground="red")
            return
            
        # 获取所有文件名模式
        filename_patterns = []
        for _, row in df.iterrows():
            value = str(row[filename_field]).strip() if row[filename_field] else ""
            if value and value not in filename_patterns:
                filename_patterns.append(value)
                
        if not filename_patterns:
            self.preview_status_label.config(text="没有找到有效的文件名模式", foreground="orange")
            return
            
        # 开始搜索
        self.preview_status_label.config(text="正在搜索文件...", foreground="blue")
        self.preview_btn.config(state="disabled")
        
        try:
            from paper_processor import get_paper_processor
            processor = get_paper_processor()
            
            # 搜索并预览文件
            results = processor.search_and_preview_files(folder_path, filename_patterns, preview_length)
            
            # 清空之前的结果
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
                
            # 显示结果
            total_patterns = len(filename_patterns)
            matched_count = results['matched_files']
            total_files = results['total_files']
            
            self.preview_status_label.config(
                text=f"找到 {total_files} 个文件，匹配 {matched_count}/{total_patterns} 个模式", 
                foreground="green" if matched_count > 0 else "orange"
            )
            
            # 添加匹配结果到树形控件
            for pattern, match_info in results['matches'].items():
                if match_info['found']:
                    status = "✓ 找到"
                    status_color = 'green'
                else:
                    status = "✗ 未找到"
                    status_color = 'red'
                    
                preview_text = match_info['preview']
                if len(preview_text) > 100:  # 在表格中显示时截断
                    preview_text = preview_text[:100] + "..."
                    
                item = self.result_tree.insert('', 'end', values=(pattern, status, preview_text))
                
                # 设置颜色标记
                if match_info['found']:
                    self.result_tree.set(item, 'status', status)
                else:
                    self.result_tree.set(item, 'status', status)
                    
        except Exception as e:
            self.preview_status_label.config(text=f"搜索失败: {str(e)}", foreground="red")
            
        finally:
            self.preview_btn.config(state="normal")
        
    def on_ok(self):
        if not self.validate_input():
            return
            
        column_name = self.column_name_var.get().strip()
        filename_field = self.filename_field_var.get().strip()
        folder_path = self.folder_path_var.get().strip()
        preview_length = int(self.preview_length_var.get().strip())
        
        self.result = {
            "column_name": column_name,
            "filename_field": filename_field,
            "folder_path": folder_path,
            "preview_length": preview_length
        }
        self.on_result_callback(self.result)
        self.dialog.destroy()
        
    def on_cancel(self):
        self.result = None
        self.on_result_callback(None)
        self.dialog.destroy()
        
    def show(self):
        print("DEBUG: LongTextColumnDialog show() called.") # Debug print
        self.parent.wait_window(self.dialog)