#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
查找替换对话框
支持全表或选定列的查找和替换功能
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re

class FindReplaceDialog:
    def __init__(self, parent, table_manager):
        self.parent = parent
        self.table_manager = table_manager
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("查找和替换")
        self.dialog.geometry("500x450")
        self.dialog.resizable(True, True)
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="查找和替换", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # 查找输入
        find_frame = ttk.LabelFrame(main_frame, text="查找", padding="10")
        find_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.find_var = tk.StringVar()
        find_entry = ttk.Entry(find_frame, textvariable=self.find_var, width=50)
        find_entry.pack(fill=tk.X, pady=5)
        find_entry.focus()
        
        # 替换输入
        replace_frame = ttk.LabelFrame(main_frame, text="替换为", padding="10")
        replace_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.replace_var = tk.StringVar()
        replace_entry = ttk.Entry(replace_frame, textvariable=self.replace_var, width=50)
        replace_entry.pack(fill=tk.X, pady=5)
        
        # 选项设置
        options_frame = ttk.LabelFrame(main_frame, text="选项", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 选项复选框
        options_grid = ttk.Frame(options_frame)
        options_grid.pack(fill=tk.X)
        
        self.case_sensitive = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_grid, text="区分大小写", 
                       variable=self.case_sensitive).grid(row=0, column=0, sticky='w', padx=(0, 20))
        
        self.regex_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_grid, text="正则表达式", 
                       variable=self.regex_mode).grid(row=0, column=1, sticky='w')
        
        self.whole_word = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_grid, text="全字匹配", 
                       variable=self.whole_word).grid(row=1, column=0, sticky='w', padx=(0, 20), pady=(5, 0))
        
        # 搜索范围
        scope_frame = ttk.LabelFrame(main_frame, text="搜索范围", padding="10")
        scope_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.search_scope = tk.StringVar(value="all")
        
        scope_radio_frame = ttk.Frame(scope_frame)
        scope_radio_frame.pack(fill=tk.X)
        
        ttk.Radiobutton(scope_radio_frame, text="整个表格", 
                       variable=self.search_scope, value="all").pack(side=tk.LEFT)
        ttk.Radiobutton(scope_radio_frame, text="选定列", 
                       variable=self.search_scope, value="column").pack(side=tk.LEFT, padx=(20, 0))
        
        # 列选择框
        column_frame = ttk.Frame(scope_frame)
        column_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(column_frame, text="选择列:").pack(side=tk.LEFT)
        
        self.selected_column = tk.StringVar()
        column_combo = ttk.Combobox(column_frame, textvariable=self.selected_column,
                                   values=self.table_manager.get_column_names(),
                                   state="readonly", width=20)
        column_combo.pack(side=tk.LEFT, padx=(10, 0))
        
        # 结果显示
        result_frame = ttk.LabelFrame(main_frame, text="结果", padding="10")
        result_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.result_label = ttk.Label(result_frame, text="", foreground="blue")
        self.result_label.pack()
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 左侧按钮
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=tk.LEFT)
        
        ttk.Button(left_buttons, text="🔍 查找", 
                  command=self.find_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(left_buttons, text="🔄 替换全部", 
                  command=self.replace_all).pack(side=tk.LEFT, padx=5)
        
        # 右侧按钮
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side=tk.RIGHT)
        
        ttk.Button(right_buttons, text="关闭", 
                  command=self.on_close).pack(side=tk.RIGHT, padx=(5, 0))
        
    def find_all(self):
        """查找所有匹配项"""
        find_text = self.find_var.get().strip()
        if not find_text:
            messagebox.showwarning("警告", "请输入要查找的文本")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None or df.empty:
            messagebox.showwarning("警告", "没有数据可以搜索")
            return
            
        try:
            matches = self._find_matches(df, find_text)
            
            if matches:
                self.result_label.config(
                    text=f"找到 {len(matches)} 个匹配项",
                    foreground="green"
                )
                
                # 显示匹配详情
                detail_msg = f"找到 {len(matches)} 个匹配项：\n\n"
                for i, (row_idx, col_name, old_val, _) in enumerate(matches[:10]):  # 只显示前10个
                    detail_msg += f"{i+1}. {col_name}[第{row_idx+1}行]: {str(old_val)[:50]}...\n"
                    
                if len(matches) > 10:
                    detail_msg += f"\n... 还有 {len(matches)-10} 个匹配项"
                    
                messagebox.showinfo("查找结果", detail_msg)
            else:
                self.result_label.config(
                    text="未找到匹配项",
                    foreground="orange"
                )
                messagebox.showinfo("查找结果", "未找到匹配项")
                
        except Exception as e:
            messagebox.showerror("错误", f"查找失败: {str(e)}")
            self.result_label.config(
                text="查找失败",
                foreground="red"
            )
            
    def replace_all(self):
        """替换所有匹配项"""
        find_text = self.find_var.get().strip()
        replace_text = self.replace_var.get()
        
        if not find_text:
            messagebox.showwarning("警告", "请输入要查找的文本")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None or df.empty:
            messagebox.showwarning("警告", "没有数据可以替换")
            return
            
        try:
            matches = self._find_matches(df, find_text)
            
            if not matches:
                messagebox.showinfo("替换结果", "未找到匹配项，无需替换")
                return
                
            # 确认替换
            if not messagebox.askyesno("确认替换", 
                                     f"找到 {len(matches)} 个匹配项，确认全部替换吗？"):
                return
                
            # 执行替换
            replace_count = self._perform_replacements(df, matches, find_text, replace_text)
            
            # 更新表格显示（TableManager没有set_dataframe方法，数据已经在原地修改）
            # self.table_manager.dataframe = df  # 不需要，因为df是引用
            
            # 显示结果
            self.result_label.config(
                text=f"已替换 {replace_count} 个匹配项",
                foreground="green"
            )
            
            messagebox.showinfo("替换完成", f"成功替换了 {replace_count} 个匹配项")
            
        except Exception as e:
            messagebox.showerror("错误", f"替换失败: {str(e)}")
            self.result_label.config(
                text="替换失败",
                foreground="red"
            )
            
    def _find_matches(self, df, find_text):
        """查找匹配项"""
        matches = []
        
        # 确定搜索范围
        if self.search_scope.get() == "column":
            col_name = self.selected_column.get()
            if not col_name:
                raise ValueError("请选择要搜索的列")
            if col_name not in df.columns:
                raise ValueError(f"列 '{col_name}' 不存在")
            search_columns = [col_name]
        else:
            search_columns = df.columns
            
        # 遍历指定列
        for col_name in search_columns:
            for row_idx in df.index:
                cell_value = df.at[row_idx, col_name]
                if pd.isna(cell_value):
                    continue
                    
                cell_str = str(cell_value)
                
                # 检查是否匹配
                if self._is_match(cell_str, find_text):
                    matches.append((row_idx, col_name, cell_value, cell_str))
                    
        return matches
        
    def _is_match(self, text, pattern):
        """检查文本是否匹配模式"""
        import re  # 在函数开头导入
        
        if self.regex_mode.get():
            # 正则表达式模式
            flags = 0 if self.case_sensitive.get() else re.IGNORECASE
            try:
                return bool(re.search(pattern, text, flags))
            except re.error:
                raise ValueError("正则表达式语法错误")
        else:
            # 普通文本模式
            if not self.case_sensitive.get():
                text = text.lower()
                pattern = pattern.lower()
                
            if self.whole_word.get():
                # 全字匹配
                word_pattern = r'\b' + re.escape(pattern) + r'\b'
                flags = 0 if self.case_sensitive.get() else re.IGNORECASE
                return bool(re.search(word_pattern, text, flags))
            else:
                # 部分匹配
                return pattern in text
                
    def _perform_replacements(self, df, matches, find_text, replace_text):
        """执行替换操作"""
        import re  # 在函数开头导入
        replace_count = 0
        
        for row_idx, col_name, original_value, cell_str in matches:
            if self.regex_mode.get():
                # 正则表达式替换
                flags = 0 if self.case_sensitive.get() else re.IGNORECASE
                new_value = re.sub(find_text, replace_text, cell_str, flags=flags)
            else:
                # 普通文本替换
                if self.whole_word.get():
                    # 全字替换
                    word_pattern = r'\b' + re.escape(find_text) + r'\b'
                    flags = 0 if self.case_sensitive.get() else re.IGNORECASE
                    new_value = re.sub(word_pattern, replace_text, cell_str, flags=flags)
                else:
                    # 部分替换
                    if self.case_sensitive.get():
                        new_value = cell_str.replace(find_text, replace_text)
                    else:
                        # 大小写不敏感替换
                        pattern = re.escape(find_text)
                        new_value = re.sub(pattern, replace_text, cell_str, flags=re.IGNORECASE)
            
            # 尝试保持原始数据类型
            if new_value != cell_str:
                try:
                    # 如果原来是数字，尝试转换回数字
                    if isinstance(original_value, (int, float)):
                        try:
                            new_value = type(original_value)(new_value)
                        except (ValueError, TypeError):
                            pass  # 转换失败，保持字符串
                except:
                    pass
                    
                df.at[row_idx, col_name] = new_value
                replace_count += 1
                
        return replace_count
        
    def on_close(self):
        """关闭对话框"""
        self.dialog.destroy()
        
    def show(self):
        """显示对话框"""
        self.dialog.wait_window() 