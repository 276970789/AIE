#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JSONL文件导入对话框
支持通过匹配字段从外部JSONL文件导入数据
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os

class JsonlImportDialog:
    def __init__(self, parent, table_manager, callback):
        self.parent = parent
        self.table_manager = table_manager
        self.callback = callback
        self.result = None
        
        # JSONL数据
        self.jsonl_data = None
        self.jsonl_df = None
        self.selected_file = None
        
        # 创建对话框
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("从JSONL文件导入数据")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (400)
        y = (self.dialog.winfo_screenheight() // 2) - (300)
        self.dialog.geometry(f"800x600+{x}+{y}")
        
        self.create_widgets()
        
    def create_widgets(self):
        """创建界面组件"""
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="从JSONL文件导入数据", 
                               style='Title.TLabel')
        title_label.pack(pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="1. 选择JSONL文件", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        file_select_frame = ttk.Frame(file_frame)
        file_select_frame.pack(fill=tk.X)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_select_frame, textvariable=self.file_path_var, 
                                   state='readonly', width=50)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        ttk.Button(file_select_frame, text="📁 浏览", 
                  command=self.select_file).pack(side=tk.RIGHT)
        
        # 匹配字段区域
        self.match_frame = ttk.LabelFrame(main_frame, text="2. 选择匹配字段", padding="10")
        self.match_frame.pack(fill=tk.X, pady=(0, 15))
        
        match_info = ttk.Label(self.match_frame, 
                              text="选择用于匹配的字段（两个文件中都必须存在）：",
                              style='Subtitle.TLabel')
        match_info.pack(anchor=tk.W, pady=(0, 10))
        
        match_select_frame = ttk.Frame(self.match_frame)
        match_select_frame.pack(fill=tk.X)
        
        ttk.Label(match_select_frame, text="匹配字段:").pack(side=tk.LEFT)
        self.match_field_var = tk.StringVar()
        self.match_field_combo = ttk.Combobox(match_select_frame, textvariable=self.match_field_var,
                                             state='readonly', width=30)
        self.match_field_combo.pack(side=tk.LEFT, padx=(10, 0))
        
        # 导入字段区域
        self.import_frame = ttk.LabelFrame(main_frame, text="3. 选择要导入的字段", padding="10")
        self.import_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        import_info = ttk.Label(self.import_frame, 
                               text="选择要从JSONL文件导入的字段：",
                               style='Subtitle.TLabel')
        import_info.pack(anchor=tk.W, pady=(0, 10))
        
        # 字段选择和新列名设置
        fields_frame = ttk.Frame(self.import_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：可用字段列表
        left_frame = ttk.Frame(fields_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        ttk.Label(left_frame, text="JSONL文件中的字段:").pack(anchor=tk.W)
        
        # 字段列表框架
        fields_list_frame = ttk.Frame(left_frame)
        fields_list_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # 创建Treeview来显示字段
        self.fields_tree = ttk.Treeview(fields_list_frame, columns=('field', 'sample'), 
                                       show='headings', height=8)
        self.fields_tree.heading('field', text='字段名')
        self.fields_tree.heading('sample', text='示例值')
        self.fields_tree.column('field', width=150)
        self.fields_tree.column('sample', width=300)
        
        # 滚动条
        fields_scrollbar = ttk.Scrollbar(fields_list_frame, orient=tk.VERTICAL, 
                                        command=self.fields_tree.yview)
        self.fields_tree.configure(yscrollcommand=fields_scrollbar.set)
        
        self.fields_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fields_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 右侧：导入设置
        right_frame = ttk.Frame(fields_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        ttk.Label(right_frame, text="导入设置:").pack(anchor=tk.W)
        
        # 选中字段显示
        selected_frame = ttk.LabelFrame(right_frame, text="选中字段", padding="5")
        selected_frame.pack(fill=tk.X, pady=(5, 10))
        
        self.selected_field_var = tk.StringVar()
        self.selected_field_label = ttk.Label(selected_frame, textvariable=self.selected_field_var)
        self.selected_field_label.pack()
        
        # 新列名输入
        column_name_frame = ttk.LabelFrame(right_frame, text="新列名", padding="5")
        column_name_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.column_name_var = tk.StringVar()
        self.column_name_entry = ttk.Entry(column_name_frame, textvariable=self.column_name_var)
        self.column_name_entry.pack(fill=tk.X)
        
        # 按钮
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="🔍 预览", 
                  command=self.preview_import).pack(fill=tk.X, pady=(0, 5))
        
        # 预览区域
        self.preview_frame = ttk.LabelFrame(main_frame, text="4. 导入预览", padding="10")
        self.preview_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.preview_text = tk.Text(self.preview_frame, height=6, wrap=tk.WORD)
        preview_scrollbar = ttk.Scrollbar(self.preview_frame, orient=tk.VERTICAL, 
                                         command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X)
        
        ttk.Button(bottom_frame, text="✅ 导入", 
                  command=self.do_import).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(bottom_frame, text="❌ 取消", 
                  command=self.cancel).pack(side=tk.RIGHT)
        
        # 绑定事件
        self.fields_tree.bind('<<TreeviewSelect>>', self.on_field_select)
        
        # 禁用控件
        self.set_controls_state(False)
        
    def select_file(self):
        """选择JSONL文件"""
        file_path = filedialog.askopenfilename(
            title="选择JSONL文件",
            filetypes=[("JSONL文件", "*.jsonl"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.file_path_var.set(file_path)
            self.selected_file = file_path
            self.load_jsonl_file(file_path)
            
    def load_jsonl_file(self, file_path):
        """加载JSONL文件"""
        try:
            # 读取JSONL文件
            jsonl_data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            data = json.loads(line)
                            jsonl_data.append(data)
                        except json.JSONDecodeError as e:
                            print(f"跳过无效行: {line[:50]}... 错误: {e}")
                            
            if not jsonl_data:
                messagebox.showerror("错误", "JSONL文件为空或格式无效")
                return
                
            self.jsonl_data = jsonl_data
            self.jsonl_df = pd.DataFrame(jsonl_data)
            
            # 更新界面
            self.update_match_fields()
            self.update_fields_list()
            self.set_controls_state(True)
            
            messagebox.showinfo("成功", f"已加载 {len(jsonl_data)} 条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载JSONL文件失败: {e}")
            
    def update_match_fields(self):
        """更新匹配字段下拉框"""
        if self.jsonl_df is None:
            return
            
        # 获取当前表格的列名
        current_df = self.table_manager.get_dataframe()
        if current_df is None:
            return
            
        current_columns = set(current_df.columns)
        jsonl_columns = set(self.jsonl_df.columns)
        
        # 找到共同字段
        common_fields = current_columns.intersection(jsonl_columns)
        
        if not common_fields:
            messagebox.showwarning("警告", "当前表格和JSONL文件没有共同字段")
            return
            
        self.match_field_combo['values'] = list(common_fields)
        if common_fields:
            self.match_field_var.set(list(common_fields)[0])
            
    def update_fields_list(self):
        """更新字段列表"""
        if self.jsonl_df is None:
            return
            
        # 清空现有项目
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
            
        # 添加字段
        for column in self.jsonl_df.columns:
            # 获取示例值
            sample_values = self.jsonl_df[column].dropna().head(3).tolist()
            sample_text = ", ".join([str(v)[:50] for v in sample_values])
            if len(sample_text) > 100:
                sample_text = sample_text[:100] + "..."
                
            self.fields_tree.insert('', 'end', values=(column, sample_text))
            
    def on_field_select(self, event):
        """字段选择事件"""
        selection = self.fields_tree.selection()
        if selection:
            item = self.fields_tree.item(selection[0])
            field_name = item['values'][0]
            self.selected_field_var.set(field_name)
            
            # 自动设置列名
            if not self.column_name_var.get():
                self.column_name_var.set(field_name)
                
    def preview_import(self):
        """预览导入"""
        if not self.validate_settings():
            return
            
        try:
            match_field = self.match_field_var.get()
            selected_field = self.selected_field_var.get()
            
            # 获取当前表格
            current_df = self.table_manager.get_dataframe()
            
            # 执行匹配
            preview_data = []
            matched_count = 0
            
            for idx, row in current_df.iterrows():
                match_value = row[match_field]
                
                # 在JSONL数据中查找匹配项
                jsonl_match = self.jsonl_df[self.jsonl_df[match_field] == match_value]
                
                if not jsonl_match.empty:
                    import_value = jsonl_match.iloc[0][selected_field]
                    matched_count += 1
                    preview_data.append(f"行{idx+1}: {match_field}='{match_value}' -> {selected_field}='{import_value}'")
                else:
                    preview_data.append(f"行{idx+1}: {match_field}='{match_value}' -> 未找到匹配")
                    
            # 显示预览
            preview_text = f"匹配结果预览 (匹配了 {matched_count}/{len(current_df)} 行):\n\n"
            preview_text += "\n".join(preview_data[:20])  # 只显示前20行
            
            if len(preview_data) > 20:
                preview_text += f"\n\n... 还有 {len(preview_data) - 20} 行（总共 {len(preview_data)} 行）"
                
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, preview_text)
            
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {e}")
            
    def validate_settings(self):
        """验证设置"""
        if not self.selected_file:
            messagebox.showwarning("警告", "请选择JSONL文件")
            return False
            
        if not self.match_field_var.get():
            messagebox.showwarning("警告", "请选择匹配字段")
            return False
            
        if not self.selected_field_var.get():
            messagebox.showwarning("警告", "请选择要导入的字段")
            return False
            
        column_name = self.column_name_var.get().strip()
        if not column_name:
            messagebox.showwarning("警告", "请输入新列名")
            return False
            
        # 检查列名是否已存在
        existing_columns = self.table_manager.get_column_names()
        if column_name in existing_columns:
            messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
            return False
            
        return True
        
    def do_import(self):
        """执行导入"""
        if not self.validate_settings():
            return
            
        try:
            match_field = self.match_field_var.get()
            selected_field = self.selected_field_var.get()
            column_name = self.column_name_var.get().strip()
            
            # 准备导入数据
            import_data = {
                'match_field': match_field,
                'source_field': selected_field,
                'column_name': column_name,
                'jsonl_data': self.jsonl_df
            }
            
            self.result = import_data
            
            # 关闭对话框
            self.dialog.destroy()
            
            # 调用回调
            if self.callback:
                self.callback(self.result)
                
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")
            
    def cancel(self):
        """取消"""
        self.result = None
        self.dialog.destroy()
        
    def set_controls_state(self, enabled):
        """设置控件状态"""
        state = 'normal' if enabled else 'disabled'
        self.match_field_combo.config(state=state)
        self.column_name_entry.config(state=state)
        
    def show(self):
        """显示对话框"""
        self.dialog.wait_window()
        return self.result 