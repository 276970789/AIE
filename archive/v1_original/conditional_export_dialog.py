#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
条件筛选导出对话框
允许用户根据多种条件筛选数据后导出
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json

class ConditionalExportDialog:
    def __init__(self, parent, table_manager):
        self.parent = parent
        self.table_manager = table_manager
        self.conditions = []  # 存储筛选条件
        self.filtered_df = None  # 筛选后的数据框
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("条件筛选导出")
        self.dialog.geometry("900x700")
        self.dialog.resizable(True, True)
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # self.center_window() # 移除固定居中，让用户可以自由调整窗口大小
        self.create_widgets()
        
    # def center_window(self):
    #     # 移除固定居中，让用户可以自由调整窗口大小
    #     self.dialog.update_idletasks()
    #     x = (self.dialog.winfo_screenwidth() // 2) - (900 // 2)
    #     y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
    #     self.dialog.geometry(f"900x700+{x}+{y}")
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="条件筛选导出", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # 条件设置区域
        conditions_frame = ttk.LabelFrame(main_frame, text="筛选条件", padding="10")
        conditions_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 添加条件按钮
        add_condition_btn = ttk.Button(conditions_frame, text="➕ 添加条件", 
                                      command=self.add_condition)
        add_condition_btn.pack(anchor=tk.W, pady=(0, 10))
        
        # 条件列表区域
        self.conditions_list_frame = ttk.Frame(conditions_frame)
        self.conditions_list_frame.pack(fill=tk.X)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="筛选预览", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 预览按钮和统计信息
        preview_header = ttk.Frame(preview_frame)
        preview_header.pack(fill=tk.X, pady=(0, 10))
        
        self.preview_btn = ttk.Button(preview_header, text="🔍 预览筛选结果", 
                                     command=self.preview_filtered_data)
        self.preview_btn.pack(side=tk.LEFT)
        
        self.stats_label = ttk.Label(preview_header, text="", foreground="blue")
        self.stats_label.pack(side=tk.LEFT, padx=(15, 0))
        
        # 预览表格
        preview_table_frame = ttk.Frame(preview_frame)
        preview_table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建预览表格
        self.preview_tree = ttk.Treeview(preview_table_frame, show='headings', height=12)
        
        # 添加滚动条
        preview_v_scrollbar = ttk.Scrollbar(preview_table_frame, orient=tk.VERTICAL, 
                                           command=self.preview_tree.yview)
        preview_h_scrollbar = ttk.Scrollbar(preview_table_frame, orient=tk.HORIZONTAL, 
                                           command=self.preview_tree.xview)
        
        self.preview_tree.configure(yscrollcommand=preview_v_scrollbar.set,
                                   xscrollcommand=preview_h_scrollbar.set)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        preview_h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 导出设置区域
        export_frame = ttk.LabelFrame(main_frame, text="导出设置", padding="10")
        export_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 导出格式选择
        format_frame = ttk.Frame(export_frame)
        format_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(format_frame, text="导出格式:").pack(side=tk.LEFT)
        
        self.export_format = tk.StringVar(value="excel")
        ttk.Radiobutton(format_frame, text="Excel (.xlsx)", variable=self.export_format, 
                       value="excel").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(format_frame, text="CSV (.csv)", variable=self.export_format, 
                       value="csv").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(format_frame, text="JSONL (.jsonl)", variable=self.export_format, 
                       value="jsonl").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(format_frame, text="AI Excel (.aie)", variable=self.export_format, 
                       value="aie").pack(side=tk.LEFT, padx=(10, 0))
        
        # 字段选择
        fields_frame = ttk.Frame(export_frame)
        fields_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(fields_frame, text="导出字段:").pack(side=tk.LEFT)
        
        self.export_all_columns = tk.BooleanVar(value=True)
        ttk.Checkbutton(fields_frame, text="导出所有字段", 
                       variable=self.export_all_columns,
                       command=self.toggle_column_selection).pack(side=tk.LEFT, padx=(10, 0))
        
        self.select_columns_btn = ttk.Button(fields_frame, text="选择字段...", 
                                           command=self.select_export_columns,
                                           state="disabled")
        self.select_columns_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        self.selected_columns = []  # 存储选中的列
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(button_frame, text="取消", command=self.on_cancel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="导出", command=self.on_export).pack(side=tk.RIGHT, padx=(5, 5))
        
        # 初始化：添加一个默认条件
        self.add_condition()
        
    def add_condition(self):
        """添加筛选条件"""
        condition_frame = ttk.Frame(self.conditions_list_frame)
        condition_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 逻辑连接符（AND/OR）
        logic_var = tk.StringVar(value="AND")
        if len(self.conditions) == 0:
            # 第一个条件不显示逻辑连接符
            logic_label = ttk.Label(condition_frame, text="     ", width=6)
        else:
            logic_combo = ttk.Combobox(condition_frame, textvariable=logic_var,
                                      values=["AND", "OR"], state="readonly", width=5)
            logic_combo.pack(side=tk.LEFT, padx=(0, 5))
            logic_label = logic_combo
        
        logic_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 字段选择
        field_var = tk.StringVar()
        field_combo = ttk.Combobox(condition_frame, textvariable=field_var,
                                  values=self.table_manager.get_column_names(),
                                  state="readonly", width=15)
        field_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        # 条件类型选择
        condition_var = tk.StringVar(value="非空")
        condition_combo = ttk.Combobox(condition_frame, textvariable=condition_var,
                                      values=["非空", "为空", "等于", "不等于", "包含", "不包含", 
                                             "开始于", "结束于", "大于", "小于", "大于等于", "小于等于"],
                                      state="readonly", width=10)
        condition_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        # 值输入框
        value_var = tk.StringVar()
        value_entry = ttk.Entry(condition_frame, textvariable=value_var, width=15)
        value_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 删除按钮
        delete_btn = ttk.Button(condition_frame, text="❌", width=3,
                               command=lambda: self.remove_condition(condition_frame))
        delete_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 存储条件信息
        condition_info = {
            'frame': condition_frame,
            'logic_var': logic_var,
            'field_var': field_var,
            'condition_var': condition_var,
            'value_var': value_var
        }
        
        self.conditions.append(condition_info)
        
        # 绑定条件类型变化事件
        condition_combo.bind('<<ComboboxSelected>>', 
                           lambda e: self.on_condition_type_change(condition_info))
        
        # 初始状态设置
        self.on_condition_type_change(condition_info)
        
    def on_condition_type_change(self, condition_info):
        """条件类型改变时的处理"""
        condition_type = condition_info['condition_var'].get()
        value_entry = None
        
        # 找到对应的值输入框
        for widget in condition_info['frame'].winfo_children():
            if isinstance(widget, ttk.Entry):
                value_entry = widget
                break
                
        if value_entry:
            if condition_type in ["非空", "为空"]:
                value_entry.config(state="disabled")
                condition_info['value_var'].set("")
            else:
                value_entry.config(state="normal")
        
    def remove_condition(self, condition_frame):
        """删除筛选条件"""
        # 找到并移除对应的条件信息
        for condition in self.conditions[:]:
            if condition['frame'] == condition_frame:
                self.conditions.remove(condition)
                break
        
        # 销毁框架
        condition_frame.destroy()
        
        # 如果删除的是第一个条件，需要更新后续条件的逻辑连接符显示
        if self.conditions:
            first_condition = self.conditions[0]
            first_frame = first_condition['frame']
            for widget in first_frame.winfo_children():
                if isinstance(widget, ttk.Combobox) or isinstance(widget, ttk.Label):
                    if widget.cget('width') <= 6:  # 可能是逻辑连接符
                        widget.destroy()
                        # 添加空白标签
                        empty_label = ttk.Label(first_frame, text="     ", width=6)
                        empty_label.pack(side=tk.LEFT, padx=(0, 5))
                        break
    
    def preview_filtered_data(self):
        """预览筛选后的数据"""
        if not self.conditions:
            messagebox.showwarning("警告", "请至少添加一个筛选条件")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None or df.empty:
            messagebox.showwarning("警告", "没有数据可以筛选")
            return
            
        try:
            # 构建筛选条件
            filter_mask = None
            
            for i, condition in enumerate(self.conditions):
                field = condition['field_var'].get()
                condition_type = condition['condition_var'].get()
                value = condition['value_var'].get().strip()
                logic = condition['logic_var'].get()
                
                if not field:
                    messagebox.showwarning("警告", f"第{i+1}个条件的字段不能为空")
                    return
                    
                if field not in df.columns:
                    messagebox.showwarning("警告", f"字段 '{field}' 不存在")
                    return
                
                # 创建当前条件的掩码
                current_mask = self.create_condition_mask(df, field, condition_type, value)
                
                if filter_mask is None:
                    filter_mask = current_mask
                else:
                    if logic == "AND":
                        filter_mask = filter_mask & current_mask
                    else:  # OR
                        filter_mask = filter_mask | current_mask
            
            # 应用筛选
            self.filtered_df = df[filter_mask].copy()
            
            # 更新统计信息
            total_rows = len(df)
            filtered_rows = len(self.filtered_df)
            self.stats_label.config(
                text=f"总行数: {total_rows}, 筛选后: {filtered_rows} 行 ({filtered_rows/total_rows*100:.1f}%)"
            )
            
            # 更新预览表格
            self.update_preview_table()
            
        except Exception as e:
            messagebox.showerror("错误", f"筛选数据时出错: {str(e)}")
    
    def create_condition_mask(self, df, field, condition_type, value):
        """创建条件掩码"""
        series = df[field].astype(str)  # 统一转换为字符串处理
        
        if condition_type == "非空":
            return (df[field].notna()) & (df[field].astype(str).str.strip() != "")
        elif condition_type == "为空":
            return (df[field].isna()) | (df[field].astype(str).str.strip() == "")
        elif condition_type == "等于":
            return series == value
        elif condition_type == "不等于":
            return series != value
        elif condition_type == "包含":
            return series.str.contains(value, case=False, na=False)
        elif condition_type == "不包含":
            return ~series.str.contains(value, case=False, na=False)
        elif condition_type == "开始于":
            return series.str.startswith(value, na=False)
        elif condition_type == "结束于":
            return series.str.endswith(value, na=False)
        elif condition_type in ["大于", "小于", "大于等于", "小于等于"]:
            try:
                # 尝试数值比较
                numeric_series = pd.to_numeric(df[field], errors='coerce')
                numeric_value = float(value)
                
                if condition_type == "大于":
                    return numeric_series > numeric_value
                elif condition_type == "小于":
                    return numeric_series < numeric_value
                elif condition_type == "大于等于":
                    return numeric_series >= numeric_value
                elif condition_type == "小于等于":
                    return numeric_series <= numeric_value
            except (ValueError, TypeError):
                # 如果无法转换为数值，使用字符串比较
                if condition_type == "大于":
                    return series > value
                elif condition_type == "小于":
                    return series < value
                elif condition_type == "大于等于":
                    return series >= value
                elif condition_type == "小于等于":
                    return series <= value
        
        # 默认返回全部为False的掩码
        return pd.Series([False] * len(df), index=df.index)
    
    def update_preview_table(self):
        """更新预览表格"""
        # 清空现有数据
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
            
        if self.filtered_df is None or self.filtered_df.empty:
            return
            
        # 设置列
        columns = list(self.filtered_df.columns)
        self.preview_tree["columns"] = columns
        
        for col in columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=120, anchor='w')
        
        # 插入数据（最多显示100行）
        display_rows = min(100, len(self.filtered_df))
        for i in range(display_rows):
            row = self.filtered_df.iloc[i]
            values = []
            for val in row:
                str_val = str(val) if val is not None else ""
                if len(str_val) > 50:
                    str_val = str_val[:47] + "..."
                values.append(str_val)
            self.preview_tree.insert("", "end", values=values)
            
        if len(self.filtered_df) > 100:
            # 添加提示信息
            self.preview_tree.insert("", "end", values=["..."] + ["仅显示前100行"] + ["..."] * (len(columns)-2))
    
    def toggle_column_selection(self):
        """切换列选择模式"""
        if self.export_all_columns.get():
            self.select_columns_btn.config(state="disabled")
            self.selected_columns = []
        else:
            self.select_columns_btn.config(state="normal")
    
    def select_export_columns(self):
        """选择要导出的列"""
        if self.filtered_df is None:
            messagebox.showwarning("警告", "请先预览筛选结果")
            return
            
        # 创建列选择对话框
        column_dialog = tk.Toplevel(self.dialog)
        column_dialog.title("选择导出字段")
        column_dialog.geometry("400x650")
        column_dialog.resizable(True, True)
        column_dialog.transient(self.dialog)
        column_dialog.grab_set()
        
        # 移除固定居中，让用户可以自由调整窗口大小
        # column_dialog.update_idletasks()
        # x = (column_dialog.winfo_screenwidth() // 2) - (400 // 2)
        # y = (column_dialog.winfo_screenheight() // 2) - (500 // 2)
        # column_dialog.geometry(f"400x500+{x}+{y}")
        
        main_frame = ttk.Frame(column_dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="选择要导出的字段:", font=('Arial', 12, 'bold')).pack(pady=(0, 10))
        
        # 全选/取消全选
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(select_frame, text="全选", 
                  command=lambda: self.select_all_columns(checkboxes, True)).pack(side=tk.LEFT)
        ttk.Button(select_frame, text="取消全选", 
                  command=lambda: self.select_all_columns(checkboxes, False)).pack(side=tk.LEFT, padx=(5, 0))
        
        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(15, 0))
        
        # 列选择区域
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        checkboxes = {}
        for col in self.filtered_df.columns:
            var = tk.BooleanVar(value=col in self.selected_columns or not self.selected_columns)
            checkboxes[col] = var
            ttk.Checkbutton(scrollable_frame, text=col, variable=var).pack(anchor=tk.W, pady=1)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def on_column_ok():
            self.selected_columns = [col for col, var in checkboxes.items() if var.get()]
            column_dialog.destroy()
        
        def on_column_cancel():
            column_dialog.destroy()
        
        ttk.Button(button_frame, text="确定", command=on_column_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=on_column_cancel).pack(side=tk.RIGHT)
    
    def select_all_columns(self, checkboxes, state):
        """全选或取消全选列"""
        for var in checkboxes.values():
            var.set(state)
    
    def on_export(self):
        """执行导出"""
        if self.filtered_df is None or self.filtered_df.empty:
            messagebox.showwarning("警告", "没有筛选结果可以导出，请先预览筛选结果")
            return
        
        # 确定要导出的列
        if self.export_all_columns.get():
            export_df = self.filtered_df.copy()
        else:
            if not self.selected_columns:
                messagebox.showwarning("警告", "请选择要导出的字段")
                return
            export_df = self.filtered_df[self.selected_columns].copy()
        
        # 选择保存路径
        format_type = self.export_format.get()
        if format_type == "excel":
            file_path = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
        elif format_type == "csv":
            file_path = filedialog.asksaveasfilename(
                title="保存CSV文件",
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
            )
        elif format_type == "jsonl":
            file_path = filedialog.asksaveasfilename(
                title="保存JSONL文件",
                defaultextension=".jsonl",
                filetypes=[("JSONL文件", "*.jsonl"), ("所有文件", "*.*")]
            )
        elif format_type == "aie": # 新增AI Excel格式
            file_path = filedialog.asksaveasfilename(
                title="保存AI Excel文件",
                defaultextension=".aie",
                filetypes=[("AI Excel文件", "*.aie"), ("所有文件", "*.*")]
            )
        
        if not file_path:
            return
        
        try:
            # 执行导出
            if format_type == "excel":
                export_df.to_excel(file_path, index=False)
            elif format_type == "csv":
                export_df.to_csv(file_path, index=False, encoding='utf-8-sig')
            elif format_type == "jsonl":
                with open(file_path, 'w', encoding='utf-8') as f:
                    for _, row in export_df.iterrows():
                        row_dict = row.to_dict()
                        json_line = json.dumps(row_dict, ensure_ascii=False)
                        f.write(json_line + '\n')
            elif format_type == "aie": # 保存为aie文件
                # 获取AI列和长文本列配置
                ai_columns = self.table_manager.get_ai_columns()
                long_text_columns = self.table_manager.get_long_text_columns()

                # 构建完整的项目数据，包含表格数据和列配置
                project_data = {
                    "format_version": "1.0", # 与ProjectManager中的版本保持一致
                    "table_data": {
                        "columns": list(export_df.columns),
                        "data": export_df.to_dict('records'),
                        "row_count": len(export_df),
                        "col_count": len(export_df.columns)
                    },
                    "ai_config": {
                        "ai_columns": ai_columns,
                        "ai_column_count": len(ai_columns),
                        "prompt_templates": {col_name: config for col_name, config in ai_columns.items()} # 假设ai_columns已经包含完整配置
                    },
                    "long_text_config": {
                        "long_text_columns": long_text_columns,
                        "long_text_column_count": len(long_text_columns)
                    },
                    "normal_columns": [col for col in export_df.columns if col not in ai_columns and col not in long_text_columns]
                }
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(project_data, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("成功", f"已成功导出 {len(export_df)} 行数据到:\n{file_path}")
            # self.dialog.destroy() # 不再自动关闭对话框
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def on_cancel(self):
        """取消导出"""
        self.dialog.destroy()
    
    def show(self):
        """显示对话框"""
        self.dialog.wait_window() 