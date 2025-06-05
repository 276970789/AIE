#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
随机导出对话框
支持完全随机导出和按字段分层导出两种模式
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

class RandomExportDialog:
    def __init__(self, parent, table_manager):
        self.parent = parent
        self.table_manager = table_manager
        self.result = None
        
        # 创建对话框窗口
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("🎲 随机导出数据")
        self.dialog.geometry("650x550")
        self.dialog.resizable(True, True)
        self.dialog.minsize(600, 500)
        
        # 设置模态
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.center_window()
        
        # 创建界面
        self.create_widgets()
        
        # 绑定关闭事件
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # 绑定键盘快捷键
        self.dialog.bind('<Return>', lambda e: self.export_data())
        self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        self.dialog.bind('<Control-p>', lambda e: self.preview_export())
        
        # 设置焦点
        self.dialog.focus_set()
        
    def center_window(self):
        """居中显示窗口"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (650 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (550 // 2)
        self.dialog.geometry(f"650x550+{x}+{y}")
        
    def create_widgets(self):
        """创建界面组件"""
        # 主容器
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="随机导出数据", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # 数据信息
        df = self.table_manager.get_dataframe()
        info_text = f"当前数据：{len(df)}行 × {len(df.columns)}列"
        info_label = ttk.Label(main_frame, text=info_text, 
                              font=('Arial', 10))
        info_label.pack(pady=(0, 15))
        
        # 导出模式选择
        mode_frame = ttk.LabelFrame(main_frame, text="导出模式", padding=15)
        mode_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.export_mode = tk.StringVar(value="random")
        
        # 完全随机模式
        random_radio = ttk.Radiobutton(mode_frame, text="完全随机导出", 
                                      variable=self.export_mode, value="random",
                                      command=self.on_mode_change)
        random_radio.pack(anchor=tk.W, pady=2)
        
        random_desc = ttk.Label(mode_frame, text="从所有数据中随机选择指定数量的行", 
                               font=('Arial', 9), foreground='gray')
        random_desc.pack(anchor=tk.W, padx=(20, 0), pady=(0, 10))
        
        # 分层随机模式
        stratified_radio = ttk.Radiobutton(mode_frame, text="按字段分层导出 (推荐)", 
                                          variable=self.export_mode, value="stratified",
                                          command=self.on_mode_change)
        stratified_radio.pack(anchor=tk.W, pady=2)
        
        stratified_desc = ttk.Label(mode_frame, text="按某个字段的值分层，每个值随机选择指定数量（适合prompt调试）", 
                                   font=('Arial', 9), foreground='gray')
        stratified_desc.pack(anchor=tk.W, padx=(20, 0))
        
        # 参数设置区域
        self.params_frame = ttk.LabelFrame(main_frame, text="参数设置", padding=15)
        self.params_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 完全随机模式的参数
        self.random_params_frame = ttk.Frame(self.params_frame)
        
        ttk.Label(self.random_params_frame, text="导出数量:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.random_count_var = tk.StringVar(value="10")
        random_count_entry = ttk.Entry(self.random_params_frame, textvariable=self.random_count_var, width=10)
        random_count_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # 快捷按钮
        quick_frame = ttk.Frame(self.random_params_frame)
        quick_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        ttk.Label(quick_frame, text="快捷选择:", font=('Arial', 9)).pack(side=tk.LEFT)
        for count in [5, 10, 20, 50]:
            btn = ttk.Button(quick_frame, text=str(count), width=3,
                           command=lambda c=count: self.random_count_var.set(str(c)))
            btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 分层模式的参数
        self.stratified_params_frame = ttk.Frame(self.params_frame)
        
        ttk.Label(self.stratified_params_frame, text="分层字段:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.field_var = tk.StringVar()
        self.field_combo = ttk.Combobox(self.stratified_params_frame, textvariable=self.field_var, 
                                       width=20, state="readonly")
        self.field_combo.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        self.field_combo.bind('<<ComboboxSelected>>', self.on_field_change)
        
        ttk.Label(self.stratified_params_frame, text="每个值的数量:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.per_value_count_var = tk.StringVar(value="1")
        per_value_entry = ttk.Entry(self.stratified_params_frame, textvariable=self.per_value_count_var, width=10)
        per_value_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # 快捷按钮
        quick_frame2 = ttk.Frame(self.stratified_params_frame)
        quick_frame2.grid(row=1, column=2, sticky=tk.W, padx=(10, 0), pady=5)
        
        ttk.Label(quick_frame2, text="快捷:", font=('Arial', 9)).pack(side=tk.LEFT)
        for count in [1, 2, 3, 5]:
            btn = ttk.Button(quick_frame2, text=str(count), width=3,
                           command=lambda c=count: self.per_value_count_var.set(str(c)))
            btn.pack(side=tk.LEFT, padx=(2, 0))
        
        # 字段值预览
        self.values_frame = ttk.Frame(self.stratified_params_frame)
        self.values_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W+tk.E, pady=(10, 0))
        
        ttk.Label(self.values_frame, text="字段值预览:").pack(anchor=tk.W)
        
        # 创建滚动文本框显示字段值
        values_scroll_frame = ttk.Frame(self.values_frame)
        values_scroll_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        self.values_text = tk.Text(values_scroll_frame, height=6, width=50, 
                                  font=('Arial', 9), state='disabled')
        values_scrollbar = ttk.Scrollbar(values_scroll_frame, orient=tk.VERTICAL, 
                                        command=self.values_text.yview)
        self.values_text.configure(yscrollcommand=values_scrollbar.set)
        
        self.values_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        values_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始化字段列表
        self.update_field_list()
        
        # 预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="导出预览", padding=15)
        preview_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.preview_label = ttk.Label(preview_frame, text="请选择导出参数", 
                                      font=('Arial', 10), foreground='gray')
        self.preview_label.pack()
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 预览按钮
        preview_btn = ttk.Button(button_frame, text="🔍 预览 (Ctrl+P)", command=self.preview_export)
        preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 导出按钮
        export_btn = ttk.Button(button_frame, text="📤 导出 (Enter)", command=self.export_data)
        export_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 取消按钮
        cancel_btn = ttk.Button(button_frame, text="取消 (Esc)", command=self.on_cancel)
        cancel_btn.pack(side=tk.RIGHT)
        
        # 初始显示模式
        self.on_mode_change()
        
        # 绑定参数变化事件
        self.random_count_var.trace('w', lambda *args: self.update_preview())
        self.per_value_count_var.trace('w', lambda *args: self.update_preview())
        
    def update_field_list(self):
        """更新字段列表"""
        df = self.table_manager.get_dataframe()
        columns = list(df.columns)
        self.field_combo['values'] = columns
        if columns:
            self.field_combo.set(columns[0])
            self.on_field_change()
            
    def on_mode_change(self):
        """模式切换事件"""
        mode = self.export_mode.get()
        
        # 隐藏所有参数框架
        self.random_params_frame.pack_forget()
        self.stratified_params_frame.pack_forget()
        
        # 显示对应的参数框架
        if mode == "random":
            self.random_params_frame.pack(fill=tk.X)
        else:
            self.stratified_params_frame.pack(fill=tk.X)
            
        # 更新预览
        self.update_preview()
        
    def on_field_change(self, event=None):
        """字段选择变化事件"""
        field_name = self.field_var.get()
        if not field_name:
            return
            
        df = self.table_manager.get_dataframe()
        if field_name in df.columns:
            # 获取字段的唯一值和计数
            value_counts = df[field_name].value_counts()
            
            # 更新字段值预览
            self.values_text.config(state='normal')
            self.values_text.delete(1.0, tk.END)
            
            preview_text = f"共 {len(value_counts)} 个不同的值:\n\n"
            for value, count in value_counts.items():
                preview_text += f"• {value}: {count}行\n"
            
            # 添加总计信息
            total_rows = value_counts.sum()
            preview_text += f"\n总计: {total_rows}行数据"
                
            self.values_text.insert(1.0, preview_text)
            self.values_text.config(state='disabled')
            
        # 更新预览
        self.update_preview()
        
    def update_preview(self):
        """更新导出预览"""
        # 安全检查：确保preview_label已经创建
        if not hasattr(self, 'preview_label'):
            return
            
        try:
            mode = self.export_mode.get()
            df = self.table_manager.get_dataframe()
            
            if mode == "random":
                count = int(self.random_count_var.get() or 0)
                max_count = len(df)
                actual_count = min(count, max_count)
                preview_text = f"🎲 将随机导出 {actual_count} 行数据（共 {max_count} 行可用）"
                if count > max_count:
                    preview_text += f"\n⚠️ 请求数量({count})超过可用数据，已调整为{actual_count}行"
                
            else:  # stratified
                field_name = self.field_var.get()
                per_value_count = int(self.per_value_count_var.get() or 0)
                
                if field_name and field_name in df.columns:
                    value_counts = df[field_name].value_counts()
                    total_export = 0
                    
                    for value, available_count in value_counts.items():
                        actual_count = min(per_value_count, available_count)
                        total_export += actual_count
                        
                    preview_text = f"🎯 将按 '{field_name}' 字段分层导出，每个值最多 {per_value_count} 行\n"
                    preview_text += f"📊 预计导出 {total_export} 行数据（共 {len(value_counts)} 个不同值）"
                    
                    # 如果某些值的数据不足，给出提示
                    insufficient_values = []
                    for value, available_count in value_counts.items():
                        if available_count < per_value_count:
                            insufficient_values.append(f"{value}({available_count}行)")
                    
                    if insufficient_values:
                        preview_text += f"\n⚠️ 数据不足的值: {', '.join(insufficient_values)}"
                else:
                    preview_text = "请选择分层字段"
                    
            self.preview_label.config(text=preview_text, foreground='black')
            
        except ValueError:
            if hasattr(self, 'preview_label'):
                self.preview_label.config(text="请输入有效的数字", foreground='red')
        except Exception as e:
            if hasattr(self, 'preview_label'):
                self.preview_label.config(text=f"预览错误: {str(e)}", foreground='red')
            
    def preview_export(self):
        """预览导出结果"""
        try:
            export_df = self.get_export_data()
            if export_df is None:
                return
                
            # 创建预览窗口
            preview_window = tk.Toplevel(self.dialog)
            preview_window.title("导出预览")
            preview_window.geometry("800x600")
            
            # 创建文本框显示预览
            text_frame = ttk.Frame(preview_window, padding=10)
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            preview_text = tk.Text(text_frame, wrap=tk.NONE, font=('Consolas', 9))
            
            # 添加滚动条
            v_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=preview_text.yview)
            h_scrollbar = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=preview_text.xview)
            preview_text.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            # 布局
            preview_text.grid(row=0, column=0, sticky=tk.NSEW)
            v_scrollbar.grid(row=0, column=1, sticky=tk.NS)
            h_scrollbar.grid(row=1, column=0, sticky=tk.EW)
            
            text_frame.grid_rowconfigure(0, weight=1)
            text_frame.grid_columnconfigure(0, weight=1)
            
            # 显示数据
            preview_content = f"导出数据预览 ({len(export_df)} 行 × {len(export_df.columns)} 列)\n"
            preview_content += "=" * 60 + "\n\n"
            preview_content += export_df.to_string(max_rows=100, max_cols=10)
            
            if len(export_df) > 100:
                preview_content += f"\n\n... 还有 {len(export_df) - 100} 行数据 ..."
                
            preview_text.insert(1.0, preview_content)
            preview_text.config(state='disabled')
            
            # 关闭按钮
            close_btn = ttk.Button(preview_window, text="关闭", 
                                  command=preview_window.destroy)
            close_btn.pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("预览错误", f"预览失败: {str(e)}")
            
    def get_export_data(self):
        """获取要导出的数据"""
        try:
            df = self.table_manager.get_dataframe()
            mode = self.export_mode.get()
            
            if mode == "random":
                count = int(self.random_count_var.get() or 0)
                if count <= 0:
                    messagebox.showerror("参数错误", "导出数量必须大于0")
                    return None
                    
                # 随机采样
                actual_count = min(count, len(df))
                export_df = df.sample(n=actual_count, random_state=None).reset_index(drop=True)
                
            else:  # stratified
                field_name = self.field_var.get()
                per_value_count = int(self.per_value_count_var.get() or 0)
                
                if not field_name or field_name not in df.columns:
                    messagebox.showerror("参数错误", "请选择有效的分层字段")
                    return None
                    
                if per_value_count <= 0:
                    messagebox.showerror("参数错误", "每个值的数量必须大于0")
                    return None
                    
                # 按字段值分层采样
                export_dfs = []
                for value in df[field_name].unique():
                    value_df = df[df[field_name] == value]
                    actual_count = min(per_value_count, len(value_df))
                    if actual_count > 0:
                        sampled_df = value_df.sample(n=actual_count, random_state=None)
                        export_dfs.append(sampled_df)
                        
                if not export_dfs:
                    messagebox.showerror("导出错误", "没有可导出的数据")
                    return None
                    
                export_df = pd.concat(export_dfs, ignore_index=True)
                
            return export_df
            
        except ValueError as e:
            messagebox.showerror("参数错误", "请输入有效的数字")
            return None
        except Exception as e:
            messagebox.showerror("导出错误", f"获取导出数据失败: {str(e)}")
            return None
            
    def export_data(self):
        """导出数据"""
        try:
            export_df = self.get_export_data()
            if export_df is None:
                return
                
            # 选择保存文件
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            mode_name = "随机导出" if self.export_mode.get() == "random" else "分层导出"
            default_filename = f"{mode_name}_{timestamp}.aie"
            
            file_path = filedialog.asksaveasfilename(
                title="保存导出项目文件",
                defaultextension=".aie",
                filetypes=[
                    ("AI Excel项目文件", "*.aie"),
                    ("Excel文件", "*.xlsx"),
                    ("CSV文件", "*.csv"),
                    ("所有文件", "*.*")
                ],
                initialfile=default_filename
            )
            
            if not file_path:
                return
                
            # 保存文件
            if file_path.lower().endswith('.aie'):
                # 保存为AI Excel项目格式，包含AI列配置
                self.save_as_aie_project(export_df, file_path)
            elif file_path.lower().endswith('.csv'):
                export_df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                export_df.to_excel(file_path, index=False)
                
            # 显示成功信息
            mode_name = "随机导出" if self.export_mode.get() == "random" else "分层导出"
            file_ext = os.path.splitext(file_path)[1].lower()
            
            success_msg = f"🎉 {mode_name}成功！\n\n"
            success_msg += f"📊 导出行数: {len(export_df)} 行\n"
            success_msg += f"📁 保存位置: {file_path}\n\n"
            
            if file_ext == '.aie':
                success_msg += f"✨ 已保存为AI Excel项目文件\n"
                success_msg += f"🔧 包含完整的AI列配置\n"
                success_msg += f"💡 可直接在程序中打开使用\n\n"
            
            if self.export_mode.get() == "stratified":
                field_name = self.field_var.get()
                unique_values = export_df[field_name].nunique()
                success_msg += f"🎯 分层字段: {field_name}\n"
                success_msg += f"📈 包含类别: {unique_values} 个"
            
            messagebox.showinfo("导出成功", success_msg)
            
            # 关闭对话框
            self.result = {"success": True, "file_path": file_path, "row_count": len(export_df)}
            self.dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("导出错误", f"导出失败: {str(e)}")
            
    def save_as_aie_project(self, export_df, file_path):
        """保存为AI Excel项目格式"""
        try:
            import json
            
            # 获取原始的AI列配置
            ai_columns = self.table_manager.get_ai_columns()
            
            # 获取原始的长文本列配置
            long_text_columns = {}
            if hasattr(self.table_manager, 'long_text_columns'):
                long_text_columns = self.table_manager.long_text_columns.copy()
            
            # 获取列宽信息（如果有的话）
            column_widths = {}
            if hasattr(self.table_manager, 'column_widths'):
                column_widths = self.table_manager.column_widths.copy()
            
            # 创建项目数据
            project_data = {
                "dataframe": export_df.to_dict('records'),
                "ai_columns": ai_columns,
                "long_text_columns": long_text_columns,
                "column_widths": column_widths,
                "export_info": {
                    "export_mode": self.export_mode.get(),
                    "export_time": datetime.now().isoformat(),
                    "original_rows": len(self.table_manager.get_dataframe()),
                    "exported_rows": len(export_df)
                }
            }
            
            # 如果是分层导出，添加分层信息
            if self.export_mode.get() == "stratified":
                field_name = self.field_var.get()
                project_data["export_info"]["stratified_field"] = field_name
                project_data["export_info"]["per_value_count"] = int(self.per_value_count_var.get())
                
                # 添加各类别的导出统计
                if field_name in export_df.columns:
                    category_stats = export_df[field_name].value_counts().to_dict()
                    project_data["export_info"]["category_distribution"] = category_stats
            
            # 保存到文件
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
                
            print(f"成功保存AI Excel项目文件: {file_path}")
            
        except Exception as e:
            raise Exception(f"保存项目文件失败: {str(e)}")
    
    def on_cancel(self):
        """取消操作"""
        self.result = {"success": False}
        self.dialog.destroy()
        
    def show(self):
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result

def show_random_export_dialog(parent, table_manager):
    """显示随机导出对话框"""
    dialog = RandomExportDialog(parent, table_manager)
    return dialog.show() 