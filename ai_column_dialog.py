#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI列配置对话框
用于设置新建列的名称、类型和prompt模板
"""

import tkinter as tk
from tkinter import ttk, messagebox

class AIColumnDialog:
    def __init__(self, parent, existing_columns):
        self.parent = parent
        self.existing_columns = existing_columns
        self.result = None
        
        # 创建对话框窗口 - 增大尺寸
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("新建列")
        self.dialog.geometry("700x600")  # 从600x500增加到700x600
        self.dialog.resizable(True, True)
        
        # 设置模态
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.center_window()
        
        # 创建界面
        self.create_widgets()
        
    def center_window(self):
        """窗口居中"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (700 // 2)  # 更新居中计算
        y = (self.dialog.winfo_screenheight() // 2) - (600 // 2)  # 更新居中计算
        self.dialog.geometry(f"700x600+{x}+{y}")  # 更新几何尺寸
        
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 列名输入
        ttk.Label(main_frame, text="列名:").pack(anchor=tk.W, pady=(0, 5))
        self.column_name_var = tk.StringVar()
        self.column_name_entry = ttk.Entry(main_frame, textvariable=self.column_name_var, width=50)
        self.column_name_entry.pack(fill=tk.X, pady=(0, 10))
        
        # 列类型选择
        ttk.Label(main_frame, text="列类型:").pack(anchor=tk.W, pady=(0, 5))
        self.column_type_var = tk.StringVar(value="ai")
        
        type_frame = ttk.Frame(main_frame)
        type_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Radiobutton(type_frame, text="AI处理列", variable=self.column_type_var, 
                       value="ai", command=self.on_type_change).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="普通列", variable=self.column_type_var, 
                       value="normal", command=self.on_type_change).pack(side=tk.LEFT)
        
        # AI模型选择
        model_config_frame = ttk.Frame(main_frame)
        model_config_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(model_config_frame, text="AI模型:").pack(side=tk.LEFT, padx=(0, 10))
        self.model_var = tk.StringVar(value="gpt-4.1")
        self.model_combo = ttk.Combobox(model_config_frame, textvariable=self.model_var, 
                                       values=["gpt-4.1", "o1"], state="readonly", width=15)
        self.model_combo.pack(side=tk.LEFT)
        
        # 模型说明
        ttk.Label(model_config_frame, text="  (gpt-4.1: 快速响应 | o1: 深度推理)", 
                 foreground="gray", font=('Microsoft YaHei UI', 8)).pack(side=tk.LEFT, padx=(10, 0))
        
        # 保存模型配置框架的引用，用于显示/隐藏
        self.model_config_frame = model_config_frame
        
        # Prompt模板输入区域
        self.prompt_frame = ttk.LabelFrame(main_frame, text="AI Prompt模板", padding="10")
        self.prompt_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 10))
        
        # 提示信息
        tip_text = "在prompt中使用 {列名} 来引用其他字段的值\n例如: 请将以下{category}类的英文query翻译成中文：{query}"
        ttk.Label(self.prompt_frame, text=tip_text, foreground="gray").pack(anchor=tk.W, pady=(0, 5))
        
        # 可用字段显示
        if self.existing_columns:
            fields_label = ttk.Label(self.prompt_frame, text="可用字段: {文本内容}", foreground="red")
            fields_label.pack(anchor=tk.W, pady=(0, 5))
            
            # 创建可选择的字段文本框
            fields_frame = ttk.Frame(self.prompt_frame)
            fields_frame.pack(fill=tk.X, pady=(0, 10))
            
            # 字段文本框 - 只读但可选择复制
            self.fields_text = tk.Text(fields_frame, height=3, wrap=tk.WORD, 
                                     background='#f8f9fa', relief='solid', borderwidth=1,
                                     font=('Consolas', 9))
            self.fields_text.pack(fill=tk.X)
            
            # 填充字段内容，每行显示几个字段
            fields_content = ""
            fields_list = [f"{{{col}}}" for col in self.existing_columns]
            
            # 按行排列字段，每行最多4个
            for i in range(0, len(fields_list), 4):
                line_fields = fields_list[i:i+4]
                fields_content += "  ".join(line_fields) + "\n"
            
            self.fields_text.insert("1.0", fields_content.strip())
            self.fields_text.config(state=tk.DISABLED)  # 设为只读但可选择
            
            # 添加右键复制菜单
            self.create_fields_context_menu()
            
            # 提示标签
            tip_label = ttk.Label(fields_frame, text="💡 双击字段名可快速复制到剪贴板", 
                                foreground="gray", font=('Microsoft YaHei UI', 8))
            tip_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Prompt文本框 - 减少高度为按钮留出空间
        self.prompt_text = tk.Text(self.prompt_frame, height=8, wrap=tk.WORD, width=80)  # 从height=12减少到8
        self.prompt_text.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.prompt_frame, orient=tk.VERTICAL, command=self.prompt_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.prompt_text.configure(yscrollcommand=scrollbar.set)
        
        # 按钮框架 - 确保按钮可见，不使用expand
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 10))  # 移除side=tk.BOTTOM
        
        # 创建按钮并确保可见性
        cancel_btn = ttk.Button(button_frame, text="取消", command=self.on_cancel, width=10)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 10))
        
        ok_btn = ttk.Button(button_frame, text="确定", command=self.on_ok, width=10)
        ok_btn.pack(side=tk.RIGHT, padx=(0, 5))
        
        # 设置初始焦点
        self.column_name_entry.focus()
        
        # 绑定回车键
        self.dialog.bind('<Return>', lambda e: self.on_ok())
        self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        
    def on_type_change(self):
        """列类型改变时的处理"""
        if self.column_type_var.get() == "ai":
            # 显示AI相关配置
            self.model_config_frame.pack(fill=tk.X, pady=(0, 10))
            self.prompt_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            self.prompt_text.config(state=tk.NORMAL)
        else:
            # 隐藏AI相关配置
            self.model_config_frame.pack_forget()
            self.prompt_frame.pack_forget()
            
    def validate_input(self):
        """验证输入"""
        column_name = self.column_name_var.get().strip()
        
        # 检查列名是否为空
        if not column_name:
            messagebox.showerror("错误", "请输入列名")
            return False
            
        # 检查列名是否已存在
        if column_name in self.existing_columns:
            messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
            return False
            
        # 如果是AI列，检查prompt是否为空
        if self.column_type_var.get() == "ai":
            prompt = self.prompt_text.get("1.0", tk.END).strip()
            if not prompt:
                messagebox.showerror("错误", "请输入AI处理的prompt模板")
                return False
                
        return True
        
    def on_ok(self):
        """确定按钮处理"""
        if self.validate_input():
            column_name = self.column_name_var.get().strip()
            is_ai_column = self.column_type_var.get() == "ai"
            
            if is_ai_column:
                prompt_template = self.prompt_text.get("1.0", tk.END).strip()
                ai_model = self.model_var.get()
                self.result = (column_name, prompt_template, True, ai_model)
            else:
                self.result = (column_name, "", False, None)
                
            self.dialog.destroy()
            
    def on_cancel(self):
        """取消按钮处理"""
        self.result = None
        self.dialog.destroy()
        
    def show(self):
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 

    def create_fields_context_menu(self):
        """创建字段文本框的右键菜单"""
        self.fields_menu = tk.Menu(self.dialog, tearoff=0)
        self.fields_menu.add_command(label="📋 复制选中内容", command=self.copy_selected_field)
        self.fields_menu.add_command(label="📋 复制全部字段", command=self.copy_all_fields)
        
        # 绑定右键菜单
        self.fields_text.bind("<Button-3>", self.show_fields_menu)
        # 绑定双击复制
        self.fields_text.bind("<Double-Button-1>", self.on_field_double_click)
        
    def show_fields_menu(self, event):
        """显示字段右键菜单"""
        try:
            self.fields_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.fields_menu.grab_release()
            
    def copy_selected_field(self):
        """复制选中的字段"""
        try:
            # 临时启用文本框以获取选中内容
            self.fields_text.config(state=tk.NORMAL)
            selected_text = self.fields_text.selection_get()
            self.fields_text.config(state=tk.DISABLED)
            
            if selected_text.strip():
                self.dialog.clipboard_clear()
                self.dialog.clipboard_append(selected_text.strip())
                # 简单的视觉反馈
                original_bg = self.fields_text.cget('background')
                self.fields_text.config(background='#e6ffe6')
                self.dialog.after(200, lambda: self.fields_text.config(background=original_bg))
        except tk.TclError:
            # 没有选中内容
            pass
            
    def copy_all_fields(self):
        """复制所有字段"""
        self.fields_text.config(state=tk.NORMAL)
        all_text = self.fields_text.get("1.0", tk.END).strip()
        self.fields_text.config(state=tk.DISABLED)
        
        if all_text:
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(all_text)
            # 视觉反馈
            original_bg = self.fields_text.cget('background')
            self.fields_text.config(background='#e6ffe6')
            self.dialog.after(200, lambda: self.fields_text.config(background=original_bg))
            
    def on_field_double_click(self, event):
        """双击字段时的处理"""
        try:
            # 获取双击位置的单词
            self.fields_text.config(state=tk.NORMAL)
            
            # 获取点击位置
            index = self.fields_text.index(f"@{event.x},{event.y}")
            
            # 选择当前单词（字段）
            word_start = self.fields_text.index(f"{index} wordstart")
            word_end = self.fields_text.index(f"{index} wordend")
            
            # 选中单词
            self.fields_text.tag_remove(tk.SEL, "1.0", tk.END)
            self.fields_text.tag_add(tk.SEL, word_start, word_end)
            
            # 获取选中的文本
            selected_word = self.fields_text.get(word_start, word_end)
            
            self.fields_text.config(state=tk.DISABLED)
            
            # 如果是字段格式，复制到剪贴板
            if selected_word.startswith('{') and selected_word.endswith('}'):
                self.dialog.clipboard_clear()
                self.dialog.clipboard_append(selected_word)
                
                # 视觉反馈
                original_bg = self.fields_text.cget('background')
                self.fields_text.config(background='#e6ffe6')
                self.dialog.after(300, lambda: self.fields_text.config(background=original_bg))
                
        except tk.TclError:
            pass 