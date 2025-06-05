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
        
        # 创建对话框窗口 - 增大尺寸以容纳参数配置
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("新建AI列")
        self.dialog.geometry("700x650")  # 增加高度
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
        x = (self.dialog.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (650 // 2)
        self.dialog.geometry(f"700x650+{x}+{y}")
        
    def create_widgets(self):
        """创建界面组件"""
        # 创建一个Canvas和Scrollbar来实现滚动
        canvas = tk.Canvas(self.dialog)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 主框架
        main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 列名输入框（将在多字段模式下隐藏）
        self.column_name_frame = ttk.Frame(main_frame)
        self.column_name_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(self.column_name_frame, text="列名:").pack(anchor=tk.W, pady=(0, 5))
        self.column_name_var = tk.StringVar()
        self.column_name_entry = ttk.Entry(self.column_name_frame, textvariable=self.column_name_var, width=50)
        self.column_name_entry.pack(fill=tk.X)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            try:
                # 检查canvas是否仍然存在且有效
                if hasattr(self, 'canvas') and self.canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except (tk.TclError, AttributeError):
                # Canvas已被销毁或不存在，忽略错误
                pass
        self._mousewheel_callback = _on_mousewheel
        canvas.bind_all("<MouseWheel>", self._mousewheel_callback)
        
        # 存储canvas引用以便后续使用
        self.canvas = canvas
        self.scrollbar = scrollbar
        
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
        
        # 输出模式选择
        output_mode_frame = ttk.LabelFrame(main_frame, text="输出模式", padding="10")
        output_mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.output_mode_var = tk.StringVar(value="single")
        ttk.Radiobutton(output_mode_frame, text="单字段输出 (传统模式)", 
                       variable=self.output_mode_var, value="single", 
                       command=self.on_output_mode_change).pack(anchor='w', pady=2)
        ttk.Radiobutton(output_mode_frame, text="多字段输出 (JSON格式)", 
                       variable=self.output_mode_var, value="multi", 
                       command=self.on_output_mode_change).pack(anchor='w', pady=2)
        
        # 多字段配置框架
        self.multi_field_frame = ttk.LabelFrame(main_frame, text="多字段配置", padding="10")
        
        # 任务描述输入（多字段模式专用）
        task_desc_frame = ttk.Frame(self.multi_field_frame)
        task_desc_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(task_desc_frame, text="任务描述:").pack(anchor=tk.W, pady=(0, 5))
        self.task_description_var = tk.StringVar()
        self.task_description_entry = ttk.Entry(task_desc_frame, textvariable=self.task_description_var, width=50)
        self.task_description_entry.pack(fill=tk.X)
        ttk.Label(task_desc_frame, text="💡 将作为临时列名，存储AI原始响应，处理完成后会解析为多个字段列", 
                 foreground="gray", font=('Microsoft YaHei UI', 8)).pack(anchor=tk.W, pady=(2, 0))
        
        # 字段列表管理
        fields_control_frame = ttk.Frame(self.multi_field_frame)
        fields_control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(fields_control_frame, text="输出字段:").pack(side=tk.LEFT, padx=(0, 10))
        
        # 字段输入框和添加按钮
        self.new_field_var = tk.StringVar()
        self.new_field_entry = ttk.Entry(fields_control_frame, textvariable=self.new_field_var, width=15)
        self.new_field_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(fields_control_frame, text="添加字段", command=self.add_output_field, width=8).pack(side=tk.LEFT, padx=(0, 10))
        
        # 字段列表
        fields_list_frame = ttk.Frame(self.multi_field_frame)
        fields_list_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.fields_listbox = tk.Listbox(fields_list_frame, height=4, selectmode=tk.SINGLE)
        self.fields_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 字段操作按钮
        fields_buttons_frame = ttk.Frame(fields_list_frame)
        fields_buttons_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        
        ttk.Button(fields_buttons_frame, text="删除", command=self.remove_output_field, width=6).pack(pady=2)
        ttk.Button(fields_buttons_frame, text="清空", command=self.clear_output_fields, width=6).pack(pady=2)
        
        # 预设字段按钮
        preset_frame = ttk.Frame(self.multi_field_frame)
        preset_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(preset_frame, text="常用预设:").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(preset_frame, text="query/answer", command=lambda: self.add_preset_fields(["query", "answer"]), width=12).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(preset_frame, text="query/golden_answer", command=lambda: self.add_preset_fields(["query", "golden_answer"]), width=14).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(preset_frame, text="question/answer/explanation", command=lambda: self.add_preset_fields(["question", "answer", "explanation"]), width=18).pack(side=tk.LEFT)
        
        # JSON示例显示
        example_frame = ttk.Frame(self.multi_field_frame)
        example_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(example_frame, text="JSON输出格式示例:", foreground="gray", font=('Microsoft YaHei UI', 8)).pack(anchor='w')
        
        self.json_example_text = tk.Text(example_frame, height=3, wrap=tk.WORD, 
                                        background='#f0f0f0', relief='solid', borderwidth=1,
                                        font=('Consolas', 8), state=tk.DISABLED)
        self.json_example_text.pack(fill=tk.X, pady=(2, 0))
        
        # 初始化字段列表
        self.output_fields = []
        
        # 处理参数配置框架
        self.params_frame = ttk.LabelFrame(main_frame, text="处理参数配置", padding="10")
        self.params_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 并发数设置
        concurrent_frame = ttk.Frame(self.params_frame)
        concurrent_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(concurrent_frame, text="并发数:").pack(side=tk.LEFT, padx=(0, 10))
        self.max_workers_var = tk.IntVar(value=3)
        max_workers_spinbox = ttk.Spinbox(concurrent_frame, from_=1, to=999, 
                                         textvariable=self.max_workers_var, width=5)
        max_workers_spinbox.pack(side=tk.LEFT)
        ttk.Label(concurrent_frame, text="  (同时处理的任务数，建议1-5)", 
                 foreground="gray", font=('Microsoft YaHei UI', 8)).pack(side=tk.LEFT, padx=(5, 0))
        
        # 请求延迟设置
        delay_frame = ttk.Frame(self.params_frame)
        delay_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(delay_frame, text="请求延迟:").pack(side=tk.LEFT, padx=(0, 10))
        self.request_delay_var = tk.DoubleVar(value=0.5)
        delay_spinbox = ttk.Spinbox(delay_frame, from_=0.1, to=5.0, increment=0.1,
                                   textvariable=self.request_delay_var, width=5)
        delay_spinbox.pack(side=tk.LEFT)
        ttk.Label(delay_frame, text="秒  (避免API限流，建议0.3-1.0)", 
                 foreground="gray", font=('Microsoft YaHei UI', 8)).pack(side=tk.LEFT, padx=(5, 0))
        
        # 重试次数设置
        retry_frame = ttk.Frame(self.params_frame)
        retry_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(retry_frame, text="重试次数:").pack(side=tk.LEFT, padx=(0, 10))
        self.max_retries_var = tk.IntVar(value=2)
        retry_spinbox = ttk.Spinbox(retry_frame, from_=0, to=5, 
                                   textvariable=self.max_retries_var, width=5)
        retry_spinbox.pack(side=tk.LEFT)
        ttk.Label(retry_frame, text="  (API失败时的重试次数，建议1-3)", 
                 foreground="gray", font=('Microsoft YaHei UI', 8)).pack(side=tk.LEFT, padx=(5, 0))
        
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
        
        # Prompt文本框 - 增大高度以便编辑
        self.prompt_text = tk.Text(self.prompt_frame, height=10, wrap=tk.WORD, width=80)  # 增加到10行
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
        
        # 初始状态：只显示AI相关配置
        self.on_type_change()
        
        # 初始化输出模式
        self.on_output_mode_change()
        
        # 绑定回车键
        self.dialog.bind('<Return>', lambda e: self.on_ok())
        self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        
        # 绑定字段输入框回车键
        self.new_field_entry.bind('<Return>', lambda e: self.add_output_field())
        
        # 布局Canvas和Scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def on_type_change(self):
        """列类型改变时的处理 - Simplified: AI is the only type"""
        # column_type = self.column_type_var.get() # No longer needed
        
        # Always show AI config
        self.model_config_frame.pack(fill=tk.X, pady=(0, 10))
        self.params_frame.pack(fill=tk.X, pady=(0, 10))
        self.prompt_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        # self.long_text_frame.pack_forget() # long_text_frame is removed
        
    def on_output_mode_change(self):
        """输出模式改变时的处理"""
        mode = self.output_mode_var.get()
        
        if mode == "multi":
            # 隐藏列名输入框（多字段模式下不需要用户输入列名）
            self.column_name_frame.pack_forget()
            # 显示多字段配置
            self.multi_field_frame.pack(fill=tk.X, pady=(0, 10), after=self.model_config_frame)
            # 更新JSON示例
            self.update_json_example()
        else:
            # 显示列名输入框（单字段模式需要用户输入列名）
            self.column_name_frame.pack(fill=tk.X, pady=(0, 10), before=self.model_config_frame)
            # 隐藏多字段配置
            self.multi_field_frame.pack_forget()
            
    def add_output_field(self):
        """添加输出字段"""
        field_name = self.new_field_var.get().strip()
        
        if not field_name:
            messagebox.showwarning("警告", "请输入字段名")
            return
            
        # 检查字段名是否已存在
        if field_name in self.output_fields:
            messagebox.showwarning("警告", f"字段 '{field_name}' 已存在")
            return
            
        # 检查字段名格式（简单验证）
        import re
        if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', field_name):
            messagebox.showerror("错误", "字段名只能包含字母、数字和下划线，且不能以数字开头")
            return
            
        # 添加到列表
        self.output_fields.append(field_name)
        self.fields_listbox.insert(tk.END, field_name)
        
        # 清空输入框
        self.new_field_var.set("")
        
        # 更新JSON示例
        self.update_json_example()
        
    def remove_output_field(self):
        """删除选中的输出字段"""
        selection = self.fields_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请选择要删除的字段")
            return
            
        index = selection[0]
        field_name = self.output_fields[index]
        
        # 从列表中删除
        del self.output_fields[index]
        self.fields_listbox.delete(index)
        
        # 更新JSON示例
        self.update_json_example()
        
    def clear_output_fields(self):
        """清空所有输出字段"""
        if self.output_fields:
            result = messagebox.askyesno("确认", "确定要清空所有字段吗？")
            if result:
                self.output_fields.clear()
                self.fields_listbox.delete(0, tk.END)
                self.update_json_example()
                
    def add_preset_fields(self, fields):
        """添加预设字段"""
        added_fields = []
        
        for field in fields:
            if field not in self.output_fields:
                self.output_fields.append(field)
                self.fields_listbox.insert(tk.END, field)
                added_fields.append(field)
                
        if added_fields:
            self.update_json_example()
            messagebox.showinfo("添加成功", f"已添加字段: {', '.join(added_fields)}")
        else:
            messagebox.showinfo("提示", "所有字段都已存在")
            
    def update_json_example(self):
        """更新JSON示例显示"""
        if not self.output_fields:
            example_text = "请先添加输出字段"
        else:
            # 生成JSON示例
            example_dict = {}
            for field in self.output_fields:
                example_dict[field] = f"这里是{field}的内容"
                
            import json
            example_text = json.dumps(example_dict, ensure_ascii=False, indent=2)
            
        # 更新显示
        self.json_example_text.config(state=tk.NORMAL)
        self.json_example_text.delete("1.0", tk.END)
        self.json_example_text.insert("1.0", example_text)
        self.json_example_text.config(state=tk.DISABLED)
    
    def update_filename_field_list(self): # REMOVED - was for long_text_frame
        pass
        # """更新文件名字段列表"""
        # if self.existing_columns:
            # self.filename_field_combo['values'] = self.existing_columns
            # if self.existing_columns:
                # self.filename_field_combo.set(self.existing_columns[0])
        # else:
            # self.filename_field_combo['values'] = []
    
    def test_search(self): # REMOVED - was for long_text_frame
        pass
        # """测试搜索功能"""
        # folder_path = self.folder_path_var.get().strip()
        # if not folder_path:
            # self.test_result_label.config(text="请先选择文件夹", foreground="red")
            # return
        
        # try:
            # from paper_processor import get_paper_processor
            # processor = get_paper_processor()
            
            # # 搜索文件
            # files_map = processor.find_files_in_folder(folder_path)
            
            # if files_map:
                # count = len(files_map)
                # # 显示前几个文件名作为示例
                # sample_files = list(files_map.keys())[:3]
                # sample_text = ", ".join(sample_files)
                # if len(files_map) > 3:
                    # sample_text += "..."
                
                # self.test_result_label.config(
                    # text=f"找到 {count} 个txt文件 (如: {sample_text})", 
                    # foreground="green"
                # )
            # else:
                # self.test_result_label.config(text="未找到任何txt文件", foreground="orange")
                
        # except Exception as e:
            # self.test_result_label.config(text=f"搜索失败: {str(e)}", foreground="red")
            
    def validate_input(self):
        """验证输入"""
        # 检查输出模式
        output_mode = self.output_mode_var.get()
        
        if output_mode == "single":
            # 单字段模式：需要验证列名
            column_name = self.column_name_var.get().strip()
            if not column_name:
                messagebox.showerror("错误", "请输入列名")
                return False
                
            if column_name in self.existing_columns:
                messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
                return False
        else:
            # 多字段模式：检查任务描述和输出字段
            task_description = self.task_description_var.get().strip()
            if not task_description:
                messagebox.showerror("错误", "请输入任务描述")
                return False
                
            if not self.output_fields:
                messagebox.showerror("错误", "多字段模式下请至少添加一个输出字段")
                return False
                
            # 检查任务描述（临时列名）是否冲突
            if task_description in self.existing_columns:
                messagebox.showerror("错误", f"任务描述 '{task_description}' 与现有列名冲突")
                return False
                
            # 多字段模式下，最终会生成以字段名为名的列
            # 检查字段名是否会与现有列冲突
            conflicting_fields = []
            for field in self.output_fields:
                if field in self.existing_columns:
                    conflicting_fields.append(field)
                    
            if conflicting_fields:
                messagebox.showerror("错误", f"以下字段名与现有列名冲突: {', '.join(conflicting_fields)}")
                return False
            
        # 验证prompt
        prompt = self.prompt_text.get("1.0", tk.END).strip()
        if not prompt:
            messagebox.showerror("错误", "请输入AI Prompt模板")
            return False
            
        # 多字段模式：检查prompt是否包含JSON要求
        if self.output_mode_var.get() == "multi":
            prompt_lower = prompt.lower()
            if "json" not in prompt_lower:
                result = messagebox.askyesno("确认", 
                    "您的prompt模板中没有明确要求AI输出JSON格式。\n"
                    "多字段模式需要AI输出标准JSON格式。\n\n"
                    "是否自动在prompt末尾添加JSON格式要求？")
                if result:
                    # 自动添加JSON格式要求
                    json_instruction = f"\n\n请以JSON格式输出，包含以下字段：{', '.join(self.output_fields)}"
                    prompt += json_instruction
                    self.prompt_text.delete("1.0", tk.END)
                    self.prompt_text.insert("1.0", prompt)
            
        # 验证处理参数
        max_workers = self.max_workers_var.get()
        if max_workers < 1:
            messagebox.showerror("错误", "并发数必须大于等于1")
            return False
            
        request_delay = self.request_delay_var.get()
        if request_delay < 0.1 or request_delay > 5.0:
            messagebox.showerror("错误", "请求延迟必须在0.1-5.0秒之间")
            return False
            
        max_retries = self.max_retries_var.get()
        if max_retries < 0 or max_retries > 5:
            messagebox.showerror("错误", "重试次数必须在0-5之间")
            return False
            
        return True
    
    def on_ok(self):
        """确定按钮点击事件"""
        if not self.validate_input():
            return
            
        model = self.model_var.get()
        prompt = self.prompt_text.get("1.0", tk.END).strip()
        output_mode = self.output_mode_var.get()
        
        # 获取列名：单字段模式从输入框获取，多字段模式从任务描述获取
        if output_mode == "single":
            column_name = self.column_name_var.get().strip()
        else:
            column_name = self.task_description_var.get().strip()  # 多字段模式使用任务描述作为临时列名
        
        # 获取处理参数
        processing_params = {
            'max_workers': self.max_workers_var.get(),
            'request_delay': self.request_delay_var.get(),
            'max_retries': self.max_retries_var.get()
        }
        
        self.result = {
            'column_name': column_name,  # 多字段模式下为任务描述
            'model': model,
            'prompt': prompt,
            'processing_params': processing_params,
            'output_mode': output_mode,
            'output_fields': self.output_fields.copy() if output_mode == "multi" else []
        }
        
        # 解绑滚轮事件
        if hasattr(self, '_mousewheel_callback'):
            try:
                self.dialog.unbind_all("<MouseWheel>")
            except:
                pass
        self.dialog.destroy()
        
    def on_cancel(self):
        """取消按钮点击事件"""
        self.result = None
        # 解绑滚轮事件
        if hasattr(self, '_mousewheel_callback'):
            try:
                self.dialog.unbind_all("<MouseWheel>")
            except:
                pass
        self.dialog.destroy()
        
    def show(self):
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result
        
    def create_fields_context_menu(self):
        """创建字段文本框的右键菜单"""
        self.fields_menu = tk.Menu(self.dialog, tearoff=0)
        self.fields_menu.add_command(label="复制选中字段", command=self.copy_selected_field)
        self.fields_menu.add_command(label="复制所有字段", command=self.copy_all_fields)
        
        # 绑定右键菜单
        self.fields_text.bind("<Button-3>", self.show_fields_menu)
        
    def show_fields_menu(self, event):
        """显示字段右键菜单"""
        try:
            self.fields_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.fields_menu.grab_release()
            
    def copy_selected_field(self):
        """复制选中的字段"""
        try:
            # 临时启用文本框以获取选择
            self.fields_text.config(state=tk.NORMAL)
            
            # 获取选中文本
            if self.fields_text.tag_ranges(tk.SEL):
                selected_text = self.fields_text.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
                if selected_text:
                    self.dialog.clipboard_clear()
                    self.dialog.clipboard_append(selected_text)
                    messagebox.showinfo("复制成功", f"已复制: {selected_text}")
                else:
                    messagebox.showwarning("提示", "请先选择要复制的字段")
            else:
                messagebox.showwarning("提示", "请先选择要复制的字段")
                
        finally:
            # 恢复只读状态
            self.fields_text.config(state=tk.DISABLED)
            
    def copy_all_fields(self):
        """复制所有字段"""
        try:
            # 临时启用文本框以获取内容
            self.fields_text.config(state=tk.NORMAL)
            all_text = self.fields_text.get("1.0", tk.END).strip()
            
            if all_text:
                self.dialog.clipboard_clear()
                self.dialog.clipboard_append(all_text)
                messagebox.showinfo("复制成功", "已复制所有字段到剪贴板")
            
        finally:
            # 恢复只读状态
            self.fields_text.config(state=tk.DISABLED)
            
    def on_field_double_click(self, event):
        """字段双击事件 - 复制到剪贴板"""
        try:
            # 临时启用文本框
            self.fields_text.config(state=tk.NORMAL)
            
            # 获取点击位置的字符
            index = self.fields_text.index(f"@{event.x},{event.y}")
            
            # 获取当前行
            line_start = self.fields_text.index(f"{index} linestart")
            line_end = self.fields_text.index(f"{index} lineend")
            line_text = self.fields_text.get(line_start, line_end)
            
            # 找到点击的字段
            import re
            fields_in_line = re.findall(r'\{[^}]+\}', line_text)
            
            if fields_in_line:
                # 简单选择第一个字段（或者可以改进为选择最接近的）
                selected_field = fields_in_line[0]
                self.dialog.clipboard_clear()
                self.dialog.clipboard_append(selected_field)
                messagebox.showinfo("复制成功", f"已复制字段: {selected_field}")
                
        except Exception as e:
            print(f"双击复制错误: {e}")
        finally:
            # 恢复只读状态
            self.fields_text.config(state=tk.DISABLED) 