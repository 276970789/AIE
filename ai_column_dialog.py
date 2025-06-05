#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI列配置对话框
用于设置新建列的名称、类型和prompt模板，也支持编辑现有AI列
"""

import tkinter as tk
from tkinter import ttk, messagebox

class AIColumnDialog:
    def __init__(self, parent, existing_columns, edit_mode=False, edit_column_name=None, edit_config=None):
        self.parent = parent
        self.existing_columns = existing_columns
        self.edit_mode = edit_mode
        self.edit_column_name = edit_column_name
        self.edit_config = edit_config or {}
        self.result = None
        
        # 创建对话框窗口 - 增大尺寸以容纳参数配置
        self.dialog = tk.Toplevel(parent)
        if edit_mode:
            self.dialog.title(f"编辑AI列配置 - {edit_column_name}")
        else:
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
        
        # 如果是编辑模式，预填充数据
        if edit_mode and edit_config:
            self.populate_edit_data()
        
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
        
        # 列名输入框（在编辑模式下显示为只读标签）
        self.column_name_frame = ttk.Frame(main_frame)
        self.column_name_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(self.column_name_frame, text="列名:").pack(anchor=tk.W, pady=(0, 5))
        
        if self.edit_mode:
            # 编辑模式：显示为只读标签
            self.column_name_display = ttk.Label(self.column_name_frame, text=self.edit_column_name, 
                                               style='Subtitle.TLabel', background='#f8f9fa', 
                                               relief='solid', padding=5)
            self.column_name_display.pack(fill=tk.X)
        else:
            # 新建模式：显示为输入框
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
                                       values=["gpt-4.1", "o1","claude-3-7-sonnet-20250219"], state="readonly", width=15)
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
        
        self.task_description_var = tk.StringVar() # Always create the variable

        if not self.edit_mode:
            ttk.Label(task_desc_frame, text="任务描述:").pack(anchor=tk.W, pady=(0, 5))
            self.task_description_entry = ttk.Entry(task_desc_frame, textvariable=self.task_description_var, width=50)
            self.task_description_entry.pack(fill=tk.X)
            ttk.Label(task_desc_frame, text="💡 将作为临时列名，存储AI原始响应，处理完成后会解析为多个字段列", 
                     foreground="gray", font=('Microsoft YaHei UI', 8)).pack(anchor=tk.W, pady=(2, 0))
        else:
            # In edit mode, display the AI column name (task description) as a read-only label
            ttk.Label(task_desc_frame, text="主AI列名 (任务描述):").pack(anchor=tk.W, pady=(0, 5))
            self.task_description_label = ttk.Label(task_desc_frame, textvariable=self.task_description_var, width=50, relief="sunken", padding=(2,2))
            self.task_description_label.pack(fill=tk.X)
        
        # 字段处理方式选择
        field_mode_frame = ttk.Frame(self.multi_field_frame)
        field_mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(field_mode_frame, text="字段处理方式:").pack(anchor=tk.W, pady=(0, 5))
        self.field_mode_var = tk.StringVar(value="predefined")
        ttk.Radiobutton(field_mode_frame, text="预定义字段 (推荐)", 
                       variable=self.field_mode_var, value="predefined",
                       command=self.on_field_mode_change).pack(anchor='w', pady=2)
        ttk.Radiobutton(field_mode_frame, text="自动解析JSON字段 (实验性)", 
                       variable=self.field_mode_var, value="auto_parse",
                       command=self.on_field_mode_change).pack(anchor='w', pady=2)
        
        # 字段列表管理框架
        self.predefined_fields_frame = ttk.Frame(self.multi_field_frame)
        self.predefined_fields_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 字段列表管理
        fields_control_frame = ttk.Frame(self.predefined_fields_frame)
        fields_control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(fields_control_frame, text="输出字段:").pack(side=tk.LEFT, padx=(0, 10))
        
        # 字段输入框和添加按钮
        self.new_field_var = tk.StringVar()
        self.new_field_entry = ttk.Entry(fields_control_frame, textvariable=self.new_field_var, width=15)
        self.new_field_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(fields_control_frame, text="添加", command=self.add_output_field).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(fields_control_frame, text="删除", command=self.remove_output_field).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(fields_control_frame, text="清空", command=self.clear_output_fields).pack(side=tk.LEFT, padx=(0, 5))
        
        # 解析prompt功能
        parse_frame = ttk.Frame(fields_control_frame)
        parse_frame.pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(parse_frame, text="解析Prompt", command=self.parse_prompt_for_fields).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(parse_frame, text="从下方Prompt中自动提取字段", 
                 foreground="gray", font=('Microsoft YaHei UI', 8)).pack(side=tk.LEFT)
        
        # 字段列表显示
        fields_list_frame = ttk.Frame(self.predefined_fields_frame)
        fields_list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        ttk.Label(fields_list_frame, text="已添加的字段:").pack(anchor=tk.W, pady=(0, 5))
        self.fields_listbox = tk.Listbox(fields_list_frame, height=4)
        self.fields_listbox.pack(fill=tk.BOTH, expand=True)
        
        # 自动解析说明框架
        self.auto_parse_frame = ttk.Frame(self.multi_field_frame)
        
        auto_parse_info = ttk.Label(self.auto_parse_frame, 
                                   text="🔍 自动解析模式说明:\n"
                                        "• AI返回JSON后，系统会自动提取其中的所有字段\n"
                                        "• 适合不确定输出字段名称的场景\n"
                                        "• 建议在prompt中要求AI输出标准JSON格式",
                                   foreground="blue", 
                                   font=('Microsoft YaHei UI', 9),
                                   justify=tk.LEFT)
        auto_parse_info.pack(anchor=tk.W, pady=(0, 10))
        
        # JSON示例显示
        json_example_frame = ttk.LabelFrame(self.multi_field_frame, text="JSON输出示例", padding="5")
        json_example_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.json_example_text = tk.Text(json_example_frame, height=4, wrap=tk.WORD, 
                                        background='#f8f9fa', relief='flat',
                                        font=('Consolas', 9))
        self.json_example_text.pack(fill=tk.X)
        self.json_example_text.config(state=tk.DISABLED)
        
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
        prompt_frame = ttk.LabelFrame(main_frame, text="AI Prompt模板", padding="10")
        prompt_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 提示信息
        tip_text = "在prompt中使用 {列名} 来引用其他字段的值\n例如: 请将以下{category}类的英文query翻译成中文：{query}"
        ttk.Label(prompt_frame, text=tip_text, foreground="gray").pack(anchor=tk.W, pady=(0, 5))
        
        # 可用字段显示
        available_columns = [col for col in self.existing_columns if col != self.edit_column_name] if self.edit_mode else self.existing_columns
        if available_columns:
            fields_label = ttk.Label(prompt_frame, text="可用字段: (双击复制)", foreground="red")
            fields_label.pack(anchor=tk.W, pady=(0, 5))
            
            # 创建可选择的字段文本框
            fields_frame = ttk.Frame(prompt_frame)
            fields_frame.pack(fill=tk.X, pady=(0, 10))
            
            # 字段文本框 - 只读但可选择复制
            self.fields_text = tk.Text(fields_frame, height=3, wrap=tk.WORD, 
                                 background='#f8f9fa', relief='solid', borderwidth=1,
                                 font=('Consolas', 9))
            self.fields_text.pack(fill=tk.X)
            
            # 填充字段内容
            fields_content = ""
            fields_list = [f"{{{col}}}" for col in available_columns]
            
            # 按行排列字段，每行最多4个
            for i in range(0, len(fields_list), 4):
                line_fields = fields_list[i:i+4]
                fields_content += "  ".join(line_fields) + "\n"
            
            self.fields_text.insert("1.0", fields_content.strip())
            self.fields_text.config(state=tk.DISABLED)
            
            # 添加双击复制功能
            def on_field_double_click(event):
                try:
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
                    self.fields_text.config(state=tk.DISABLED)
                except:
                    self.fields_text.config(state=tk.DISABLED)
            
            self.fields_text.bind("<Double-Button-1>", on_field_double_click)
            
            # 提示标签
            tip_label = ttk.Label(fields_frame, text="💡 双击字段名可快速复制到剪贴板", 
                                foreground="gray", font=('Microsoft YaHei UI', 8))
            tip_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Prompt文本框
        self.prompt_text = tk.Text(prompt_frame, height=8, wrap=tk.WORD, width=80)
        self.prompt_text.pack(fill=tk.BOTH, expand=True)
        self.prompt_text.focus()
        
        # 滚动条
        prompt_scrollbar = ttk.Scrollbar(prompt_frame, orient=tk.VERTICAL, command=self.prompt_text.yview)
        prompt_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.prompt_text.configure(yscrollcommand=prompt_scrollbar.set)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 根据模式设置按钮文本
        save_text = "保存配置" if self.edit_mode else "创建列"
        ttk.Button(button_frame, text=save_text, command=self.on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=self.on_cancel).pack(side=tk.RIGHT)
        
        # 布局canvas和scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定快捷键
        self.dialog.bind('<Control-Return>', lambda e: self.on_ok())
        self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        
        # 初始化界面状态
        self.on_output_mode_change()
        self.on_field_mode_change()
        
    def populate_edit_data(self):
        """在编辑模式下预填充数据"""
        if not self.edit_mode or not self.edit_config:
            return
            
        # 填充模型 (Model first, as it's independent of output_mode)
        model = self.edit_config.get("model", "gpt-4.1")
        self.model_var.set(model)

        # 填充输出模式并立即更新UI
        output_mode = self.edit_config.get("output_mode", "single")
        self.output_mode_var.set(output_mode)
        self.on_output_mode_change() # Crucial: Update UI to show/hide multi_field_frame

        # 填充Prompt
        prompt = self.edit_config.get("prompt", "")
        self.prompt_text.insert("1.0", prompt)
        
        # 如果是多字段模式，将列名填充到任务描述中 (which is now a label in edit mode)
        if output_mode == "multi":
            if self.edit_column_name: # Should always be true in edit mode
                self.task_description_var.set(self.edit_column_name)
        
        # 填充字段处理方式 (after on_output_mode_change ensures multi_field_frame is visible if needed)
        field_mode = self.edit_config.get("field_mode", "predefined")
        self.field_mode_var.set(field_mode)
        # self.on_field_mode_change() # This will be called at the end

        # 填充输出字段 (after on_output_mode_change ensures multi_field_frame is visible if needed)
        self.output_fields.clear() # Clear any existing fields from previous UI state
        self.fields_listbox.delete(0, tk.END) # Clear listbox display
        output_fields_data = self.edit_config.get("output_fields", [])
        if output_fields_data:
            self.output_fields = output_fields_data.copy()
            for field in self.output_fields:
                self.fields_listbox.insert(tk.END, field)
        
        # 填充处理参数
        processing_params = self.edit_config.get("processing_params", {})
        self.max_workers_var.set(processing_params.get("max_workers", 3))
        self.request_delay_var.set(processing_params.get("request_delay", 0.5))
        self.max_retries_var.set(processing_params.get("max_retries", 2))
        
        # 最后更新依赖于字段模式的UI和JSON示例
        self.on_field_mode_change() 
        self.update_json_example()
        
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
                
    def parse_prompt_for_fields(self):
        """从Prompt中解析字段"""
        prompt_content = self.prompt_text.get("1.0", tk.END).strip()
        if not prompt_content:
            messagebox.showwarning("提示", "请先输入Prompt内容")
            return
        
        # 提取JSON相关的字段
        import re
        
        # 尝试匹配常见的JSON字段模式
        patterns = [
            r'"([^"]+)"\s*:', # 匹配 "field":
            r"'([^']+)'\s*:", # 匹配 'field':
            r'(\w+)\s*:', # 匹配 field:
            r'{\s*"([^"]+)"', # 匹配 {"field"
            r'{\s*\'([^\']+)\'', # 匹配 {'field'
        ]
        
        found_fields = set()
        for pattern in patterns:
            matches = re.findall(pattern, prompt_content, re.IGNORECASE)
            found_fields.update(matches)
        
        # 过滤掉一些常见的非字段词
        exclude_words = {
            'type', 'format', 'example', 'value', 'data', 'result', 'output',
            'json', 'response', 'answer', 'text', 'content', 'message'
        }
        found_fields = {f for f in found_fields if f.lower() not in exclude_words and len(f) > 1}
        
        if not found_fields:
            messagebox.showinfo("解析结果", "未在Prompt中找到明确的字段名\n\n请尝试在Prompt中明确指定输出字段，例如：\n'请输出JSON格式，包含以下字段：\"translation\": 翻译内容, \"confidence\": 置信度'")
            return
        
        # 显示解析结果给用户选择
        self.show_field_selection_dialog(found_fields)
    
    def show_field_selection_dialog(self, found_fields):
        """显示字段选择对话框"""
        selection_dialog = tk.Toplevel(self.dialog)
        selection_dialog.title("选择要添加的字段")
        selection_dialog.geometry("400x300")
        selection_dialog.transient(self.dialog)
        selection_dialog.grab_set()
        
        # 居中显示
        selection_dialog.update_idletasks()
        x = (selection_dialog.winfo_screenwidth() - selection_dialog.winfo_width()) // 2
        y = (selection_dialog.winfo_screenheight() - selection_dialog.winfo_height()) // 2
        selection_dialog.geometry(f"+{x}+{y}")
        
        frame = ttk.Frame(selection_dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="从Prompt中解析到以下字段，请选择要添加的：").pack(anchor=tk.W, pady=(0, 10))
        
        # 复选框列表
        field_vars = {}
        checkbox_frame = ttk.Frame(frame)
        checkbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        for field in sorted(found_fields):
            var = tk.BooleanVar(value=True)
            field_vars[field] = var
            ttk.Checkbutton(checkbox_frame, text=field, variable=var).pack(anchor=tk.W, pady=2)
        
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        def add_selected_fields():
            selected_fields = [field for field, var in field_vars.items() if var.get()]
            if selected_fields:
                added_fields = []
                for field in selected_fields:
                    if field not in self.output_fields:
                        self.output_fields.append(field)
                        self.fields_listbox.insert(tk.END, field)
                        added_fields.append(field)
                
                if added_fields:
                    self.update_json_example()
                    field_names = ", ".join(added_fields)
                    messagebox.showinfo("添加成功", f"已添加字段: {field_names}")
                else:
                    messagebox.showinfo("提示", "所有字段都已存在")
            
            selection_dialog.destroy()
        
        def select_all():
            for var in field_vars.values():
                var.set(True)
        
        def select_none():
            for var in field_vars.values():
                var.set(False)
        
        ttk.Button(button_frame, text="全选", command=select_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="全不选", command=select_none).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="添加选中", command=add_selected_fields).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=selection_dialog.destroy).pack(side=tk.RIGHT)
    
    def update_field_list_display(self):
        """更新字段列表显示"""
        self.fields_listbox.delete(0, tk.END)
        for field in self.output_fields:
            self.fields_listbox.insert(tk.END, field)
            
    def update_json_example(self):
        """更新JSON示例显示"""
        mode = self.field_mode_var.get()
        
        if mode == "auto_parse":
            example_text = """自动解析模式示例:
{
  "field1": "自动检测到的字段1",
  "field2": "自动检测到的字段2",
  "any_key": "任意字段名都会被解析"
}"""
        elif not self.output_fields:
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
        
        if not self.edit_mode:
            # 新建模式的验证
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
        else:
            # 编辑模式的验证
            if output_mode == "multi" and not self.output_fields:
                messagebox.showerror("错误", "多字段模式下请至少添加一个输出字段")
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
        
        # 获取列名
        if self.edit_mode:
            column_name = self.edit_column_name
        else:
            # 新建模式：单字段模式从输入框获取，多字段模式从任务描述获取
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
            'column_name': column_name,
            'model': model,
            'prompt': prompt,
            'processing_params': processing_params,
            'output_mode': output_mode,
            'output_fields': self.output_fields.copy() if output_mode == "multi" else [],
            'field_mode': self.field_mode_var.get() if output_mode == "multi" else "predefined",
            'edit_mode': self.edit_mode  # 添加编辑模式标识
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

    def on_field_mode_change(self):
        """字段处理方式改变时的处理"""
        mode = self.field_mode_var.get()
        
        if mode == "predefined":
            # 显示字段列表管理
            self.predefined_fields_frame.pack(fill=tk.X, pady=(0, 10))
        else:
            # 隐藏字段列表管理
            self.predefined_fields_frame.pack_forget()
            
        # 更新JSON示例
        self.update_json_example()
        
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