#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI Excel 批量数据处理工具
主程序入口
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.simpledialog
import pandas as pd
from table_manager import TableManager
from ai_processor import AIProcessor
from ai_column_dialog import AIColumnDialog
from project_manager import ProjectManager
import os
import time
import threading

class AIExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Excel Tool v2.0")
        self.root.geometry("1400x800")
        self.table_manager = TableManager()
        self.project_manager = ProjectManager()  
        self.ai_processor = AIProcessor()
        self.current_project_path = None
        self.highlighted_column = None
        self.last_sorted_column = None
        self.sort_ascending = True
        
        # 初始化行高设置（必须在setup_styles之前）
        self.row_height_settings = {
            '紧凑': 22,
            '标准': 25,
            '宽松': 30,
            '超宽松': 35
        }
        self.current_row_height = '标准'
        
        # 初始化筛选状态
        self.filter_state = {
            'active': False,
            'column': None,
            'selected_values': [],
            'all_values': [],
            'filtered_indices': []
        }
        
        # 初始化排序状态
        self.sort_state = {
            'column': None,
            'ascending': True,
            'original_order': None
        }
        
        # 添加撤销功能的历史记录
        self.undo_history = []
        self.max_undo_steps = 50  # 最大撤销步数
        
        # 创建样式
        self.setup_styles()
        
        # 创建菜单
        self.create_menu()
        
        # 创建工具栏
        self.create_toolbar()
        
        # 创建主界面
        self.create_main_frame()
        
        # 创建状态栏
        self.create_status_bar()
        
        # 绑定事件
        self.bind_events()
        
        # 显示欢迎页面
        self.show_welcome()
        
        # 设置窗口关闭处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 更新标题栏
        self.update_title()
        
    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        
        # 设置主题
        try:
            style.theme_use('clam')  # 使用更现代的主题
        except:
            pass
            
        # 现代化配色方案 - 使用更清晰的颜色
        
        # 配置主窗口背景
        self.root.configure(bg='#f8fafc')  # 浅灰白色背景
        
        # 自定义样式 - 使用更深的颜色提高对比度 和 更通用的字体
        style.configure('Title.TLabel', 
                       font=('Arial', 16, 'bold'), # 改为 Arial
                       foreground='#000000',  # 纯黑色文字，更清晰
                       background='#f8fafc')
                       
        style.configure('Subtitle.TLabel', 
                       font=('Arial', 11), # 改为 Arial
                       foreground='#2d3748',  # 深灰色文字，更清晰
                       background='#f8fafc')
                       
        style.configure('Success.TLabel', 
                       font=('Arial', 10), # 改为 Arial
                       foreground='#059669',  # 绿色
                       background='#f8fafc')
                       
        style.configure('Error.TLabel', 
                       font=('Arial', 10), # 改为 Arial
                       foreground='#dc2626',  # 红色
                       background='#f8fafc')
        
        # 现代化工具栏按钮样式
        style.configure('Toolbar.TButton', 
                       padding=(12, 8),
                       background='#ffffff',
                       foreground='#1a202c',  # 更深的文字颜色
                       borderwidth=1,
                       relief='solid',
                       bordercolor='#d1d5db',
                       font=('Arial', 10)) # 改为 Arial
                       
        style.map('Toolbar.TButton',
                 background=[('active', '#f3f4f6'),
                            ('pressed', '#e5e7eb')],
                 bordercolor=[('active', '#9ca3af'),
                             ('pressed', '#6b7280')])
        
        # 现代化LabelFrame样式
        style.configure('Modern.TLabelframe', 
                       background='#ffffff',
                       bordercolor='#e2e8f0',
                       borderwidth=1,
                       relief='solid')
                       
        style.configure('Modern.TLabelframe.Label',
                       background='#ffffff',
                       foreground='#000000',  # 纯黑色标题
                       font=('Arial', 11, 'bold'), # 改为 Arial
                       padding=(8, 4))
        
        # Frame样式
        style.configure('Modern.TFrame', background='#ffffff')
        
        # 其他组件样式
        style.configure('Modern.TLabel', 
                        font=('Arial', 10), # 改为 Arial
                        background='#ffffff', 
                        foreground='#1a202c')  # 更深的文字
        
        # 现代化Treeview样式 - 增强边框和网格线
        style.configure('Modern.Treeview',
                       background='#ffffff',
                       foreground='#1a202c',  # 更深的文字颜色
                       selectbackground='#dbeafe',  # 浅蓝色选中背景
                       selectforeground='#1e40af',  # 深蓝色选中文字
                       fieldbackground='#ffffff',
                       bordercolor='#9ca3af',  # 更深的边框颜色
                       borderwidth=1,  # 边框宽度
                       font=('Arial', 10), # 改为 Arial
                       rowheight=self.row_height_settings[self.current_row_height],  # 动态行高
                       relief='solid')  # 边框样式
                       
        style.configure('Modern.Treeview.Heading',
                       background='#f8fafc',  # 列头背景
                       foreground='#000000',  # 纯黑色列头文字
                       font=('Arial', 10, 'bold'), # 改为 Arial
                       bordercolor='#cbd5e1',  # 列头边框
                       borderwidth=1,  # 边框宽度
                       relief='solid',
                       padding=(8, 6))
        
        # 为AI列头定义特定样式
        style.configure('AI.Treeview.Heading',
                       background='#e3f2fd',  # 浅蓝色背景
                       foreground='#1a202c',  # 深色文字
                       font=('Arial', 10, 'bold'),
                       bordercolor='#90caf9',
                       borderwidth=1,
                       relief='solid',
                       padding=(8, 6))
                       
        # 为普通列头定义特定样式
        style.configure('Normal.Treeview.Heading',
                       background='#f8fafc',  # 浅灰白色背景
                       foreground='#000000',  # 纯黑色文字
                       font=('Arial', 10, 'bold'),
                       bordercolor='#cbd5e1',
                       borderwidth=1,
                       relief='solid',
                       padding=(8, 6))
        
        # 为长文本列头定义特定样式
        style.configure('LongText.Treeview.Heading',
                       background='#f0f4f8',  # Example: A slightly different light blue/gray
                       foreground='#000000',
                       font=('Arial', 10, 'bold'),
                       bordercolor='#cbd5e1',
                       borderwidth=1,
                       relief='solid',
                       padding=(8, 6))
        
        # 配置网格线和边框效果
        style.map('Modern.Treeview',
                 background=[('selected', '#dbeafe')],
                 foreground=[('selected', '#1e40af')])
        
        # 配置Treeview布局 - 添加网格线效果
        style.layout('Modern.Treeview', [
            ('Treeview.treearea', {'sticky': 'nswe'}),
            ('Treeview.border', {'sticky': 'nswe', 'border': '1', 'children': [
                ('Treeview.padding', {'sticky': 'nswe', 'children': [
                    ('Treeview.treearea', {'sticky': 'nswe'})
                ]})
            ]})
        ])
        
        # 配置列高亮专用样式
        style.configure('ColumnHighlight.Treeview',
                       background='#f0f8ff',  # 高亮背景色
                       foreground='#1a202c',
                       selectbackground='#bfdbfe',
                       selectforeground='#1e40af',
                       fieldbackground='#f0f8ff',
                       font=('Arial', 10),
                       rowheight=self.row_height_settings[self.current_row_height])
        
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建空白表格", command=self.create_blank_table, accelerator="Ctrl+N")
        file_menu.add_separator()
        
        # 项目文件操作
        file_menu.add_command(label="💾 保存项目", command=self.save_project, accelerator="Ctrl+S")
        file_menu.add_command(label="💾 另存为", command=self.save_project_as, accelerator="Ctrl+Shift+S")
        file_menu.add_command(label="📂 打开项目", command=self.load_project, accelerator="Ctrl+O")
        file_menu.add_separator()
        
        # 数据文件操作
        file_menu.add_command(label="导入Excel/CSV", command=self.import_data_file)
        file_menu.add_separator()
        
        # 导出子菜单
        export_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="📤 导出", menu=export_menu)
        
        # 快速导出
        export_menu.add_command(label="📊 快速导出Excel", command=self.export_excel)
        export_menu.add_command(label="📋 快速导出CSV", command=self.export_csv)
        export_menu.add_command(label="📄 快速导出JSONL", command=self.export_jsonl)
        export_menu.add_separator()
        
        # 选择性导出
        export_menu.add_command(label="🎯 选择字段导出", command=self.show_export_selection, accelerator="Ctrl+E")
        export_menu.add_command(label="🔍 条件筛选导出", command=self.show_conditional_export, accelerator="Ctrl+Alt+E")
        export_menu.add_command(label="⚡ 使用上次选择快速导出", command=self.quick_export_excel, accelerator="Ctrl+Shift+E")
        
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit, accelerator="Ctrl+Q")
        
        # 数据操作菜单（合并编辑和AI功能）
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="📊 数据操作", menu=data_menu)
        
        # 插入操作
        insert_menu = tk.Menu(data_menu, tearoff=0)
        data_menu.add_cascade(label="➕ 插入", menu=insert_menu)
        insert_menu.add_command(label="← 左侧插入列", command=lambda: self.insert_column_dialog("left"))
        insert_menu.add_command(label="→ 右侧插入列", command=lambda: self.insert_column_dialog("right"))
        insert_menu.add_separator()
        insert_menu.add_command(label="📄 创建长文本列", command=self.create_long_text_column)
        insert_menu.add_separator()
        insert_menu.add_command(label="⬇️ 添加行", command=self.add_row)
        
        data_menu.add_command(label="🗑️ 删除列", command=self.delete_column)
        data_menu.add_separator()
        
        # AI处理
        ai_submenu = tk.Menu(data_menu, tearoff=0)
        data_menu.add_cascade(label="🤖 AI处理", menu=ai_submenu)
        ai_submenu.add_command(label="🔄 全部处理", command=self.process_all_ai, accelerator="F5")
        ai_submenu.add_command(label="📋 单列处理", command=self.process_single_column, accelerator="F6")
        ai_submenu.add_command(label="⚡ 单元格处理", command=self.process_single_cell, accelerator="F7")
        ai_submenu.add_separator()
        ai_submenu.add_command(label="🔗 测试AI连接", command=self.test_ai_connection)
        
        data_menu.add_separator()
        data_menu.add_command(label="↶ 撤销", command=self.undo_action, accelerator="Ctrl+Z")
        data_menu.add_separator()
        data_menu.add_command(label="🔍 查找和替换", command=self.show_find_replace, accelerator="Ctrl+H")
        data_menu.add_command(label="🧹 清空所有数据", command=self.clear_data)
        
        # 视图菜单
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="👁️ 视图", menu=view_menu)
        
        # 排序操作
        sort_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="🔄 排序", menu=sort_submenu)
        sort_submenu.add_command(label="🔄 重置排序", command=self.reset_sort)
        sort_submenu.add_separator()
        sort_submenu.add_command(label="💡 右键列标题选择排序方式", state='disabled')
        
        # 筛选操作
        filter_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="🔍 筛选", menu=filter_submenu)
        filter_submenu.add_command(label="❌ 清除筛选", command=self.clear_filter)
        filter_submenu.add_separator()
        filter_submenu.add_command(label="💡 右键列标题选择筛选", state='disabled')
        
        view_menu.add_separator()
        
        # 行高设置
        self.row_height_var = tk.StringVar(value=self.current_row_height)
        row_height_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="📏 行高设置", menu=row_height_submenu)
        row_height_submenu.add_radiobutton(label="低 (紧凑)", variable=self.row_height_var, 
                                         value="low", command=self.change_row_height)
        row_height_submenu.add_radiobutton(label="中 (标准)", variable=self.row_height_var, 
                                         value="medium", command=self.change_row_height)
        row_height_submenu.add_radiobutton(label="高 (宽松)", variable=self.row_height_var, 
                                         value="high", command=self.change_row_height)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        
        # 绑定快捷键 - 使用bind_all确保全局生效
        self.root.bind_all('<Control-n>', lambda e: self.create_blank_table())
        self.root.bind_all('<Control-o>', lambda e: self.load_project())
        self.root.bind_all('<Control-s>', lambda e: self.save_project())
        self.root.bind_all('<Control-Shift-S>', lambda e: self.save_project_as())
        self.root.bind_all('<Control-e>', lambda e: self.show_export_selection())
        self.root.bind_all('<Control-Alt-e>', lambda e: self.show_conditional_export())
        self.root.bind_all('<Control-h>', lambda e: self.show_find_replace())
        self.root.bind_all('<Control-Shift-E>', lambda e: self.quick_export_excel())
        self.root.bind_all('<F5>', lambda e: self.process_all_ai())
        self.root.bind_all('<F6>', lambda e: self.process_single_column())
        self.root.bind_all('<F7>', lambda e: self.process_single_cell())
        
    def create_toolbar(self):
        """创建工具栏区域 - 现在用于预览面板"""
        # 工具栏容器 - 现在用作预览面板容器
        self.toolbar_container = ttk.Frame(self.root, style='Modern.TFrame')
        self.toolbar_container.pack(side=tk.TOP, fill=tk.X, padx=20, pady=(20, 0))
        
        # 创建内容预览面板
        self.create_content_preview_panel(self.toolbar_container)
        
        # 初始化导出选择状态
        self.export_selection = {
            'selected_columns': [],
            'remember_selection': True
        }
        
        # 在状态栏创建简化的状态信息
        self.info_label = None  # 将在状态栏中创建
        
    def create_main_frame(self):
        """创建主界面"""
        # 主容器 - 现代化样式
        main_container = ttk.Frame(self.root, style='Modern.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 创建水平分割的主要区域
        content_container = ttk.Frame(main_container, style='Modern.TFrame')
        content_container.pack(fill=tk.BOTH, expand=True)
        
        # 左侧表格区域
        left_panel = ttk.Frame(content_container, style='Modern.TFrame')
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 表格区域 - 更大的容器
        self.table_frame = ttk.LabelFrame(left_panel, text="📊 数据表格", style='Modern.TLabelframe', padding=16)
        self.table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 预览面板将放在工具栏位置，这里不创建
        
        # 表格标题栏（包含进度条）
        self.table_header_frame = ttk.Frame(self.table_frame, style='Modern.TFrame')
        self.table_header_frame.pack(fill=tk.X, pady=(0, 8))
        
        # 进度信息
        self.progress_info_frame = ttk.Frame(self.table_header_frame, style='Modern.TFrame')
        self.progress_info_frame.pack(side=tk.RIGHT)
        
        self.progress_label = ttk.Label(self.progress_info_frame, text="", style='Modern.TLabel', font=('Arial', 9))
        self.progress_label.pack(side=tk.RIGHT, padx=(0, 8))
        
        self.table_progress_bar = ttk.Progressbar(self.progress_info_frame, mode='determinate', length=150)
        self.table_progress_bar.pack(side=tk.RIGHT)
        
        # 创建表格容器 - 增加内边距和边框
        table_container = ttk.Frame(self.table_frame, style='Modern.TFrame')
        table_container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        
        # 创建带边框的表格外层容器
        table_border_frame = tk.Frame(table_container, 
                                     bg='#cbd5e1',  # 边框颜色
                                     bd=1, 
                                     relief='solid')
        table_border_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格内部容器
        table_inner_frame = tk.Frame(table_border_frame, bg='#ffffff')
        table_inner_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # 创建表格和滚动条的容器
        table_scroll_frame = ttk.Frame(table_inner_frame)
        table_scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格 - 使用现代化样式
        self.tree = ttk.Treeview(table_scroll_frame, show='headings', style='Modern.Treeview')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置表格的网格线和边框
        self.tree.configure(
            selectmode='browse',  # 单选模式
            show='headings'  # 只显示标题行
        )
        
        # 设置表格的外边框 - 移除不支持的选项
        # self.tree.configure(relief='solid', borderwidth=1)  # Treeview不支持这些选项
        
        # 现代化垂直滚动条
        v_scrollbar = ttk.Scrollbar(table_scroll_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # 现代化水平滚动条
        h_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # 绑定事件 - 在这里就绑定，确保始终有效
        self.bind_tree_events()
        
        # 欢迎界面
        self.welcome_frame = ttk.Frame(self.table_frame, style='Modern.TFrame')
        
    def create_content_preview_panel(self, parent):
        """创建内容预览面板 - 水平布局"""
        # 预览面板容器
        self.preview_panel = ttk.LabelFrame(parent, text="📖 内容预览", 
                                          style='Modern.TLabelframe', padding=16)
        # 初始时不显示，将在有数据时显示
        
        # 水平布局：左侧信息，右侧内容，最右侧按钮
        left_info_frame = ttk.Frame(self.preview_panel, style='Modern.TFrame')
        left_info_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        
        # 单元格信息显示
        ttk.Label(left_info_frame, text="位置:", style='Modern.TLabel', font=('Arial', 9)).pack(anchor='w')
        self.cell_info_label = ttk.Label(left_info_frame, text="未选中", 
                                        style='Subtitle.TLabel', font=('Arial', 10, 'bold'))
        self.cell_info_label.pack(anchor='w', pady=(2, 8))
        
        ttk.Label(left_info_frame, text="类型:", style='Modern.TLabel', font=('Arial', 9)).pack(anchor='w')
        self.cell_type_label = ttk.Label(left_info_frame, text="", 
                                        style='Modern.TLabel', font=('Arial', 9))
        self.cell_type_label.pack(anchor='w')
        
        # 内容显示区域 - 水平展开
        content_frame = ttk.Frame(self.preview_panel, style='Modern.TFrame')
        content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        ttk.Label(content_frame, text="内容:", style='Modern.TLabel', font=('Arial', 9)).pack(anchor='w')
        
        # 文本显示控件 - 调整为水平布局
        text_container = ttk.Frame(content_frame)
        text_container.pack(fill=tk.BOTH, expand=True, pady=(2, 0))
        
        self.preview_text = tk.Text(text_container, wrap=tk.WORD, height=4, width=60,
                                   font=('Arial', 10), relief='solid', borderwidth=1)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        preview_scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, 
                                         command=self.preview_text.yview)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        # AI Prompt 预览区域
        self.prompt_preview_frame = ttk.LabelFrame(content_frame, text="💬 AI请求Prompt", 
                                                   style='Modern.TLabelframe', padding="10")
        self.prompt_preview_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Prompt文本框和开关
        prompt_text_container = ttk.Frame(self.prompt_preview_frame)
        prompt_text_container.pack(fill=tk.BOTH, expand=True)
        
        self.prompt_text = tk.Text(prompt_text_container, wrap=tk.WORD, height=3, width=60,
                                  font=('Arial', 10), relief='solid', borderwidth=1)
        self.prompt_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        prompt_scrollbar = ttk.Scrollbar(prompt_text_container, orient=tk.VERTICAL, 
                                         command=self.prompt_text.yview)
        prompt_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.prompt_text.configure(yscrollcommand=prompt_scrollbar.set)
        
        # 一键开关 - 用于显示/隐藏完整Prompt
        self.show_full_prompt_var = tk.BooleanVar(value=False)
        self.show_full_prompt_checkbox = ttk.Checkbutton(self.prompt_preview_frame, 
                                                        text="显示完整Prompt", 
                                                        variable=self.show_full_prompt_var, 
                                                        command=self.toggle_full_prompt_display)
        self.show_full_prompt_checkbox.pack(anchor=tk.E, pady=(5,0))
        
        # 初始状态隐藏Prompt预览
        self.prompt_preview_frame.pack_forget()
        
        # 操作按钮 - 垂直排列在右侧
        button_frame = ttk.Frame(self.preview_panel, style='Modern.TFrame')
        button_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        ttk.Label(button_frame, text="操作:", style='Modern.TLabel', font=('Arial', 9)).pack(anchor='w')
        
        # 按钮垂直排列
        self.edit_button = ttk.Button(button_frame, text="✏️编辑", 
                                     command=self.edit_from_preview, state='disabled', width=8)
        self.edit_button.pack(pady=(2, 3))
        
        self.copy_button = ttk.Button(button_frame, text="📋 复制", 
                                     command=self.copy_from_preview, state='disabled', width=8)
        self.copy_button.pack(pady=3)
        
        self.clear_button = ttk.Button(button_frame, text="🗑️清空", 
                                      command=self.clear_preview, state='disabled', width=8)
        self.clear_button.pack(pady=3)
        
        # 初始状态显示
        self.preview_text.insert("1.0", "选择一个单元格来查看其内容...")
        self.preview_text.config(state='disabled')
        
        # 存储当前选中的单元格信息
        self.current_preview_cell = None
    
    def update_content_preview(self, row_index, col_name, content):
        """更新内容预览"""
        try:
            # 更新单元格信息
            self.cell_info_label.config(text=f"{col_name} [第{row_index+1}行]")
            
            # 获取列类型信息
            if self.table_manager.is_long_text_column(col_name):
                self.cell_type_label.config(text="长文本列", foreground="green") # Example color
            elif col_name in self.table_manager.get_ai_columns(): # Check AI after Long Text
                ai_columns = self.table_manager.get_ai_columns()
                config = ai_columns[col_name]
                if isinstance(config, dict):
                    model = config.get("model", "gpt-4.1")
                    self.cell_type_label.config(text=f"AI列 ({model})", foreground="blue")
                    
                    # 显示Prompt预览区域
                    self.prompt_preview_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                    
                    # 构建并显示完整Prompt
                    full_prompt = self.ai_processor.build_full_prompt(row_index, col_name, self.table_manager)
                    self.prompt_text.config(state='normal')
                    self.prompt_text.delete("1.0", tk.END)
                    self.prompt_text.insert("1.0", full_prompt)
                    self.prompt_text.config(state='disabled')
                    
                    # 重置开关状态，默认不显示完整prompt，但框显示
                    self.show_full_prompt_var.set(False)
                    self.toggle_full_prompt_display() # 根据初始值设置显示状态
                    self.show_full_prompt_checkbox.config(state='normal')

                else:
                    self.cell_type_label.config(text="AI列 (gpt-4.1)", foreground="blue")
                    # 如果是旧格式，也显示Prompt预览，但只显示prompt template
                    self.prompt_preview_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                    self.prompt_text.config(state='normal')
                    self.prompt_text.delete("1.0", tk.END)
                    self.prompt_text.insert("1.0", self.table_manager.get_ai_column_prompt(col_name))
                    self.prompt_text.config(state='disabled')
                    self.show_full_prompt_var.set(False)
                    self.toggle_full_prompt_display()
                    self.show_full_prompt_checkbox.config(state='normal')

            else:
                self.cell_type_label.config(text="普通列", foreground="gray")
                # 隐藏Prompt预览区域
                self.prompt_preview_frame.pack_forget()
                self.show_full_prompt_checkbox.config(state='disabled')
                self.prompt_text.config(state='normal')
                self.prompt_text.delete("1.0", tk.END)
                self.prompt_text.config(state='disabled')

            # 更新内容显示
            self.preview_text.config(state='normal')
            self.preview_text.delete("1.0", tk.END)
            
            if content and str(content).strip():
                self.preview_text.insert("1.0", str(content))
                # 启用操作按钮
                self.edit_button.config(state='normal')
                self.copy_button.config(state='normal')
                self.clear_button.config(state='normal')
            else:
                self.preview_text.insert("1.0", "[空内容]")
                # 启用编辑按钮，禁用复制按钮
                self.edit_button.config(state='normal')
                self.copy_button.config(state='disabled')
                self.clear_button.config(state='disabled')
            
            self.preview_text.config(state='disabled')
            
            # 保存当前选中信息
            self.current_preview_cell = {
                'row_index': row_index,
                'col_name': col_name,
                'content': content
            }
            
        except Exception as e:
            print(f"更新内容预览失败: {e}")
    
    def clear_content_preview(self):
        """清空内容预览"""
        self.cell_info_label.config(text="未选中单元格")
        self.cell_type_label.config(text="")
        
        self.preview_text.config(state='normal')
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", "选择一个单元格来查看其内容...")
        self.preview_text.config(state='disabled')
        
        # 禁用操作按钮
        self.edit_button.config(state='disabled')
        self.copy_button.config(state='disabled')
        self.clear_button.config(state='disabled')
        
        self.current_preview_cell = None
    
    def edit_from_preview(self):
        """从预览面板编辑内容"""
        if self.current_preview_cell:
            row_index = self.current_preview_cell['row_index']
            col_name = self.current_preview_cell['col_name']
            content = self.current_preview_cell['content']
            self.edit_cell_dialog(row_index, col_name, str(content) if content else "")
    
    def copy_from_preview(self):
        """从预览面板复制内容"""
        if self.current_preview_cell and self.current_preview_cell['content']:
            content = str(self.current_preview_cell['content'])
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.update_status("内容已复制到剪贴板", "success")
    
    def clear_preview(self):
        """清空当前单元格内容"""
        if self.current_preview_cell:
            result = messagebox.askyesno("确认清空", "确定要清空当前单元格的内容吗？")
            if result:
                row_index = self.current_preview_cell['row_index']
                col_name = self.current_preview_cell['col_name']
                
                # 更新数据框
                df = self.table_manager.get_dataframe()
                if df is not None:
                    df.iloc[row_index, df.columns.get_loc(col_name)] = ""
                    
                    # 刷新显示
                    self.update_table_display()
                    self.update_title()  # 更新标题栏
                    self.update_content_preview(row_index, col_name, "")
                    self.update_status(f"已清空 {col_name} [第{row_index+1}行]", "success")
        
    def bind_tree_events(self):
        """绑定表格事件"""
        # 右键菜单
        self.tree.bind("<Button-3>", self.show_context_menu)
        # 双击编辑
        self.tree.bind("<Double-1>", self.on_cell_double_click)
        # 单击选择 - 改进选中逻辑
        self.tree.bind("<Button-1>", self.on_cell_click)
        
        # 列拖拽事件
        self.tree.bind("<Button-1>", self.on_column_drag_start, add='+')
        self.tree.bind("<B1-Motion>", self.on_column_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_column_drag_end)
        
        # 全局快捷键绑定
        self.root.bind_all("<Control-s>", lambda e: self.save_project())
        self.root.bind_all("<Control-S>", lambda e: self.save_project())
        self.root.bind_all("<Control-z>", lambda e: self.undo_action())
        self.root.bind_all("<Control-Z>", lambda e: self.undo_action())
        
        # 添加选中状态追踪
        self.selection_info = {
            'type': None,  # 'cell', 'column', 'row' 
            'row_index': None,
            'column_index': None,
            'column_name': None
        }
        
        # 列高亮状态
        self.highlighted_column = None
        
        # 拖拽状态变量
        self.drag_data = {
            'dragging': False,
            'start_column': None,
            'start_x': 0,
            'target_column': None,
            'drag_indicator': None
        }
        
    def show_welcome(self):
        """显示欢迎界面"""
        # 隐藏表格
        for widget in self.table_frame.winfo_children():
            widget.pack_forget()
            
        # 隐藏预览面板
        self.preview_panel.pack_forget()
            
        # 显示欢迎界面
        self.welcome_frame.pack(fill=tk.BOTH, expand=True)
        
        # 欢迎内容容器
        welcome_container = ttk.Frame(self.welcome_frame, style='Modern.TFrame')
        welcome_container.place(relx=0.5, rely=0.5, anchor='center')
        
        # 主标题 - 简化版本
        title_frame = ttk.Frame(welcome_container, style='Modern.TFrame')
        title_frame.pack(pady=(0, 30))
        
        ttk.Label(title_frame, text="AI批量数据处理工具", 
                 style='Title.TLabel', font=('Arial', 18, 'bold')).pack(pady=(8, 0))
        
        # 操作按钮区域
        action_frame = ttk.Frame(welcome_container, style='Modern.TFrame')
        action_frame.pack(pady=(0, 30))
        
        # 创建现代化的大按钮
        btn_frame = ttk.Frame(action_frame, style='Modern.TFrame')
        btn_frame.pack()
        
        # 新建按钮
        new_btn = ttk.Button(btn_frame, text="📄 新建空白表格", command=self.create_blank_table,
                            style='Toolbar.TButton', width=18)
        new_btn.pack(side=tk.LEFT, padx=8)
        
        # 打开项目按钮
        open_btn = ttk.Button(btn_frame, text="📂 打开项目文件", command=self.load_project,
                             style='Toolbar.TButton', width=18)
        open_btn.pack(side=tk.LEFT, padx=8)
        
        # 导入数据按钮
        import_btn = ttk.Button(btn_frame, text="📥 导入数据文件", command=self.import_data_file,
                               style='Toolbar.TButton', width=18)
        import_btn.pack(side=tk.LEFT, padx=8)
        

        
    def hide_welcome(self):
        """隐藏欢迎界面，显示表格"""
        print("隐藏欢迎界面，显示表格")
        self.welcome_frame.pack_forget()
        
        # 显示预览面板在工具栏区域
        self.preview_panel.pack(fill=tk.X, pady=(0, 10))
        print(f"预览面板已显示在工具栏区域")
        
        # 确保表格组件可见
        for widget in self.table_frame.winfo_children():
            if widget != self.welcome_frame:
                widget.pack_configure()
        
        # 确保事件绑定有效
        self.bind_tree_events()
        
        # 强制更新界面
        self.root.update_idletasks()
        
    def create_blank_table(self):
        """创建空白表格"""
        # 检查未保存的更改
        if not self.check_unsaved_changes_before_action("创建新的空白表格"):
            return
            
        success = self.table_manager.create_blank_table()
        if success:
            # 清除当前项目路径
            self.current_project_path = None
            self.hide_welcome()
            self.update_table_display()
            self.info_label.config(text="📄 新建表格")
            self.update_status("已创建空白表格", "success")
            self.update_title()  # 更新标题栏
        else:
            messagebox.showerror("错误", "创建空白表格失败")
            
    def create_status_bar(self):
        """创建状态栏"""
        self.status_frame = ttk.Frame(self.root, style='Modern.TFrame', padding=12)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 20))
        
        # 状态栏内容容器
        status_content = ttk.Frame(self.status_frame, style='Modern.TFrame')
        status_content.pack(fill=tk.X)
        
        # 左侧状态
        left_status = ttk.Frame(status_content, style='Modern.TFrame')
        left_status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        status_icon = ttk.Label(left_status, text="📊", style='Modern.TLabel', font=('Arial', 12))
        status_icon.pack(side=tk.LEFT, padx=(0, 8))
        
        self.status_label = ttk.Label(left_status, text="就绪", style='Modern.TLabel', 
                                     font=('Arial', 10))
        self.status_label.pack(side=tk.LEFT)
        
        # 添加文件信息标签
        ttk.Label(left_status, text=" | 文件:", style='Modern.TLabel', font=('Arial', 10)).pack(side=tk.LEFT, padx=(20, 0))
        self.info_label = ttk.Label(left_status, text="无", style='Subtitle.TLabel', 
                                   font=('Arial', 10, 'bold'))
        self.info_label.pack(side=tk.LEFT, padx=(8, 0))
        
        # 右侧进度
        right_status = ttk.Frame(status_content, style='Modern.TFrame')
        right_status.pack(side=tk.RIGHT)
        
        progress_label = ttk.Label(right_status, text="进度:", style='Modern.TLabel', 
                                  font=('Arial', 10))
        progress_label.pack(side=tk.RIGHT, padx=(0, 8))
        
        self.progress_bar = ttk.Progressbar(right_status, mode='determinate', length=200)
        self.progress_bar.pack(side=tk.RIGHT, pady=2)
        
    def update_status(self, message, status_type="normal"):
        """更新状态信息"""
        self.status_label.config(text=message)
        if status_type == "success":
            self.status_label.config(style='Success.TLabel')
        elif status_type == "error":
            self.status_label.config(style='Error.TLabel')
        else:
            self.status_label.config(style='TLabel')
            
    def bind_events(self):
        """绑定事件"""
        pass  # 事件绑定在hide_welcome中处理
        
    def show_context_menu(self, event):
        """显示右键菜单 - 基于选中状态的智能菜单"""
        context_menu = tk.Menu(self.root, tearoff=0)
        
        # 根据选中状态显示不同菜单
        if self.selection_info['type'] == 'column':
            # 选中列的菜单
            col_name = self.selection_info['column_name']
            ai_columns = self.table_manager.get_ai_columns()
            is_ai_col = col_name in ai_columns
            col_type = "AI列" if is_ai_col else "普通列"
            
            # 列标题
            context_menu.add_command(label=f"📊 {col_type}: {col_name}", state='disabled')
            context_menu.add_separator()
            
            # 列编辑操作
            context_menu.add_command(
                label="✏️ 编辑列名",
                command=lambda: self.edit_column_name(col_name)
            )
            
            if is_ai_col:
                context_menu.add_command(
                    label="🤖 编辑AI功能",
                    command=lambda: self.edit_ai_prompt(col_name)
                )
                context_menu.add_separator()
                
                # AI处理操作
                context_menu.add_command(
                    label="📊 查看处理进度",
                    command=lambda: self.show_ai_column_progress(col_name)
                )
                context_menu.add_command(
                    label="⚡ AI处理整列(新版)",
                    command=lambda: self.process_ai_column_concurrent(col_name)
                )
                context_menu.add_command(
                    label="⚡ AI处理整列(旧版)",
                    command=lambda: self.process_entire_column(col_name)
                )
                context_menu.add_separator()
                
                context_menu.add_command(
                    label="📝 转换为普通列",
                    command=lambda: self.convert_to_normal_column(col_name)
                )
            
            context_menu.add_separator()
            
            # 筛选操作
            context_menu.add_command(
                label="🔍 筛选数据",
                command=lambda: self.show_filter_dialog(col_name)
            )
            if self.filter_state['active'] and self.filter_state['column'] == col_name:
                context_menu.add_command(
                    label="❌ 清除筛选",
                    command=self.clear_filter
                )
            
            context_menu.add_separator()
            
            # 排序操作
            sort_submenu = tk.Menu(context_menu, tearoff=0)
            context_menu.add_cascade(label="🔄 排序", menu=sort_submenu)
            sort_submenu.add_command(
                label="↑ 升序排序",
                command=lambda: self.sort_by_column(col_name, ascending=True)
            )
            sort_submenu.add_command(
                label="↓ 降序排序", 
                command=lambda: self.sort_by_column(col_name, ascending=False)
            )
            if self.sort_state['column'] is not None:
                sort_submenu.add_separator()
                sort_submenu.add_command(
                    label="🔄 重置排序",
                    command=self.reset_sort
                )
            
            # 列操作
            col_index = self.selection_info['column_index']
            context_menu.add_command(
                label="➕ 左侧插入列",
                command=lambda: self.insert_column_at_position(col_index, "left")
            )
            context_menu.add_command(
                label="➕ 右侧插入列", 
                command=lambda: self.insert_column_at_position(col_index + 1, "right")
            )
            context_menu.add_command(
                label="🗑️ 删除此列",
                command=lambda: self.delete_specific_column(col_name)
            )
            
        elif self.selection_info['type'] == 'cell':
            # 选中单元格的菜单
            row_index = self.selection_info['row_index']
            col_name = self.selection_info['column_name']
            ai_columns = self.table_manager.get_ai_columns()
            is_ai_cell = col_name in ai_columns
            cell_type = "AI单元格" if is_ai_cell else "单元格"
            
            # 单元格标题
            context_menu.add_command(
                label=f"📍 {cell_type}: {col_name}[第{row_index+1}行]", 
                state='disabled'
            )
            context_menu.add_separator()
            
            # 单元格编辑
            context_menu.add_command(
                label="✏️ 编辑内容",
                command=lambda: self.edit_specific_cell(row_index, col_name)
            )
            
            if is_ai_cell:
                context_menu.add_separator()
                # AI处理操作
                context_menu.add_command(
                    label="🤖 AI处理此单元格",
                    command=lambda: self.process_specific_cell(row_index, col_name)
                )
                context_menu.add_command(
                    label="⚡ AI处理此行所有AI列",
                    command=lambda: self.process_selected_row(row_index)
                )
            
            context_menu.add_separator()
            
            # 行操作
            context_menu.add_command(
                label="⬆️ 上方插入行",
                command=lambda: self.insert_row_at_position(row_index, "above")
            )
            context_menu.add_command(
                label="⬇️ 下方插入行",
                command=lambda: self.insert_row_at_position(row_index + 1, "below")
            )
            context_menu.add_command(
                label=f"🗑️ 删除第{row_index+1}行",
                command=lambda: self.delete_selected_row(row_index)
            )
            
        else:
            # 未选中或空白区域的菜单
            context_menu.add_command(label="📊 表格操作", state='disabled')
            context_menu.add_separator()
        
        # 通用操作
        if self.selection_info['type'] is None:  # 只在空白区域显示
            context_menu.add_separator()
            context_menu.add_command(label="➕ 添加行", command=self.add_row)
            
            # 排序操作
            if self.sort_state['column'] is not None:
                context_menu.add_separator()
                context_menu.add_command(
                    label="🔄 重置排序",
                    command=self.reset_sort
                )
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
            
    def edit_column_name(self, old_name):
        """编辑列名"""
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑列名")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"400x200+{x}+{y}")
        
        ttk.Label(dialog, text=f"编辑列名", style='Title.TLabel').pack(pady=10)
        
        # 当前列名显示
        current_frame = ttk.LabelFrame(dialog, text="当前列名", padding=5)
        current_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(current_frame, text=old_name, style='Subtitle.TLabel').pack()
        
        # 新列名输入
        new_frame = ttk.LabelFrame(dialog, text="新列名", padding=5)
        new_frame.pack(fill=tk.X, padx=10, pady=5)
        
        entry_var = tk.StringVar(value=old_name)
        entry = ttk.Entry(new_frame, textvariable=entry_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        entry.select_range(0, tk.END)
        
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_save():
            new_name = entry_var.get().strip()
            if not new_name:
                messagebox.showwarning("警告", "列名不能为空")
                return
                
            if new_name == old_name:
                dialog.destroy()
                return
                
            # 检查新列名是否已存在
            existing_cols = self.table_manager.get_column_names()
            if new_name in existing_cols:
                messagebox.showerror("错误", f"列名 '{new_name}' 已存在")
                return
                
            # 执行重命名
            success = self.table_manager.rename_column(old_name, new_name)
            if success:
                self.update_table_display()
                self.update_title()  # 更新标题栏
                self.update_status(f"列名已更改: {old_name} → {new_name}", "success")
                messagebox.showinfo("成功", f"列名已更改为: {new_name}")
                dialog.destroy()
            else:
                messagebox.showerror("错误", "重命名失败")
                
        def on_cancel():
            dialog.destroy()
            
        ttk.Button(button_frame, text="💾 保存", command=on_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="❌ 取消", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        dialog.bind('<Return>', lambda e: on_save())
        dialog.bind('<Escape>', lambda e: on_cancel())
        
    def edit_ai_prompt(self, col_name):
        """编辑AI提示词 - 使用和新建相同的界面"""
        ai_columns = self.table_manager.get_ai_columns()
        if col_name not in ai_columns:
            messagebox.showwarning("警告", f"'{col_name}' 不是AI列")
            return
            
        # 获取当前配置
        config = ai_columns[col_name]
        if isinstance(config, dict):
            current_prompt = config.get("prompt", "")
            current_model = config.get("model", "gpt-4.1")
            current_params = config.get("processing_params", {
                'max_workers': 3,
                'request_delay': 0.5,
                'max_retries': 2
            })
        else:
            # 向后兼容旧格式
            current_prompt = config
            current_model = "gpt-4.1"
            current_params = {
                'max_workers': 3,
                'request_delay': 0.5,
                'max_retries': 2
            }
        
        # 使用 AI 列对话框的相似设计，但预填充现有数据
        from ai_column_dialog import AIColumnDialog
        
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑AI列配置 - {col_name}")
        dialog.geometry("700x700")  # 增加高度以容纳处理参数
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (700 // 2)
        y = (dialog.winfo_screenheight() // 2) - (700 // 2)
        dialog.geometry(f"700x700+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 列名显示（不可编辑）
        ttk.Label(main_frame, text="列名:").pack(anchor=tk.W, pady=(0, 5))
        column_name_display = ttk.Label(main_frame, text=col_name, style='Subtitle.TLabel', 
                                       background='#f8f9fa', relief='solid', padding=5)
        column_name_display.pack(fill=tk.X, pady=(0, 10))
        
        # AI模型选择
        model_config_frame = ttk.Frame(main_frame)
        model_config_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(model_config_frame, text="AI模型:").pack(side=tk.LEFT, padx=(0, 10))
        model_var = tk.StringVar(value=current_model)
        model_combo = ttk.Combobox(model_config_frame, textvariable=model_var, 
                                  values=["gpt-4.1", "o1"], state="readonly", width=15)
        model_combo.pack(side=tk.LEFT)
        
        # 模型说明
        ttk.Label(model_config_frame, text="  (gpt-4.1: 快速响应 | o1: 深度推理)", 
                 foreground="gray", font=('Arial', 8)).pack(side=tk.LEFT, padx=(10, 0))
        
        # 处理参数配置框架
        params_frame = ttk.LabelFrame(main_frame, text="处理参数配置", padding="10")
        params_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 并发数设置
        concurrent_frame = ttk.Frame(params_frame)
        concurrent_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(concurrent_frame, text="并发数:").pack(side=tk.LEFT, padx=(0, 10))
        max_workers_var = tk.IntVar(value=current_params.get('max_workers', 3))
        max_workers_spinbox = ttk.Spinbox(concurrent_frame, from_=1, to=10, 
                                         textvariable=max_workers_var, width=5)
        max_workers_spinbox.pack(side=tk.LEFT)
        ttk.Label(concurrent_frame, text="  (同时处理的任务数，建议1-5)", 
                 foreground="gray", font=('Arial', 8)).pack(side=tk.LEFT, padx=(5, 0))
        
        # 请求延迟设置
        delay_frame = ttk.Frame(params_frame)
        delay_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(delay_frame, text="请求延迟:").pack(side=tk.LEFT, padx=(0, 10))
        request_delay_var = tk.DoubleVar(value=current_params.get('request_delay', 0.5))
        delay_spinbox = ttk.Spinbox(delay_frame, from_=0.1, to=5.0, increment=0.1,
                                   textvariable=request_delay_var, width=5)
        delay_spinbox.pack(side=tk.LEFT)
        ttk.Label(delay_frame, text="秒  (避免API限流，建议0.3-1.0)", 
                 foreground="gray", font=('Arial', 8)).pack(side=tk.LEFT, padx=(5, 0))
        
        # 重试次数设置
        retry_frame = ttk.Frame(params_frame)
        retry_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(retry_frame, text="重试次数:").pack(side=tk.LEFT, padx=(0, 10))
        max_retries_var = tk.IntVar(value=current_params.get('max_retries', 2))
        retry_spinbox = ttk.Spinbox(retry_frame, from_=0, to=5, 
                                   textvariable=max_retries_var, width=5)
        retry_spinbox.pack(side=tk.LEFT)
        ttk.Label(retry_frame, text="  (API失败时的重试次数，建议1-3)", 
                 foreground="gray", font=('Arial', 8)).pack(side=tk.LEFT, padx=(5, 0))
        
        # Prompt模板输入区域
        prompt_frame = ttk.LabelFrame(main_frame, text="AI Prompt模板", padding="10")
        prompt_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 提示信息
        tip_text = "在prompt中使用 {列名} 来引用其他字段的值\n例如: 请将以下{category}类的英文query翻译成中文：{query}"
        ttk.Label(prompt_frame, text=tip_text, foreground="gray").pack(anchor=tk.W, pady=(0, 5))
        
        # 可用字段显示
        existing_columns = [col for col in self.table_manager.get_column_names() if col != col_name]
        if existing_columns:
            fields_label = ttk.Label(prompt_frame, text="可用字段: (双击复制)", foreground="red")
            fields_label.pack(anchor=tk.W, pady=(0, 5))
            
            # 创建可选择的字段文本框
            fields_frame = ttk.Frame(prompt_frame)
            fields_frame.pack(fill=tk.X, pady=(0, 10))
            
            # 字段文本框 - 只读但可选择复制
            fields_text = tk.Text(fields_frame, height=3, wrap=tk.WORD, 
                                 background='#f8f9fa', relief='solid', borderwidth=1,
                                 font=('Consolas', 9))
            fields_text.pack(fill=tk.X)
            
            # 填充字段内容
            fields_content = ""
            fields_list = [f"{{{col}}}" for col in existing_columns]
            
            # 按行排列字段，每行最多4个
            for i in range(0, len(fields_list), 4):
                line_fields = fields_list[i:i+4]
                fields_content += "  ".join(line_fields) + "\n"
            
            fields_text.insert("1.0", fields_content.strip())
            fields_text.config(state=tk.DISABLED)
            
            # 添加双击复制功能
            def on_field_double_click(event):
                try:
                    fields_text.config(state=tk.NORMAL)
                    # 获取点击位置的字符
                    index = fields_text.index(f"@{event.x},{event.y}")
                    # 获取当前行
                    line_start = fields_text.index(f"{index} linestart")
                    line_end = fields_text.index(f"{index} lineend")
                    line_text = fields_text.get(line_start, line_end)
                    
                    # 找到点击的字段
                    import re
                    fields_in_line = re.findall(r'\{[^}]+\}', line_text)
                    if fields_in_line:
                        # 简单选择第一个字段（或者可以改进为选择最接近的）
                        selected_field = fields_in_line[0]
                        dialog.clipboard_clear()
                        dialog.clipboard_append(selected_field)
                        messagebox.showinfo("复制成功", f"已复制字段: {selected_field}")
                    fields_text.config(state=tk.DISABLED)
                except:
                    fields_text.config(state=tk.DISABLED)
            
            fields_text.bind("<Double-Button-1>", on_field_double_click)
            
            # 提示标签
            tip_label = ttk.Label(fields_frame, text="💡 双击字段名可快速复制到剪贴板", 
                                foreground="gray", font=('Arial', 8))
            tip_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Prompt文本框
        prompt_text = tk.Text(prompt_frame, height=8, wrap=tk.WORD, width=80)
        prompt_text.pack(fill=tk.BOTH, expand=True)
        prompt_text.insert("1.0", current_prompt)
        prompt_text.focus()
        
        # 滚动条
        scrollbar = ttk.Scrollbar(prompt_frame, orient=tk.VERTICAL, command=prompt_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        prompt_text.configure(yscrollcommand=scrollbar.set)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def on_save():
            new_prompt = prompt_text.get("1.0", tk.END).strip()
            if not new_prompt:
                messagebox.showwarning("警告", "提示词不能为空")
                return
                
            # 验证提示词模板
            is_valid, message = self.table_manager.validate_prompt_template(new_prompt)
            if not is_valid:
                result = messagebox.askyesno("模板验证", 
                                           f"提示词模板可能有问题：{message}\n\n是否仍要保存？")
                if not result:
                    return
            
            # 获取处理参数
            processing_params = {
                'max_workers': max_workers_var.get(),
                'request_delay': request_delay_var.get(),
                'max_retries': max_retries_var.get()
            }
                    
            # 更新AI列配置（包含模型信息和处理参数）
            new_model = model_var.get()
            self.table_manager.update_ai_column_config(col_name, new_prompt, new_model, processing_params)
            self.update_title()  # 更新标题栏
            self.update_status(f"已更新AI列配置: {col_name} (模型: {new_model})", "success")
            messagebox.showinfo("成功", f"AI列配置已更新\n模型: {new_model}\n并发数: {processing_params['max_workers']}")
            dialog.destroy()
                
        def on_cancel():
            dialog.destroy()
            
        ttk.Button(button_frame, text="保存配置", command=on_save).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.RIGHT)
        
        # 绑定快捷键
        dialog.bind('<Control-Return>', lambda e: on_save())
        dialog.bind('<Escape>', lambda e: on_cancel())
        

            
    def convert_to_normal_column(self, col_name):
        """将AI列转换为普通列"""
        result = messagebox.askyesno("确认转换", 
                                   f"确定要将AI列 '{col_name}' 转换为普通列吗？\n\n"
                                   f"这将删除该列的AI配置，但保留现有数据。")
        if result:
            self.table_manager.convert_to_normal_column(col_name)
            self.update_table_display()
            self.update_title()  # 更新标题栏
            self.update_status(f"已转换为普通列: {col_name}", "success")
            messagebox.showinfo("成功", f"列 '{col_name}' 已转换为普通列")
    
    def delete_specific_column(self, column_name):
        """删除指定列"""
        ai_columns = self.table_manager.get_ai_columns()
        is_ai_col = column_name in ai_columns
        col_type = "AI列" if is_ai_col else "普通列"
        
        # 确认删除
        result = messagebox.askyesno("确认删除", 
                                   f"确定要删除{col_type} '{column_name}' 吗？\n\n"
                                   f"此操作将删除该列的所有数据，无法撤销。")
        if result:
            # 执行删除
            success = self.table_manager.delete_column(column_name)
            if success:
                self.update_table_display()
                self.update_title()  # 更新标题栏
                self.update_status(f"已删除列: {column_name}", "success")
                messagebox.showinfo("成功", f"列 '{column_name}' 已删除")
            else:
                messagebox.showerror("错误", f"删除列 '{column_name}' 失败")
    
    def edit_specific_cell(self, row_index, col_name):
        """编辑指定单元格"""
        df = self.table_manager.get_dataframe()
        if df is not None:
            current_value = str(df.iloc[row_index, df.columns.get_loc(col_name)])
            self.edit_cell_dialog(row_index, col_name, current_value)
    
    def process_specific_cell(self, row_index, col_name):
        """处理指定的AI单元格"""
        ai_columns = self.table_manager.get_ai_columns()
        if col_name in ai_columns:
            # 获取AI列的配置
            config = ai_columns[col_name]
            if isinstance(config, dict):
                prompt = config.get("prompt", "")
                model = config.get("model", "gpt-4.1")
                output_mode = config.get("output_mode", "single")
                output_fields = config.get("output_fields", [])
            else:
                # 向后兼容
                prompt = config
                model = "gpt-4.1"
                output_mode = "single"
                output_fields = []
            
            # 处理单个单元格
            try:
                success, result = self.ai_processor.process_single_cell(
                    self.table_manager.get_dataframe(),
                    row_index,
                    col_name,
                    prompt,
                    model,
                    self.table_manager,
                    output_fields if output_mode == "multi" else None
                )
                
                if success:
                    self.update_table_display()
                    if output_mode == "multi" and isinstance(result, dict):
                        self.update_status(f"单元格 {col_name}[{row_index+1}] 多字段处理完成 (提取了 {len(result)} 个字段)", "success")
                    else:
                        self.update_status(f"单元格 {col_name}[{row_index+1}] 处理完成", "success")
                else:
                    # 如果是JSON解析失败，提供重试选项
                    if "JSON解析失败" in str(result) and output_mode == "multi":
                        retry_result = messagebox.askyesnocancel("解析失败", 
                                                               f"JSON解析失败: {result}\n\n"
                                                               f"是否手动重新解析当前响应？\n"
                                                               f"点击'是'重新解析，'否'跳过，'取消'查看原始响应")
                        if retry_result is True:  # 重新解析
                            self.retry_parse_cell(row_index, col_name, output_fields)
                        elif retry_result is False:  # 跳过
                            pass
                        else:  # 查看原始响应
                            self.show_original_response(row_index, col_name)
                    else:
                        self.update_status("单元格处理失败", "error")
                    
            except Exception as e:
                messagebox.showerror("错误", f"处理单元格时出错: {str(e)}")
                self.update_status("单元格处理失败", "error")
    
    def retry_parse_cell(self, row_index, col_name, expected_fields):
        """重试解析单个单元格的JSON响应"""
        try:
            success, result = self.ai_processor.retry_parse_single_cell(
                self.table_manager.get_dataframe(),
                row_index,
                col_name,
                expected_fields
            )
            
            if success:
                self.update_table_display()
                self.update_status(f"重新解析成功: {col_name}[{row_index+1}]", "success")
                messagebox.showinfo("成功", f"重新解析成功！\n{result}")
            else:
                self.update_status("重新解析失败", "error")
                messagebox.showerror("失败", f"重新解析失败: {result}")
                
        except Exception as e:
            messagebox.showerror("错误", f"重新解析时出错: {str(e)}")
            
    def show_original_response(self, row_index, col_name):
        """显示原始AI响应"""
        try:
            df = self.table_manager.get_dataframe()
            current_value = str(df.loc[row_index, col_name])
            
            # 提取原始响应（去掉解析错误标记）
            if "[解析错误:" in current_value:
                original_response = current_value.split("\n\n[解析错误:")[0]
            else:
                original_response = current_value
                
            # 创建对话框显示原始响应
            dialog = tk.Toplevel(self.root)
            dialog.title(f"原始AI响应 - {col_name}[{row_index+1}]")
            dialog.geometry("600x400")
            dialog.transient(self.root)
            dialog.grab_set()
            
            # 居中显示
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (300)
            y = (dialog.winfo_screenheight() // 2) - (200)
            dialog.geometry(f"600x400+{x}+{y}")
            
            # 文本框显示原始响应
            text_frame = ttk.Frame(dialog)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, width=70, height=20)
            text_widget.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
            
            scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.insert("1.0", original_response)
            text_widget.config(state=tk.DISABLED)
            
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="复制内容", 
                      command=lambda: self.copy_to_clipboard(original_response)).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="关闭", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
            
        except Exception as e:
            messagebox.showerror("错误", f"显示原始响应时出错: {str(e)}")
            
    def copy_to_clipboard(self, text):
        """复制文本到剪贴板"""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("复制成功", "内容已复制到剪贴板")
        
    def show_filter_dialog(self, column_name):
        """显示筛选对话框"""
        df = self.table_manager.get_dataframe()
        if df is None or column_name not in df.columns:
            messagebox.showwarning("警告", "无法筛选此列")
            return
            
        # 获取列的所有唯一值
        unique_values = df[column_name].astype(str).unique()
        unique_values = sorted([v for v in unique_values if v.strip() != ''])  # 排序并去除空值
        
        if not unique_values:
            messagebox.showinfo("提示", "此列没有可筛选的数据")
            return
            
        # 创建筛选对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"筛选 - {column_name}")
        dialog.geometry("450x600")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (225)
        y = (dialog.winfo_screenheight() // 2) - (300)
        dialog.geometry(f"450x600+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text=f"筛选列: {column_name}", 
                               font=('Microsoft YaHei UI', 14, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # 统计信息
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        total_count = len(df)
        unique_count = len(unique_values)
        ttk.Label(info_frame, text=f"总行数: {total_count} | 唯一值: {unique_count}",
                 foreground="gray").pack()
        
        # 操作按钮框架
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(control_frame, text="全选", command=lambda: self.toggle_all_filter_items(True, checkboxes)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="全不选", command=lambda: self.toggle_all_filter_items(False, checkboxes)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="反选", command=lambda: self.toggle_filter_selection(checkboxes)).pack(side=tk.LEFT)
        
        # 搜索框
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT, padx=(0, 5))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 值列表框架
        list_frame = ttk.LabelFrame(main_frame, text="选择要显示的值", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 滚动框架
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 创建复选框
        checkboxes = {}
        checkbox_vars = {}
        
        # 如果当前列已有筛选，使用之前的选择
        if self.filter_state['active'] and self.filter_state['column'] == column_name:
            selected_values = set(self.filter_state['selected_values'])
        else:
            selected_values = set(unique_values)  # 默认全选
        
        for value in unique_values:
            var = tk.BooleanVar(value=value in selected_values)
            checkbox_vars[value] = var
            
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, pady=1)
            
            checkbox = ttk.Checkbutton(frame, text=str(value), variable=var)
            checkbox.pack(side=tk.LEFT)
            checkboxes[value] = (checkbox, var)
            
            # 显示该值的行数
            count = len(df[df[column_name].astype(str) == str(value)])
            count_label = ttk.Label(frame, text=f"({count})", foreground="gray")
            count_label.pack(side=tk.RIGHT)
        
        # 搜索功能
        def filter_checkboxes():
            search_text = search_var.get().lower()
            for value, (checkbox, var) in checkboxes.items():
                if search_text in str(value).lower():
                    checkbox.pack(side=tk.LEFT)
                    checkbox.master.pack(fill=tk.X, pady=1)
                else:
                    checkbox.master.pack_forget()
        
        search_var.trace('w', lambda *args: filter_checkboxes())
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def apply_filter():
            selected = [value for value, (_, var) in checkboxes.items() if var.get()]
            
            if not selected:
                messagebox.showwarning("警告", "请至少选择一个值")
                return
                
            # 应用筛选
            self.apply_filter(column_name, selected)
            dialog.destroy()
            
        def cancel_filter():
            dialog.destroy()
            
        ttk.Button(button_frame, text="应用筛选", command=apply_filter).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=cancel_filter).pack(side=tk.RIGHT)
        
        # 绑定快捷键
        dialog.bind('<Return>', lambda e: apply_filter())
        dialog.bind('<Escape>', lambda e: cancel_filter())
        
    def toggle_all_filter_items(self, select_all, checkboxes):
        """全选/全不选筛选项"""
        for _, (_, var) in checkboxes.items():
            var.set(select_all)
            
    def toggle_filter_selection(self, checkboxes):
        """反选筛选项"""
        for _, (_, var) in checkboxes.items():
            var.set(not var.get())
            
    def apply_filter(self, column_name, selected_values):
        """应用筛选"""
        df = self.table_manager.get_dataframe()
        if df is None or column_name not in df.columns:
            return
            
        # 找到匹配的行索引
        mask = df[column_name].astype(str).isin(selected_values)
        filtered_indices = df[mask].index.tolist()
        
        # 更新筛选状态
        self.filter_state = {
            'active': True,
            'column': column_name,
            'selected_values': selected_values,
            'all_values': df[column_name].astype(str).unique().tolist(),
            'filtered_indices': filtered_indices
        }
        
        # 更新显示
        self.update_table_display()
        
        # 更新状态栏
        total_count = len(df)
        filtered_count = len(filtered_indices)
        self.update_status(f"已筛选 {column_name}: 显示 {filtered_count}/{total_count} 行", "success")
        
    def clear_filter(self):
        """清除筛选"""
        self.filter_state = {
            'active': False,
            'column': None,
            'selected_values': [],
            'all_values': [],
            'filtered_indices': []
        }
        
        # 更新显示
        self.update_table_display()
        self.update_status("已清除筛选", "success")
        
    def get_filtered_dataframe(self):
        """获取筛选后的数据框"""
        df = self.table_manager.get_dataframe()
        if df is None:
            return None
            
        if self.filter_state['active'] and self.filter_state['filtered_indices']:
            return df.iloc[self.filter_state['filtered_indices']].reset_index(drop=True)
        else:
            return df
    
    def process_entire_column(self, col_name):
        """处理整个AI列"""
        ai_columns = self.table_manager.get_ai_columns()
        if col_name not in ai_columns:
            messagebox.showwarning("警告", f"'{col_name}' 不是AI列")
            return
        
        df = self.table_manager.get_dataframe()
        if df is None:
            return
            
        # 确认处理
        row_count = len(df)
        result = messagebox.askyesno("确认处理", 
                                   f"即将处理整个 '{col_name}' 列，共 {row_count} 行。\n"
                                   f"这可能需要一些时间，是否继续？")
        if not result:
            return
            
        try:
            self.update_status(f"正在处理整列 {col_name}...", "normal")
            
            success_count = 0
            
            # 获取AI列配置
            config = ai_columns[col_name]
            if isinstance(config, dict):
                prompt_template = config.get("prompt", "")
                model = config.get("model", "gpt-4.1")
            else:
                # 向后兼容
                prompt_template = config
                model = "gpt-4.1"
            
            # 处理每一行，优化性能
            for row_index in range(len(df)):
                try:
                    success, result = self.ai_processor.process_single_cell(
                        df, row_index, col_name, prompt_template, model, self.table_manager
                    )
                    
                    if success:
                        success_count += 1
                        
                    # 更新进度条 - 使用新的表格进度条
                    self.update_table_progress(row_index + 1, row_count, f"处理 {col_name}")
                    
                    # 减少界面更新频率，每5行更新一次显示
                    if (row_index + 1) % 5 == 0 or row_index == len(df) - 1:
                        self.update_table_display()
                    
                except Exception as e:
                    print(f"处理第{row_index+1}行时出错: {e}")
                    
            # 最终更新显示
            self.update_table_display()
            self.update_status(f"列 {col_name} 处理完成 ({success_count}/{row_count})", "success")
            messagebox.showinfo("完成", f"列 '{col_name}' 处理完成！\n成功: {success_count}/{row_count}")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理列时出错: {str(e)}")
            selection = self.tree.selection()
            if not selection:
                self.update_status("请先选择一个单元格", "error")
                return
                
            item = selection[0]
            column = self.tree.identify_column(event.x)
            
            if not column:
                self.update_status("无法识别列位置", "error")
                return
                
            # 获取列索引和列名
            col_index = int(column.replace('#', '')) - 1
            if col_index < 0:
                return
                
            df = self.table_manager.get_dataframe()
            if df is None:
                self.update_status("没有数据可编辑", "error")
                return
                
            column_names = list(df.columns)
            if col_index >= len(column_names):
                return
                
            col_name = column_names[col_index]
            
            # 获取当前值
            values = self.tree.item(item, 'values')
            if col_index < len(values):
                current_value = values[col_index]
                # 处理被截断的文本，从原始数据获取完整值
                row_index = self.tree.index(item)
                current_value = str(df.iloc[row_index, col_index])
            else:
                current_value = ""
                
            # 获取行索引
            row_index = self.tree.index(item)
            
            # 创建编辑对话框
            self.edit_cell_dialog(row_index, col_name, current_value)
            
        except Exception as e:
            self.update_status(f"编辑失败: {str(e)}", "error")
            print(f"双击编辑错误: {e}")
            
    def edit_cell_dialog(self, row_index, col_name, current_value):
        """单元格编辑对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑单元格")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")
        
        # 标题标签
        title_label = ttk.Label(dialog, text=f"编辑第 {row_index+1} 行 - {col_name}", 
                               style='Title.TLabel')
        title_label.pack(pady=10)
        
        # 当前值显示（如果很长的话）
        if len(str(current_value)) > 100:
            preview_frame = ttk.LabelFrame(dialog, text="当前内容预览", padding=5)
            preview_frame.pack(fill=tk.X, padx=10, pady=5)
            
            preview_text = str(current_value)[:200] + "..." if len(str(current_value)) > 200 else str(current_value)
            ttk.Label(preview_frame, text=preview_text, wraplength=450).pack()
        
        # 编辑框架
        edit_frame = ttk.LabelFrame(dialog, text="编辑内容", padding=5)
        edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 文本框
        text_widget = tk.Text(edit_frame, wrap=tk.WORD, height=12, width=60)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", str(current_value))
        text_widget.focus()
        
        # 滚动条
        scrollbar = ttk.Scrollbar(edit_frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def save_changes():
            try:
                new_value = text_widget.get("1.0", tk.END).strip()
                
                # 保存状态用于撤销
                self.save_state(f"编辑单元格 {col_name}[第{row_index+1}行]")
                
                # 更新数据框
                df = self.table_manager.get_dataframe()
                df.iloc[row_index, df.columns.get_loc(col_name)] = new_value
                
                # 标记有更改
                self.table_manager.changes_made = True
                
                # 刷新表格显示
                self.update_table_display()
                self.update_title()  # 更新标题栏
                
                # 更新预览面板
                if self.current_preview_cell and self.current_preview_cell['row_index'] == row_index and self.current_preview_cell['col_name'] == col_name:
                    self.update_content_preview(row_index, col_name, new_value)
                
                self.update_status(f"已更新 {col_name} (第{row_index+1}行)", "success")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("保存失败", f"保存时出错: {str(e)}")
                
        def cancel_edit():
            dialog.destroy()
            
        # 按钮
        save_btn = ttk.Button(button_frame, text="💾 保存", command=save_changes)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="❌ 取消", command=cancel_edit)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # 提示标签
        tip_label = ttk.Label(button_frame, text="快捷键: Ctrl+Enter保存, Escape取消", 
                             style='Subtitle.TLabel')
        tip_label.pack(side=tk.LEFT, padx=20)
        
        # 绑定快捷键
        dialog.bind('<Control-Return>', lambda e: save_changes())
        dialog.bind('<Escape>', lambda e: cancel_edit())
        
        # 设置焦点到保存按钮（回车时可以触发）
        save_btn.focus_set()
        
    def save_project(self):
        """保存项目文件"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可保存，请先创建表格或导入数据")
            return
        
        # 如果已有项目文件路径，直接保存
        if self.current_project_path:
            file_path = self.current_project_path
        else:
            # 新项目，弹出文件选择对话框
            file_path = filedialog.asksaveasfilename(
                title="保存项目文件",
                defaultextension=".aie",
                filetypes=[
                    ("AI Excel项目文件", "*.aie"),
                    ("所有文件", "*.*")
                ]
            )
        
        if file_path:
            try:
                self.update_status("正在保存项目...", "normal")
                self.root.update()
                
                success, message = self.project_manager.save_project(
                    file_path, self.table_manager, self.ai_processor, self._get_column_widths()
                )
                
                if success:
                    # 更新当前项目路径
                    self.current_project_path = file_path
                    filename = os.path.basename(file_path)
                    self.info_label.config(text=f"📁 {filename}")
                    self.update_status(f"项目已保存到: {filename}", "success")
                    messagebox.showinfo("成功", f"项目已保存到: {filename}")
                    self.table_manager.reset_changes_flag() # 保存成功后重置更改标志
                    self.update_title()  # 更新标题栏，移除未保存标记
                    return True, "项目保存成功"
                else:
                    self.update_status("保存失败", "error")
                    messagebox.showerror("错误", message)
                    
            except Exception as e:
                messagebox.showerror("错误", f"保存项目时出错: {str(e)}")
                self.update_status("保存失败", "error")

    def save_project_as(self):
        """另存为项目文件"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可保存，请先创建表格或导入数据")
            return
            
        # 总是弹出文件选择对话框
        file_path = filedialog.asksaveasfilename(
            title="另存为项目文件",
            defaultextension=".aie",
            filetypes=[
                ("AI Excel项目文件", "*.aie"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.update_status("正在保存项目...", "normal")
                self.root.update()
                
                success, message = self.project_manager.save_project(
                    file_path, self.table_manager, self.ai_processor
                )
                
                if success:
                    # 更新当前项目路径为新路径
                    self.current_project_path = file_path
                    filename = os.path.basename(file_path)
                    self.info_label.config(text=f"📁 {filename}")
                    self.update_status(f"项目已另存为: {filename}", "success")
                    messagebox.showinfo("成功", f"项目已另存为: {filename}")
                    self.table_manager.reset_changes_flag() # 另存为成功后重置更改标志
                    self.update_title()  # 更新标题栏，移除未保存标记
                else:
                    self.update_status("另存为失败", "error")
                    messagebox.showerror("错误", message)
                    
            except Exception as e:
                messagebox.showerror("错误", f"另存为项目时出错: {str(e)}")
                self.update_status("另存为失败", "error")

    def load_project(self):
        """加载项目文件"""
        # 检查未保存的更改
        if not self.check_unsaved_changes_before_action("加载新项目"):
            return
            
        file_path = filedialog.askopenfilename(
            title="选择项目文件",
            filetypes=[
                ("AI Excel项目文件", "*.aie"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.update_status("正在加载项目...", "normal")
                self.root.update()
                
                success, message, column_widths = self.project_manager.load_project(
                    file_path, self.table_manager
                )
                
                if success:
                    # 记录当前项目文件路径
                    self.current_project_path = file_path
                    self.hide_welcome()
                    self.update_table_display(column_widths=column_widths) # 传递列宽
                    filename = os.path.basename(file_path)
                    self.info_label.config(text=f"📁 {filename}")
                    self.update_status(f"项目已加载: {filename}", "success")
                    messagebox.showinfo("成功", message)
                    self.table_manager.reset_changes_flag() # 加载成功后重置更改标志
                    self.update_title()  # 更新标题栏
                else:
                    self.update_status("加载失败", "error")
                    messagebox.showerror("错误", message)
                    
            except Exception as e:
                messagebox.showerror("错误", f"加载项目时出错: {str(e)}")
                self.update_status("加载失败", "error")

    def import_data_file(self):
        """导入文件"""
        # 检查未保存的更改
        if not self.check_unsaved_changes_before_action("导入新数据文件"):
            return
            
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("JSONL文件", "*.jsonl"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.update_status("正在导入文件...", "normal")
                self.root.update()
                
                success = self.table_manager.load_file(file_path)
                if success:
                    # 清除项目文件路径（导入数据文件不是项目文件）
                    self.current_project_path = None
                    self.hide_welcome()
                    self.update_table_display()
                    filename = os.path.basename(file_path)
                    self.info_label.config(text=f"📁 {filename}")
                    self.update_status(f"已导入: {filename}", "success")
                    self.update_title()  # 更新标题栏
                else:
                    messagebox.showerror("错误", "文件导入失败")
                    self.update_status("导入失败", "error")
                    
            except Exception as e:
                messagebox.showerror("错误", f"导入文件时出错: {str(e)}")
                self.update_status("导入失败", "error")

    def update_table_display(self, column_widths=None):
        """更新表格显示"""
        print("开始更新表格显示")
        
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 获取原始数据框和筛选后的数据框
        original_df = self.table_manager.get_dataframe()
        if original_df is not None:
            # 获取要显示的数据（筛选后的）
            if self.filter_state['active'] and self.filter_state['filtered_indices']:
                df = original_df.iloc[self.filter_state['filtered_indices']].reset_index(drop=True)
            else:
                df = original_df
            print(f"数据框大小: {df.shape}")
            print(f"列名: {list(df.columns)}")
            
            # 设置列
            columns = list(df.columns)
            self.tree["columns"] = columns
            self.tree["show"] = "headings"
            
            # 设置列标题和宽度，并添加边框效果
            ai_columns = self.table_manager.get_ai_columns()
            long_text_columns = self.table_manager.get_long_text_columns()
            print(f"DEBUG: main.py - update_table_display - AI Columns: {ai_columns}") # Debug print
            print(f"DEBUG: main.py - update_table_display - Long Text Columns: {long_text_columns}") # Debug print
            for i, col in enumerate(columns):
                display_col_name = col
                heading_style = 'Normal.Treeview.Heading'  # Default style
                
                if col in ai_columns:
                    display_col_name = f"★ {col} ★"  # AI列添加星号
                    heading_style = 'AI.Treeview.Heading'
                elif col in long_text_columns:
                    display_col_name = f"📄 {col}"  # 长文本列添加文档图标
                    heading_style = 'LongText.Treeview.Heading'
                
                # 添加排序指示符（如果当前正在排序这一列）
                sort_indicator = self.get_sort_indicator(col)
                display_col_name += sort_indicator
                
                # 添加筛选指示符（如果当前正在筛选这一列）
                if self.filter_state['active'] and self.filter_state['column'] == col:
                    display_col_name += " 🔍"
                    
                self.tree.heading(col, text=display_col_name,
                                 anchor='w')  # 左对齐，不绑定点击事件
                
                # 如果有保存的列宽，在后面应用，这里先用默认逻辑
                if column_widths is None:
                    # 根据内容调整列宽 - 增大最小宽度改善可读性
                    max_width = max(len(col) * 12, 150)  # 最小150像素，字体更大
                    self.tree.column(col, width=max_width, minwidth=120, anchor='w')  # 左对齐
            
            # 如果从项目加载了列宽，则应用它们
            if column_widths:
                self._apply_column_widths(column_widths)
                
            # 配置行的边框标签样式
            self.tree.tag_configure('row_border', 
                                   background='#ffffff',
                                   foreground='#1a202c')
            
            # 配置交替行颜色（斑马纹效果）
            self.tree.tag_configure('odd_row', 
                                   background='#f8fafc',  # 浅灰背景
                                   foreground='#1a202c')
            self.tree.tag_configure('even_row', 
                                   background='#ffffff',  # 白色背景
                                   foreground='#1a202c')
                                   
            # 配置AI列单元格样式
            self.tree.tag_configure('ai_cell', 
                                   background='#e3f2fd',  # 浅蓝色背景
                                   foreground='#1a202c')  # 深色文字
                
            # 插入数据并应用行样式
            for index, row in df.iterrows():
                values = []
                for val in row:
                    # 处理长文本显示 - 增加显示长度
                    str_val = str(val) if val is not None else ""
                    if len(str_val) > 80:  # 增加显示长度
                        str_val = str_val[:77] + "..."
                    values.append(str_val)
                
                # 使用交替行颜色创建网格效果
                row_tag = 'odd_row' if index % 2 == 0 else 'even_row'
                item = self.tree.insert("", "end", values=values, tags=(row_tag,))
                print(f"插入第{index+1}行: {values}")
                

                
            # 更新表格标题
            row_count = len(df)
            col_count = len(df.columns)
            ai_count = len(self.table_manager.get_ai_columns())
            
            # 构建表格标题，包含筛选状态
            title = f"📊 数据表格 - {row_count}行 {col_count}列 (AI列: {ai_count})"
            if self.filter_state['active']:
                original_count = len(original_df)
                filter_column = self.filter_state['column']
                title += f" | 🔍 已筛选 {filter_column}: {row_count}/{original_count} 行"
            
            self.table_frame.config(text=title)
            
            # 恢复列高亮效果
            if hasattr(self, 'highlighted_column') and self.highlighted_column is not None:
                self.highlight_column(self.highlighted_column)
            
            print(f"表格更新完成，显示{row_count}行{col_count}列")
        else:
            print("数据框为空")
            
        # 配置选择模式为单个单元格
        self.tree.configure(selectmode='browse')  # 只能选择一个项目

    def create_ai_column(self):
        """新建AI列"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格或导入数据文件")
            return
            
        dialog = AIColumnDialog(self.root, self.table_manager.get_column_names())
        result = dialog.show()
        
        if result:
            column_name = result['column_name']
            prompt_template = result['prompt']
            ai_model = result['model']
            processing_params = result['processing_params']
            output_mode = result.get('output_mode', 'single')
            output_fields = result.get('output_fields', [])
            
            # 检查列名冲突（对话框已经做了验证，这里做最后检查）
            existing_columns = self.table_manager.get_column_names()
            if column_name in existing_columns:
                messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
                return

            # 保存状态用于撤销
            self.save_state(f"创建AI列: {column_name}")
            
            self.table_manager.add_ai_column(column_name, prompt_template, ai_model, processing_params, output_mode, output_fields)
            
            if output_mode == "multi" and output_fields:
                self.update_status(f"已添加多字段AI列: {column_name} (模型: {ai_model}, 字段: {', '.join(output_fields)})", "success")
            else:
                self.update_status(f"已添加AI列: {column_name} (模型: {ai_model})", "success")
            
            self.update_table_display()
            self.update_title()  # 更新标题栏
            
    def create_normal_column(self, position=None, side=None): # Added position and side parameters
        """新建普通列. 若提供了position, 则在该位置插入,否则追加到末尾.""" # Docstring updated
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格或导入数据文件")
            return
            
        # 询问用户创建方式
        choice = self.ask_normal_column_creation_type()
        if choice is None: # User cancelled
            return
            
        if choice == "manual":
            # 手动创建
            column_name = tk.simpledialog.askstring("新建列", "请输入列名:")
            if column_name and column_name.strip():
                column_name = column_name.strip()
                if column_name not in self.table_manager.get_column_names():
                    success = False
                    if position is not None:
                        success = self.table_manager.insert_column_at_position(position, column_name)
                        status_message_verb = f"已在{side}侧插入普通列" if side else "已插入普通列"
                    else:
                        success = self.table_manager.add_normal_column(column_name)
                        status_message_verb = "已添加列"
                    
                    if success:
                        self.update_table_display()
                        self.update_status(f"{status_message_verb}: {column_name}", "success")
                        self.update_title()  # 更新标题栏
                    else:
                        messagebox.showerror("错误", f"创建普通列 '{column_name}' 失败")
                else:
                    messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
        elif choice == "jsonl":
            # 从JSONL导入
            # Pass position to import_from_jsonl. If None, it will append.
            self.import_from_jsonl(position=position, side=side) 
            
    def ask_normal_column_creation_type(self):
        """询问普通列创建方式"""
        dialog = tk.Toplevel(self.root)
        dialog.title("选择创建方式")
        dialog.geometry("350x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (175)
        y = (dialog.winfo_screenheight() // 2) - (100)
        dialog.geometry(f"350x200+{x}+{y}")

        result = [None]  # 使用列表存储结果

        ttk.Label(dialog, text="选择普通列创建方式:", style='Title.TLabel').pack(pady=15)

        choice_var = tk.StringVar(value="manual")

        ttk.Radiobutton(dialog, text="📝 手动创建空白列", variable=choice_var, value="manual").pack(anchor=tk.W, padx=30, pady=5)
        ttk.Radiobutton(dialog, text="📥 从JSONL文件导入", variable=choice_var, value="jsonl").pack(anchor=tk.W, padx=30, pady=5)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)

        def on_ok():
            result[0] = choice_var.get()
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        ttk.Button(button_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        # 绑定回车和ESC键
        dialog.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())

        dialog.wait_window()
        return result[0]
        
    def import_from_jsonl(self, position=None, side=None): # Added position and side
        """从JSONL文件导入数据. 若提供position, 则尝试在该位置插入.""" # Docstring updated
        from jsonl_import_dialog import JsonlImportDialog
        
        def on_import_result(result):
            if result:
                match_field = result['match_field']
                source_field = result['source_field']
                column_name = result['column_name']
                jsonl_data = result['jsonl_data']
                
                # 执行导入
                # Pass position to table_manager.import_from_jsonl
                success, message = self.table_manager.import_from_jsonl(
                    match_field, source_field, column_name, jsonl_data, position=position
                )
                
                if success:
                    self.update_table_display()
                    # Modify status message if side is available (for insertion context)
                    status_message = message
                    if position is not None and side:
                        # Example: "已在左侧通过JSONL导入: new_col" - customize as needed
                        status_message = f"已在{side}侧通过JSONL导入: {column_name}" 
                                       
                    self.update_status(status_message, "success")
                    self.update_title()
                    messagebox.showinfo("成功", message) # Original message for dialog
                else:
                    messagebox.showerror("失败", message)
        
        try:
            dialog = JsonlImportDialog(self.root, self.table_manager, on_import_result)
            dialog.show()
        except Exception as e:
            messagebox.showerror("错误", f"打开JSONL导入对话框失败: {e}")
                
    def create_long_text_column(self):
        """创建长文本列"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格或导入数据文件")
            return
            
        print("DEBUG: Entering create_long_text_column") # Debug print

        from long_text_column_dialog import LongTextColumnDialog
        
        def on_result(result):
            if result:
                column_name = result['column_name']
                filename_field = result['filename_field']
                folder_path = result['folder_path']
                preview_length = result['preview_length']
                
                # 添加长文本列
                success = self.table_manager.add_long_text_column(
                    column_name, filename_field, folder_path, preview_length
                )
                
                if success:
                    self.update_table_display()
                    self.update_status(f"已添加长文本列: {column_name}", "success")
                else:
                    messagebox.showerror("错误", "创建长文本列失败")
        
        try:
            dialog = LongTextColumnDialog(self.root, self.table_manager, on_result)
            print("DEBUG: LongTextColumnDialog instance created.") # Debug print
            dialog.show()
            print("DEBUG: LongTextColumnDialog closed.") # Debug print
        except Exception as e:
            print(f"ERROR: Failed to open LongTextColumnDialog: {e}") # Debug print
            messagebox.showerror("错误", f"打开长文本列对话框失败: {e}")
                
    def add_row(self):
        """添加新行"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格")
            return
        
        # 保存状态用于撤销
        self.save_state("添加行")
            
        success = self.table_manager.add_row()
        if success:
            self.update_table_display()
            self.update_status("已添加新行", "success")
            self.update_title()  # 更新标题栏
    
    def insert_row_at_position(self, position, direction):
        """在指定位置插入行"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格")
            return
            
        success = self.table_manager.insert_row_at_position(position)
        if success:
            self.update_table_display()
            side = "上方" if direction == "above" else "下方"
            self.update_status(f"已在第{position+1}行{side}插入新行", "success")
        else:
            messagebox.showerror("错误", "插入行失败")
    
    def insert_column_dialog(self, direction):
        """插入列对话框"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格或导入数据文件")
            return
            
        # 获取当前选中列的位置
        position = 0  # 默认位置
        # Ensure self.selection_info and 'column_index' exist before accessing
        current_col_index = None
        if hasattr(self, 'selection_info') and self.selection_info:
            current_col_index = self.selection_info.get('column_index')

        if current_col_index is not None:
            col_index = current_col_index
            position = col_index if direction == "left" else col_index + 1
        else:
            # 如果没有选中，根据方向决定插入开头还是末尾
            df = self.table_manager.get_dataframe()
            if df is not None:
                 position = 0 if direction == "left" else len(df.columns)
            # If df is None, position remains 0, though create_table checks should prevent this path.

        # 询问用户要插入哪种类型的列
        column_type_choice = self.ask_column_type_for_insertion()
        if column_type_choice is None: # 用户取消
            return

        column_name = None
        success = False # Initialize success to False
        side = "左" if direction == "left" else "右"

        if column_type_choice == "ai":
            from ai_column_dialog import AIColumnDialog
            dialog = AIColumnDialog(self.root, self.table_manager.get_column_names())
            result = dialog.show()
            if result:
                column_name = result['column_name']
                prompt_template = result['prompt']
                ai_model = result['model']
                processing_params = result['processing_params']
                output_mode = result.get('output_mode', 'single')
                output_fields = result.get('output_fields', [])
                if column_name in self.table_manager.get_column_names():
                    messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
                    return
                success = self.table_manager.insert_column_at_position(
                    position, column_name, prompt_template=prompt_template, is_ai_column=True, 
                    ai_model=ai_model, processing_params=processing_params, output_mode=output_mode, 
                    output_fields=output_fields
                )
                if success:
                    if output_mode == "multi" and output_fields:
                        self.update_status(f"已在{side}侧插入多字段AI列: {column_name} (字段: {', '.join(output_fields)})", "success")
                    else:
                        self.update_status(f"已在{side}侧插入AI列: {column_name}", "success")
                else:
                    messagebox.showerror("错误", "插入AI列失败")

        elif column_type_choice == "long_text":
            from long_text_column_dialog import LongTextColumnDialog
            dialog = LongTextColumnDialog(self.root, self.table_manager, lambda r: None)
            dialog.show()
            result = dialog.result
            if result:
                column_name = result['column_name']
                filename_field = result['filename_field']
                folder_path = result['folder_path']
                preview_length = result['preview_length']

                if column_name in self.table_manager.get_column_names():
                    messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
                    return
                
                # 先插入普通列，再转换为长文本列
                success = self.table_manager.insert_column_at_position(position, column_name)
                if success:
                    # 添加长文本列配置
                    self.table_manager.long_text_columns[column_name] = {
                        "filename_field": filename_field,
                        "folder_path": folder_path,
                        "preview_length": preview_length
                    }
                    # 刷新长文本列内容
                    self.table_manager.refresh_long_text_column(column_name)
                    self.update_status(f"已在{side}侧插入长文本列: {column_name}", "success")
                else:
                    messagebox.showerror("错误", "插入长文本列失败")

        elif column_type_choice == "normal": # Normal column
            # 不再直接创建，而是调用 create_normal_column，它内部会处理创建方式（手动或JSONL）
            # create_normal_column 将处理列名输入、查重和实际的插入逻辑
            # 它也需要知道插入的位置 position 和 side (用于状态消息)
            self.create_normal_column(position=position, side=side)
            # success 和 column_name 的处理将移到 create_normal_column 内部
            # 因此这里不需要再处理 success 或调用 update_table_display

        # if success: # This check is now handled within create_normal_column or other specific column creation methods
        #     self.update_table_display()

    def ask_column_type_for_insertion(self):
        """询问用户要插入的列类型"""
        dialog = tk.Toplevel(self.root)
        dialog.title("选择列类型")
        dialog.geometry("300x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"300x200+{x}+{y}")

        column_type_var = tk.StringVar(value="ai") # Default to AI column
        result = [None] # Use a list to store the result

        ttk.Label(dialog, text="请选择要插入的列类型:", style='Title.TLabel').pack(pady=10)

        ttk.Radiobutton(dialog, text="AI处理列", variable=column_type_var, value="ai").pack(anchor=tk.W, padx=20)
        ttk.Radiobutton(dialog, text="长文本列", variable=column_type_var, value="long_text").pack(anchor=tk.W, padx=20)
        ttk.Radiobutton(dialog, text="普通列", variable=column_type_var, value="normal").pack(anchor=tk.W, padx=20)

        def on_ok():
            result[0] = column_type_var.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        dialog.wait_window()
        return result[0]

    def clear_data(self):
        """清空数据"""
        if self.table_manager.get_dataframe() is None:
            return
            
        result = messagebox.askyesno("确认", "确定要清空所有数据吗？")
        if result:
            self.table_manager.clear_all_data()
            self.show_welcome()
            self.update_title()  # 更新标题栏
            self.update_status("已清空数据", "success")

    def test_ai_connection(self):
        """测试AI连接"""
        try:
            self.update_status("正在测试AI连接...", "normal")
            success, message = self.ai_processor.test_connection()
            
            if success:
                messagebox.showinfo("连接测试", f"AI连接正常！\n响应: {message}")
                self.update_status("AI连接正常", "success")
            else:
                messagebox.showerror("连接测试", f"AI连接失败！\n错误: {message}")
                self.update_status("AI连接失败", "error")
                
        except Exception as e:
            messagebox.showerror("连接测试", f"测试连接时出错: {str(e)}")
            self.update_status("连接测试失败", "error")

    def update_progress(self, current, total):
        """更新进度条"""
        progress = (current / total) * 100
        self.progress_bar['value'] = progress
        self.update_status(f"正在处理: {current}/{total} ({progress:.1f}%)", "normal")
        self.root.update()
        
    def update_table_progress(self, current, total, message="正在处理"):
        """更新表格区域的进度条"""
        if total > 0:
            progress = (current / total) * 100
            self.table_progress_bar['value'] = progress
            self.progress_label.config(text=f"{message}: {current}/{total}")
            if current >= total:
                # 处理完成，延迟隐藏进度条
                self.root.after(1000, self.hide_table_progress)
        else:
            self.table_progress_bar['value'] = 0
            self.progress_label.config(text="")
        self.root.update()
        
    def hide_table_progress(self):
        """隐藏表格进度条"""
        self.table_progress_bar['value'] = 0
        self.progress_label.config(text="")
        
    def delete_column(self):
        """删除列"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据表格")
            return
            
        columns = self.table_manager.get_column_names()
        if not columns:
            messagebox.showwarning("警告", "没有可删除的列")
            return
            
        # 选择要删除的列
        dialog = tk.Toplevel(self.root)
        dialog.title("删除列")
        dialog.geometry("350x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (350 // 2)
        y = (dialog.winfo_screenheight() // 2) - (250 // 2)
        dialog.geometry(f"350x250+{x}+{y}")
        
        ttk.Label(dialog, text="选择要删除的列:", style='Title.TLabel').pack(pady=10)
        
        # 列表框
        listbox_frame = ttk.Frame(dialog)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        listbox = tk.Listbox(listbox_frame, selectmode=tk.SINGLE)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # 添加列到列表，标注AI列
        ai_columns = self.table_manager.get_ai_columns()
        for col in columns:
            display_text = f"{col}"
            if col in ai_columns:
                display_text += " (AI列)"
            listbox.insert(tk.END, display_text)
            
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_delete():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("警告", "请选择要删除的列")
                return
                
            col_index = selection[0]
            col_name = columns[col_index]
            
            # 确认删除
            is_ai_col = col_name in ai_columns
            col_type = "AI列" if is_ai_col else "普通列"
            
            result = messagebox.askyesno("确认删除", 
                                       f"确定要删除{col_type} '{col_name}' 吗？\n\n"
                                       f"此操作将删除该列的所有数据，无法撤销。")
            if result:
                # 执行删除
                success = self.table_manager.delete_column(col_name)
                if success:
                    self.update_table_display()
                    self.update_status(f"已删除列: {col_name}", "success")
                    messagebox.showinfo("成功", f"列 '{col_name}' 已删除")
                    dialog.destroy()
                else:
                    messagebox.showerror("错误", f"删除列 '{col_name}' 失败")
                    
        def on_cancel():
            dialog.destroy()
            
        ttk.Button(button_frame, text="🗑️ 删除", command=on_delete).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="❌ 取消", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # 提示信息
        tip_label = ttk.Label(dialog, text="提示：删除操作无法撤销", 
                             style='Subtitle.TLabel', foreground='red')
        tip_label.pack(pady=5)

    def export_excel(self):
        """导出Excel"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        
        if file_path:
            try:
                self.table_manager.export_excel(file_path)
                messagebox.showinfo("成功", f"已导出到: {file_path}")
                self.update_status(f"已导出: {os.path.basename(file_path)}", "success")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
                self.update_status("导出失败", "error")
                
    def export_csv(self):
        """导出CSV"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="保存CSV文件",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv")]
        )
        
        if file_path:
            try:
                self.table_manager.export_csv(file_path)
                messagebox.showinfo("成功", f"已导出到: {file_path}")
                self.update_status(f"已导出: {os.path.basename(file_path)}", "success")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
                self.update_status("导出失败", "error")
                
    def export_jsonl(self):
        """导出JSONL"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="保存JSONL文件",
            defaultextension=".jsonl",
            filetypes=[("JSONL文件", "*.jsonl")]
        )
        
        if file_path:
            try:
                self.table_manager.export_jsonl(file_path)
                messagebox.showinfo("成功", f"已导出到: {file_path}")
                self.update_status(f"已导出: {os.path.basename(file_path)}", "success")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
                self.update_status("导出失败", "error")
    
    def show_export_selection(self):
        """显示导出字段选择对话框"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        # 创建选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("选择导出字段")
        dialog.geometry("550x600")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (550 // 2)
        y = (dialog.winfo_screenheight() // 2) - (600 // 2)
        dialog.geometry(f"550x600+{x}+{y}")
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="选择要导出的字段", style='Title.TLabel').pack(pady=(0, 20))
        
        # 全选/全不选按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="全选", command=lambda: self.toggle_all_selection(checkboxes, True)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="全不选", command=lambda: self.toggle_all_selection(checkboxes, False)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="反选", command=lambda: self.toggle_all_selection(checkboxes, None)).pack(side=tk.LEFT, padx=(0, 5))
        
        # 滚动框架
        canvas = tk.Canvas(main_frame, height=350)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 获取所有列
        all_columns = self.table_manager.get_column_names()
        ai_columns = self.table_manager.get_ai_columns()
        
        # 如果没有保存的选择，默认全选
        if not self.export_selection['selected_columns']:
            self.export_selection['selected_columns'] = all_columns.copy()
        
        # 创建复选框
        checkboxes = {}
        for i, col_name in enumerate(all_columns):
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, pady=2, padx=10)
            
            var = tk.BooleanVar()
            var.set(col_name in self.export_selection['selected_columns'])
            checkboxes[col_name] = var
            
            # 复选框
            cb = ttk.Checkbutton(frame, text=col_name, variable=var)
            cb.pack(side=tk.LEFT)
            
            # AI列标识
            if col_name in ai_columns:
                config = ai_columns[col_name]
                if isinstance(config, dict):
                    model = config.get("model", "gpt-4.1")
                else:
                    model = "gpt-4.1"
                ttk.Label(frame, text=f"(AI列 - {model})", foreground="blue", font=('Arial', 8)).pack(side=tk.LEFT, padx=(10, 0))
            else:
                ttk.Label(frame, text="(普通列)", foreground="gray", font=('Arial', 8)).pack(side=tk.LEFT, padx=(10, 0))
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(20, 0))
        
        # 格式选择 - 改为竖向布局
        format_frame = ttk.LabelFrame(bottom_frame, text="导出格式", padding="20")
        format_frame.pack(fill=tk.X, pady=(0, 15))
        
        format_var = tk.StringVar(value="excel")
        
        # 使用竖向Pack布局，确保每个选项都有足够空间
        ttk.Radiobutton(format_frame, text="📊 Excel (.xlsx)", variable=format_var, value="excel").pack(anchor=tk.W, pady=3)
        ttk.Radiobutton(format_frame, text="📋 CSV (.csv)", variable=format_var, value="csv").pack(anchor=tk.W, pady=3)
        ttk.Radiobutton(format_frame, text="📄 JSON (.json)", variable=format_var, value="json").pack(anchor=tk.W, pady=3)
        ttk.Radiobutton(format_frame, text="📝 JSONL (.jsonl)", variable=format_var, value="jsonl").pack(anchor=tk.W, pady=3)
        
        # 确定取消按钮
        action_frame = ttk.Frame(bottom_frame)
        action_frame.pack(fill=tk.X)
        
        def on_export():
            # 获取选中的列
            selected = [col for col, var in checkboxes.items() if var.get()]
            if not selected:
                messagebox.showwarning("警告", "请至少选择一个字段")
                return
                
            # 保存选择
            self.export_selection['selected_columns'] = selected
            
            # 执行导出
            export_format = format_var.get()
            self.export_selected_columns(selected, export_format)
            dialog.destroy()
            
        def on_cancel():
            dialog.destroy()
            
        ttk.Button(action_frame, text="导出", command=on_export).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(action_frame, text="取消", command=on_cancel).pack(side=tk.RIGHT)
        
        # 选中数量显示
        count_label = ttk.Label(action_frame, text="", foreground="gray")
        count_label.pack(side=tk.LEFT)
        
        def update_count():
            selected_count = sum(1 for var in checkboxes.values() if var.get())
            count_label.config(text=f"已选择 {selected_count}/{len(all_columns)} 个字段")
            dialog.after(100, update_count)
            
        update_count()
    
    def toggle_all_selection(self, checkboxes, state):
        """切换所有选择状态"""
        if state is None:  # 反选
            for var in checkboxes.values():
                var.set(not var.get())
        else:  # 全选或全不选
            for var in checkboxes.values():
                var.set(state)
    
    def export_selected_columns(self, selected_columns, format_type):
        """导出选中的列"""
        try:
            df = self.table_manager.get_dataframe()
            if df is None:
                messagebox.showwarning("警告", "没有数据可导出")
                return
                
            # 过滤选中的列
            export_df = df[selected_columns].copy()
            
            # 根据格式选择文件
            if format_type == "excel":
                file_path = filedialog.asksaveasfilename(
                    title="保存Excel文件",
                    defaultextension=".xlsx",
                    filetypes=[("Excel文件", "*.xlsx")]
                )
                if file_path:
                    export_df.to_excel(file_path, index=False)
                    
            elif format_type == "csv":
                file_path = filedialog.asksaveasfilename(
                    title="保存CSV文件",
                    defaultextension=".csv",
                    filetypes=[("CSV文件", "*.csv")]
                )
                if file_path:
                    export_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                    
            elif format_type == "json":
                file_path = filedialog.asksaveasfilename(
                    title="保存JSON文件",
                    defaultextension=".json",
                    filetypes=[("JSON文件", "*.json")]
                )
                if file_path:
                    import json
                    # 转换为JSON数组格式
                    data_list = []
                    for _, row in export_df.iterrows():
                        row_dict = row.to_dict()
                        data_list.append(row_dict)
                    
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(data_list, f, ensure_ascii=False, indent=2)
                        
            elif format_type == "jsonl":
                file_path = filedialog.asksaveasfilename(
                    title="保存JSONL文件",
                    defaultextension=".jsonl",
                    filetypes=[("JSONL文件", "*.jsonl")]
                )
                if file_path:
                    import json
                    with open(file_path, 'w', encoding='utf-8') as f:
                        for _, row in export_df.iterrows():
                            row_dict = row.to_dict()
                            json_line = json.dumps(row_dict, ensure_ascii=False)
                            f.write(json_line + '\n')
            
            if 'file_path' in locals() and file_path:
                filename = os.path.basename(file_path)
                col_count = len(selected_columns)
                row_count = len(export_df)
                self.update_status(f"已导出 {col_count} 列 {row_count} 行到: {filename}", "success")
                messagebox.showinfo("成功", f"已导出 {col_count} 个字段，{row_count} 行数据到:\n{filename}")
                
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            self.update_status("导出失败", "error")
    
    def show_conditional_export(self):
        """显示条件筛选导出对话框"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
            
        try:
            from conditional_export_dialog import ConditionalExportDialog
            dialog = ConditionalExportDialog(self.root, self.table_manager)
            dialog.show()
        except Exception as e:
            messagebox.showerror("错误", f"打开条件筛选导出对话框失败: {e}")
            
    def show_find_replace(self):
        """显示查找替换对话框"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
            
        try:
            from find_replace_dialog import FindReplaceDialog
            dialog = FindReplaceDialog(self.root, self.table_manager)
            dialog.show()
            # 替换完成后刷新显示
            self.update_table_display()
        except Exception as e:
            messagebox.showerror("错误", f"打开查找替换对话框失败: {e}")
            
    def quick_export_excel(self):
        """快速导出Excel（使用上次的选择）"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        # 如果没有选择，使用所有列
        if not self.export_selection['selected_columns']:
            self.export_selection['selected_columns'] = self.table_manager.get_column_names()
        
        # 验证选中的列是否仍然存在
        current_columns = self.table_manager.get_column_names()
        valid_columns = [col for col in self.export_selection['selected_columns'] if col in current_columns]
        
        if not valid_columns:
            messagebox.showwarning("警告", "之前选择的字段已不存在，请重新选择")
            self.show_export_selection()
            return
            
        self.export_selected_columns(valid_columns, "excel")
                
    def show_help(self):
        """显示帮助"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("700x600")
        help_window.transient(self.root)
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=20, pady=20)
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        help_text = """
🚀 AI 批量数据处理工具使用说明

📋 **基本操作**：
1.  **新建表格**：创建一个空白的数据表格。
2.  **导入文件**：导入Excel、CSV或JSONL格式的现有数据。
3.  **添加列**：可添加普通列或AI处理列，支持智能定位插入。
4.  **编辑数据**：双击单元格或右键菜单进行内容编辑。
5.  **保存/加载项目**：支持保存和加载`.aie`项目文件，自动保存AI列配置和列宽信息。

🤖 **AI功能 - 三级处理模式**：
1.  **🔄 全部处理 (F5)**：处理所有AI列的所有行数据。
2.  **📋 单列处理 (F6)**：选择一个AI列处理其所有行。
3.  **⚡ 单元格处理 (F7)**：处理当前选中单元格或选中行的AI列。
4.  **新建/编辑AI列**：设置列名、AI模型（如gpt-4.1, o1）和Prompt模板。在Prompt中使用 `{列名}` 引用其他字段。

📁 **文件操作**：
-   支持导入 `.xlsx`、`.xls`、`.csv`、`.jsonl` 文件。
-   支持导出 `.xlsx`、`.csv`、`.jsonl`、`.json` 格式（可选择导出字段）。
-   自动处理中文编码。

⌨️ **快捷键**：
-   `Ctrl+N`: 新建空白表格
-   `Ctrl+O`: 打开项目
-   `Ctrl+S`: 保存项目
-   `Ctrl+Shift+S`: 另存为项目
-   `Ctrl+E`: 选择字段导出
-   `Ctrl+Shift+E`: 使用上次选择快速导出Excel
-   `F5`: 全部处理
-   `F6`: 单列处理
-   `F7`: 单元格处理

💡 **使用技巧**：
-   右键点击表格：根据选中内容（单元格、列头、空白区域）显示智能上下文菜单。
-   拖拽列头：可重新排列表格列的顺序。
-   AI列提示：AI列会在表格标题和导出选择中进行标注。
-   内容预览：选中单元格后，可在上方预览区查看完整内容并进行编辑、复制。
-   行高模式：提供低、中、高三种行高设置，提升阅读体验。
-   状态栏：实时显示操作进度和文件信息。

⚠️ **注意事项**：
-   确保 `.env` 文件中API配置正确。
-   批量处理大量数据时请耐心等待，系统已自动添加延迟。
"""
        
        text_widget.insert("1.0", help_text)
        text_widget.config(state=tk.DISABLED)
        
    def show_about(self):
        """显示关于"""
        messagebox.showinfo("关于", 
                          "🚀 AI Excel 批量数据处理工具 v2.0\n\n"
                          "一个用于批量AI数据处理的现代化桌面工具\n\n"
                          "✨ 功能特点:\n"
                          "• 现代化界面设计\n"
                          "• 空白表格创建\n"
                          "• AI批量数据处理\n"
                          "• 多格式文件支持\n"
                          "• 实时进度显示\n\n"
                          "💻 技术栈: Python + Tkinter + pandas + OpenAI")

    def on_cell_click(self, event):
        """单击单元格或列头处理 - 改进的选中逻辑"""
        try:
            # 识别点击区域
            clicked_column = self.tree.identify_column(event.x)
            clicked_item = self.tree.identify_row(event.y)
            
            # 重置选中信息
            self.selection_info = {
                'type': None,
                'row_index': None,
                'column_index': None,
                'column_name': None
            }
            
            df = self.table_manager.get_dataframe()
            if df is None:
                return
                
            # 如果点击列头
            if clicked_column and not clicked_item:
                col_index = int(clicked_column.replace('#', '')) - 1
                if 0 <= col_index < len(df.columns):
                    col_name = list(df.columns)[col_index]
                    
                    # 更新选中信息
                    self.selection_info = {
                        'type': 'column',
                        'row_index': None,
                        'column_index': col_index,
                        'column_name': col_name
                    }
                    
                    ai_columns = self.table_manager.get_ai_columns()
                    is_ai = col_name in ai_columns
                    col_type = "AI列" if is_ai else "普通列"
                    self.update_status(f"选中{col_type}: {col_name} (右键查看操作)", "normal")
                    
                    # 高亮选中的列
                    self.highlight_column(col_index)
                    
                    # 清空内容预览（因为选中的是列头，不是具体单元格）
                    self.clear_content_preview()
                    
            # 如果点击单元格
            elif clicked_column and clicked_item:
                col_index = int(clicked_column.replace('#', '')) - 1
                selection = self.tree.selection()
                
                if selection and 0 <= col_index < len(df.columns):
                    item = selection[0]
                    row_index = self.tree.index(item)
                    col_name = list(df.columns)[col_index]
                    
                    # 更新选中信息
                    self.selection_info = {
                        'type': 'cell',
                        'row_index': row_index,
                        'column_index': col_index,
                        'column_name': col_name
                    }
                    
                    # 获取单元格内容并更新预览
                    cell_content = df.iloc[row_index, col_index]
                    self.update_content_preview(row_index, col_name, cell_content)
                    
                    ai_columns = self.table_manager.get_ai_columns()
                    is_ai = col_name in ai_columns
                    cell_type = "AI单元格" if is_ai else "单元格"
                    self.update_status(f"选中{cell_type}: {col_name}[第{row_index+1}行] (双击编辑, 右键查看操作)", "normal")
                    
                    # 高亮选中单元格所在的列
                    self.highlight_column(col_index)
                    
        except Exception as e:
            print(f"选中处理错误: {e}")
            self.update_status("选中失败", "error")
            
    def highlight_column(self, col_index):
        """高亮指定列"""
        try:
            # 清除之前的高亮
            if self.highlighted_column is not None:
                self.clear_column_highlight()
            
            # 设置当前高亮列
            self.highlighted_column = col_index
            
            df = self.table_manager.get_dataframe()
            if df is not None and 0 <= col_index < len(df.columns):
                
                # 配置高亮标签样式 - 使用更明显的颜色
                highlight_tag = f"col_highlight_{col_index}"
                self.tree.tag_configure(highlight_tag, 
                                       background='#e0f2fe',  # 更明显的浅蓝背景
                                       foreground='#0f172a')  # 深色文字
                
                # 为该列的所有行添加高亮效果
                for item in self.tree.get_children():
                    # 清除之前的高亮标签
                    current_tags = list(self.tree.item(item, 'tags'))
                    current_tags = [tag for tag in current_tags if not tag.startswith('col_highlight_')]
                    
                    # 添加当前列的高亮标签
                    current_tags.append(highlight_tag)
                    self.tree.item(item, tags=current_tags)
                
                # 同时设置列头的高亮效果
                columns = list(df.columns)
                if col_index < len(columns):
                    col_name = columns[col_index]
                    ai_columns = self.table_manager.get_ai_columns()
                    
                    # 设置高亮的列头文字（保留AI图标）
                    display_col_name = col_name
                    if col_name in ai_columns:
                        display_col_name = f"🤖 {col_name}"
                    
                    # 添加星号表示选中状态
                    highlight_text = f"★ {display_col_name} ★"
                    
                    try:
                        self.tree.heading(col_name, text=highlight_text)
                        print(f"设置高亮列头: {col_name} -> {highlight_text}")
                    except Exception as e:
                        print(f"设置列头高亮错误: {e}")
                    
                print(f"成功高亮列 {col_index} ({col_name if col_index < len(columns) else 'unknown'})")
                    
        except Exception as e:
            print(f"列高亮错误: {e}")
            
    def clear_column_highlight(self):
        """清除列高亮"""
        try:
            if self.highlighted_column is not None:
                # 清除所有行的高亮标签
                for item in self.tree.get_children():
                    current_tags = list(self.tree.item(item, 'tags'))
                    current_tags = [tag for tag in current_tags if not tag.startswith('col_highlight_')]
                    self.tree.item(item, tags=current_tags)
                
                # 恢复列头文字（移除星号，并正确恢复AI列图标）
                df = self.table_manager.get_dataframe()
                if df is not None and 0 <= self.highlighted_column < len(df.columns):
                    columns = list(df.columns)
                    col_name = columns[self.highlighted_column]
                    ai_columns = self.table_manager.get_ai_columns()
                    
                    # 根据是否为AI列设置正确的显示名称
                    display_col_name = col_name
                    if col_name in ai_columns:
                        display_col_name = f"🤖 {col_name}"  # 恢复AI列图标
                        
                    try:
                        # 恢复原始列头文字（包含AI图标）
                        self.tree.heading(col_name, text=display_col_name)
                        print(f"恢复列头: {col_name} -> {display_col_name}")
                    except Exception as e:
                        print(f"恢复列头文本错误: {e}")
                    
                self.highlighted_column = None
                print("已清除列高亮")
        except Exception as e:
            print(f"清除列高亮错误: {e}")

    def process_selected_row(self, row_index):
        """处理选中行的所有AI列"""
        ai_columns = self.table_manager.get_ai_columns()
        self.process_selected_row_columns(row_index, ai_columns)
        
    def process_selected_row_columns(self, row_index, ai_columns):
        """处理选中行的指定AI列"""
        if not ai_columns:
            return
            
        try:
            self.update_status(f"正在处理第{row_index+1}行的AI列...", "normal")
            
            success_count = 0
            total_count = len(ai_columns)
            
            # 处理每个AI列
            for column_name, prompt_template in ai_columns.items():
                try:
                    success, result = self.ai_processor.process_single_cell(
                        self.table_manager.get_dataframe(),
                        row_index,
                        column_name,
                        prompt_template
                    )
                    
                    if success:
                        success_count += 1
                        
                    # 更新显示
                    self.update_table_display()
                    self.root.update()
                    
                    # 添加延迟
                    time.sleep(0.1)
                    
                except Exception as e:
                    print(f"处理列 {column_name} 时出错: {e}")
                    
            # 完成提示
            if success_count == total_count:
                self.update_status(f"第{row_index+1}行AI处理完成 ({success_count}/{total_count})", "success")
                messagebox.showinfo("成功", f"第{row_index+1}行的{success_count}个AI列处理完成！")
            else:
                self.update_status(f"第{row_index+1}行AI处理完成 ({success_count}/{total_count})", "error")
                messagebox.showwarning("部分成功", f"第{row_index+1}行：{success_count}/{total_count}个AI列处理成功")
                
        except Exception as e:
            messagebox.showerror("错误", f"处理AI列时出错: {str(e)}")
            self.update_status("AI处理失败", "error")

    def on_cell_double_click(self, event):
        """双击单元格编辑"""
        try:
            selection = self.tree.selection()
            if not selection:
                self.update_status("请先选择一个单元格", "error")
                return
                
            item = selection[0]
            column = self.tree.identify_column(event.x)
            
            if not column:
                self.update_status("无法识别列位置", "error")
                return
                
            # 获取列索引和列名
            col_index = int(column.replace('#', '')) - 1
            if col_index < 0:
                return
                
            df = self.table_manager.get_dataframe()
            if df is None:
                self.update_status("没有数据可编辑", "error")
                return
                
            column_names = list(df.columns)
            if col_index >= len(column_names):
                return
                
            col_name = column_names[col_index]
            
            # 获取当前值
            values = self.tree.item(item, 'values')
            if col_index < len(values):
                current_value = values[col_index]
                # 处理被截断的文本，从原始数据获取完整值
                row_index = self.tree.index(item)
                current_value = str(df.iloc[row_index, col_index])
            else:
                current_value = ""
                
            # 获取行索引
            row_index = self.tree.index(item)
            
            # 创建编辑对话框
            self.edit_cell_dialog(row_index, col_name, current_value)
            
        except Exception as e:
            self.update_status(f"编辑失败: {str(e)}", "error")
            print(f"双击编辑错误: {e}")

    def insert_column_at_position(self, position, direction):
        """在指定位置插入新列"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "请先创建表格或导入数据文件")
            return
            
        # 计算实际插入位置
        actual_position = position  # 使用传入的position参数
        # Ensure self.selection_info and 'column_index' exist before accessing
        current_col_index = None
        if hasattr(self, 'selection_info') and self.selection_info:
            current_col_index = self.selection_info.get('column_index')

        if current_col_index is not None:
            col_index = current_col_index
            actual_position = col_index if direction == "left" else col_index + 1
        else:
            # 如果没有选中，根据方向决定插入开头还是末尾
            df = self.table_manager.get_dataframe()
            if df is not None:
                actual_position = 0 if direction == "left" else len(df.columns)
            # If df is None, actual_position remains as passed position

        # 询问用户要插入哪种类型的列
        column_type_choice = self.ask_column_type_for_insertion()
        if column_type_choice is None: # 用户取消
            return

        column_name = None
        success = False # Initialize success to False
        side = "左" if direction == "left" else "右"

        if column_type_choice == "ai":
            from ai_column_dialog import AIColumnDialog
            dialog = AIColumnDialog(self.root, self.table_manager.get_column_names())
            result = dialog.show()
            if result:
                column_name = result['column_name']
                prompt_template = result['prompt']
                ai_model = result['model']
                processing_params = result['processing_params']
                output_mode = result.get('output_mode', 'single')
                output_fields = result.get('output_fields', [])
                if column_name in self.table_manager.get_column_names():
                    messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
                    return
                success = self.table_manager.insert_column_at_position(
                    actual_position, column_name, prompt_template=prompt_template, is_ai_column=True, 
                    ai_model=ai_model, processing_params=processing_params, output_mode=output_mode, 
                    output_fields=output_fields
                )
                if success:
                    if output_mode == "multi" and output_fields:
                        self.update_status(f"已在{side}侧插入多字段AI列: {column_name} (字段: {', '.join(output_fields)})", "success")
                    else:
                        self.update_status(f"已在{side}侧插入AI列: {column_name}", "success")
                    self.update_table_display()
                else:
                    messagebox.showerror("错误", "插入AI列失败")

        elif column_type_choice == "long_text":
            from long_text_column_dialog import LongTextColumnDialog
            dialog = LongTextColumnDialog(self.root, self.table_manager, lambda r: None)
            dialog.show()
            result = dialog.result
            if result:
                column_name = result['column_name']
                filename_field = result['filename_field']
                folder_path = result['folder_path']
                preview_length = result['preview_length']

                if column_name in self.table_manager.get_column_names():
                    messagebox.showerror("错误", f"列名 '{column_name}' 已存在")
                    return
                
                # 先插入普通列，再转换为长文本列
                success = self.table_manager.insert_column_at_position(actual_position, column_name)
                if success:
                    # 添加长文本列配置
                    self.table_manager.long_text_columns[column_name] = {
                        "filename_field": filename_field,
                        "folder_path": folder_path,
                        "preview_length": preview_length
                    }
                    # 刷新长文本列内容
                    self.table_manager.refresh_long_text_column(column_name)
                    self.update_status(f"已在{side}侧插入长文本列: {column_name}", "success")
                    self.update_table_display()
                else:
                    messagebox.showerror("错误", "插入长文本列失败")

        elif column_type_choice == "normal": # Normal column
            # 调用 create_normal_column，它内部会处理创建方式（手动或JSONL）
            # 它也需要知道插入的位置 actual_position 和 side (用于状态消息)
            self.create_normal_column(position=actual_position, side=side)

    def ask_column_type_for_insertion(self):
        """询问用户要插入的列类型"""
        dialog = tk.Toplevel(self.root)
        dialog.title("选择列类型")
        dialog.geometry("300x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"300x200+{x}+{y}")

        column_type_var = tk.StringVar(value="ai") # Default to AI column
        result = [None] # Use a list to store the result

        ttk.Label(dialog, text="请选择要插入的列类型:", style='Title.TLabel').pack(pady=10)

        ttk.Radiobutton(dialog, text="AI处理列", variable=column_type_var, value="ai").pack(anchor=tk.W, padx=20)
        ttk.Radiobutton(dialog, text="长文本列", variable=column_type_var, value="long_text").pack(anchor=tk.W, padx=20)
        ttk.Radiobutton(dialog, text="普通列", variable=column_type_var, value="normal").pack(anchor=tk.W, padx=20)

        def on_ok():
            result[0] = column_type_var.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)

        dialog.wait_window()
        return result[0]

    def on_column_drag_start(self, event):
        """开始列拖拽"""
        # 检查是否点击在列头区域
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            column = self.tree.identify_column(event.x)
            if column:
                self.drag_data['dragging'] = True
                self.drag_data['start_column'] = column
                self.drag_data['start_x'] = event.x
                self.drag_data['target_column'] = column
                
                # 改变光标样式
                self.tree.config(cursor="hand2")
                
                # 获取列信息用于视觉反馈
                col_index = int(column.replace('#', '')) - 1
                df = self.table_manager.get_dataframe()
                if df is not None and col_index < len(df.columns):
                    col_name = list(df.columns)[col_index]
                    self.update_status(f"正在拖拽列: {col_name}", "normal")
                
    def on_column_drag_motion(self, event):
        """列拖拽移动过程"""
        if self.drag_data['dragging']:
            # 识别当前位置的列
            column = self.tree.identify_column(event.x)
            if column and column != self.drag_data['target_column']:
                self.drag_data['target_column'] = column
                
                # 提供视觉反馈
                col_index = int(column.replace('#', '')) - 1
                df = self.table_manager.get_dataframe()
                if df is not None and col_index < len(df.columns):
                    target_col_name = list(df.columns)[col_index]
                    self.update_status(f"目标位置: {target_col_name}", "normal")
                
                # 改变目标列的视觉样式（高亮效果）
                self.tree.configure(cursor="exchange")
                
    def on_column_drag_end(self, event):
        """结束列拖拽"""
        if self.drag_data['dragging']:
            start_col = self.drag_data['start_column']
            target_col = self.drag_data['target_column']
            
            # 恢复光标
            self.tree.config(cursor="")
            
            # 如果拖拽到了不同的列，执行移动
            if start_col and target_col and start_col != target_col:
                self.move_column_with_animation(start_col, target_col)
            else:
                self.update_status("就绪", "normal")
            
        # 重置拖拽状态
        self.drag_data = {
            'dragging': False,
            'start_column': None,
            'start_x': 0,
            'target_column': None,
            'drag_indicator': None
        }
        
    def move_column_with_animation(self, from_column, to_column):
        """带动画效果的移动列位置"""
        try:
            # 转换列标识为索引
            from_index = int(from_column.replace('#', '')) - 1
            to_index = int(to_column.replace('#', '')) - 1
            
            df = self.table_manager.get_dataframe()
            if df is not None and 0 <= from_index < len(df.columns) and 0 <= to_index < len(df.columns):
                columns = list(df.columns)
                from_col_name = columns[from_index]
                to_col_name = columns[to_index]
                
                # 显示移动动画效果（通过状态更新）
                self.update_status(f"移动中: {from_col_name} → {to_col_name}", "normal")
                self.root.update()
                
                # 执行列移动
                success = self.table_manager.move_column(from_index, to_index)
                if success:
                    # 延迟更新以显示动画效果
                    self.root.after(100, lambda: self.update_table_display())
                    self.root.after(200, lambda: self.update_status(f"已移动列: {from_col_name} → {to_col_name}位置", "success"))
                    
        except Exception as e:
            print(f"移动列失败: {e}")
            self.update_status("移动列失败", "error")

    def delete_selected_row(self, row_index):
        """删除选中的行"""
        if self.table_manager.get_dataframe() is None:
            messagebox.showwarning("警告", "没有数据表格")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None:
            return
            
        # 确认删除
        result = messagebox.askyesno("确认删除", 
                                   f"确定要删除第 {row_index + 1} 行吗？\n\n"
                                   f"此操作将删除该行的所有数据。")
        if result:
            # 保存状态用于撤销
            self.save_state(f"删除第{row_index+1}行")
            
            # 执行删除
            success = self.table_manager.delete_row(row_index)
            if success:
                self.update_table_display()
                self.update_title()  # 更新标题栏
                self.update_status(f"已删除第 {row_index + 1} 行", "success")
            else:
                messagebox.showerror("错误", "删除行失败")

    def change_row_height(self):
        """更改行高"""
        new_height = self.row_height_var.get()
        if new_height in self.row_height_settings:
            self.current_row_height = new_height
            # 通过重新配置样式来更改行高
            style = ttk.Style()
            style.configure('Modern.Treeview',
                           rowheight=self.row_height_settings[self.current_row_height])
            self.update_status(f"已更改行高: {self.current_row_height}", "success")
        else:
            messagebox.showerror("错误", "无效的行高选择")
            
    def set_row_height(self, height):
        """设置行高（工具栏按钮使用）"""
        if height in self.row_height_settings:
            self.current_row_height = height
            self.row_height_var.set(height)  # 同步菜单选择
            
            # 通过重新配置样式来更改行高
            style = ttk.Style()
            style.configure('Modern.Treeview',
                           background='#ffffff',
                           foreground='#1a202c',
                           selectbackground='#dbeafe',
                           selectforeground='#1e40af',
                           fieldbackground='#ffffff',
                           bordercolor='#d1d5db',
                           borderwidth=1,
                           font=('Arial', 10),
                           rowheight=self.row_height_settings[self.current_row_height])
            
            # 获取中文描述
            height_names = {'low': '低 (紧凑)', 'medium': '中 (标准)', 'high': '高 (宽松)'}
            self.update_status(f"已设置行高: {height_names[height]}", "success")
        else:
            messagebox.showerror("错误", "无效的行高选择")

    def process_all_ai(self):
        """全部处理 - 处理所有AI列的所有行"""
        ai_columns = self.table_manager.get_ai_columns()
        if not ai_columns:
            messagebox.showwarning("警告", "没有AI列需要处理")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None or len(df) == 0:
            messagebox.showwarning("警告", "没有数据需要处理")
            return
            
        # 考虑筛选状态
        if self.filter_state['active'] and self.filter_state['filtered_indices']:
            # 如果有筛选，只处理筛选后的行
            rows_to_process = self.filter_state['filtered_indices']
            data_description = f"筛选后的 {len(rows_to_process)} 行"
        else:
            # 处理所有行
            rows_to_process = list(range(len(df)))
            data_description = f"所有 {len(df)} 行"
        
        # 确认处理
        total_tasks = len(rows_to_process) * len(ai_columns)
        result = messagebox.askyesno("确认全部处理", 
                                   f"即将处理所有 {len(ai_columns)} 个AI列的{data_description}数据。\n"
                                   f"总共 {total_tasks} 个任务，这可能需要较长时间。\n\n"
                                   f"是否继续？")
        if not result:
            return
            
        try:
            self.update_status("正在全部处理AI列...", "normal")
            
            success_count = 0
            current_task = 0
            
            # 处理每个AI列的筛选行，优化性能
            for col_name, prompt_template in ai_columns.items():
                for row_index in rows_to_process:
                    try:
                        success, result = self.ai_processor.process_single_cell(
                            df, row_index, col_name, prompt_template
                        )
                        
                        if success:
                            success_count += 1
                            
                        current_task += 1
                        
                        # 更新表格进度条
                        self.update_table_progress(current_task, total_tasks, "全部处理")
                        
                        # 减少界面更新频率，每10个任务更新一次显示
                        if current_task % 10 == 0 or current_task == total_tasks:
                            self.update_table_display()
                        
                        # 减少延迟，提高处理速度
                        if current_task % 5 == 0:
                            time.sleep(0.05)  # 减少延迟
                        
                    except Exception as e:
                        print(f"处理 {col_name} 第{row_index+1}行时出错: {e}")
                        current_task += 1
                        
            # 最终更新显示
            self.update_table_display()
            self.update_status(f"全部处理完成 ({success_count}/{total_tasks})", "success")
            messagebox.showinfo("完成", f"全部处理完成！\n成功: {success_count}/{total_tasks}")
            
        except Exception as e:
            messagebox.showerror("错误", f"全部处理时出错: {str(e)}")
            self.update_status("全部处理失败", "error")

    def process_single_column(self):
        """单列处理 - 选择一个AI列处理所有行"""
        ai_columns = self.table_manager.get_ai_columns()
        if not ai_columns:
            messagebox.showwarning("警告", "没有AI列需要处理")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None or len(df) == 0:
            messagebox.showwarning("警告", "没有数据需要处理")
            return
            
        # 选择要处理的AI列
        if len(ai_columns) == 1:
            # 只有一个AI列，直接处理
            col_name = list(ai_columns.keys())[0]
        else:
            # 多个AI列，让用户选择
            dialog = tk.Toplevel(self.root)
            dialog.title("选择AI列")
            dialog.geometry("400x300")
            dialog.transient(self.root)
            dialog.grab_set()
            
            # 居中显示
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
            y = (dialog.winfo_screenheight() // 2) - (300 // 2)
            dialog.geometry(f"400x300+{x}+{y}")
            
            ttk.Label(dialog, text="选择要处理的AI列:", style='Title.TLabel').pack(pady=10)
            
            # 列表框
            listbox_frame = ttk.Frame(dialog)
            listbox_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            listbox = tk.Listbox(listbox_frame, selectmode=tk.SINGLE)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            listbox.configure(yscrollcommand=scrollbar.set)
            
            # 添加AI列到列表
            for col in ai_columns.keys():
                listbox.insert(tk.END, col)
                
            selected_column = [None]  # 使用列表来存储选择结果
            
            def on_select():
                selection = listbox.curselection()
                if not selection:
                    messagebox.showwarning("警告", "请选择一个AI列")
                    return
                    
                selected_column[0] = list(ai_columns.keys())[selection[0]]
                dialog.destroy()
                
            def on_cancel():
                dialog.destroy()
                
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="确定", command=on_select).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)
            
            # 等待对话框关闭
            dialog.wait_window()
            
            col_name = selected_column[0]
            if not col_name:
                return
                
        # 考虑筛选状态
        if self.filter_state['active'] and self.filter_state['filtered_indices']:
            # 如果有筛选，只处理筛选后的行
            rows_to_process = self.filter_state['filtered_indices']
            data_description = f"筛选后的 {len(rows_to_process)} 行"
        else:
            # 处理所有行
            rows_to_process = list(range(len(df)))
            data_description = f"所有 {len(df)} 行"
        
        # 确认处理
        result = messagebox.askyesno("确认单列处理", 
                                   f"即将处理AI列 '{col_name}' 的{data_description}数据。\n\n"
                                   f"是否继续？")
        if not result:
            return
            
        try:
            self.update_status(f"正在处理列 {col_name}...", "normal")
            
            success_count = 0
            prompt_template = ai_columns[col_name]
            
            # 处理选中列的筛选行，优化性能
            for i, row_index in enumerate(rows_to_process):
                try:
                    success, result = self.ai_processor.process_single_cell(
                        df, row_index, col_name, prompt_template
                    )
                    
                    if success:
                        success_count += 1
                        
                    # 更新表格进度条
                    self.update_table_progress(i + 1, len(rows_to_process), f"处理列 {col_name}")
                    
                    # 减少界面更新频率，每3行更新一次显示
                    if (i + 1) % 3 == 0 or i == len(rows_to_process) - 1:
                        self.update_table_display()
                    
                    # 减少延迟
                    if (i + 1) % 3 == 0:
                        time.sleep(0.05)
                    
                except Exception as e:
                    print(f"处理列 {col_name} 第{row_index+1}行时出错: {e}")
                    
            # 最终更新显示
            self.update_table_display()
            self.update_status(f"列 {col_name} 处理完成 ({success_count}/{len(rows_to_process)})", "success")
            messagebox.showinfo("完成", f"列 '{col_name}' 处理完成！\n成功: {success_count}/{len(rows_to_process)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"单列处理时出错: {str(e)}")
            self.update_status("单列处理失败", "error")

    def process_single_cell(self):
        """单元格处理 - 处理当前选中的单元格"""
        # 检查选中状态
        if self.selection_info['type'] == 'cell':
            # 如果选中了单元格，处理该单元格
            row_index = self.selection_info['row_index']
            col_name = self.selection_info['column_name']
            ai_columns = self.table_manager.get_ai_columns()
            
            if col_name in ai_columns:
                self.process_specific_cell(row_index, col_name)
            else:
                messagebox.showwarning("警告", f"选中的单元格 '{col_name}' 不是AI列")
            return
            
        elif self.selection_info['type'] == 'column':
            # 如果选中了列，处理整列
            col_name = self.selection_info['column_name']
            ai_columns = self.table_manager.get_ai_columns()
            
            if col_name in ai_columns:
                self.process_entire_column(col_name)
            else:
                messagebox.showwarning("警告", f"选中的列 '{col_name}' 不是AI列")
            return
        
        # 如果没有选中，使用原来的逻辑
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个单元格或列")
            return
            
        ai_columns = self.table_manager.get_ai_columns()
        if not ai_columns:
            messagebox.showwarning("警告", "没有AI列需要处理")
            return
            
        df = self.table_manager.get_dataframe()
        if df is None:
            messagebox.showwarning("警告", "没有数据需要处理")
            return
            
        # 获取选中行的索引
        item = selection[0]
        row_index = self.tree.index(item)
        
        # 检查该行是否有AI列需要处理
        row_ai_columns = {}
        for col_name, prompt_template in ai_columns.items():
            if col_name in df.columns:
                row_ai_columns[col_name] = prompt_template
                
        if not row_ai_columns:
            messagebox.showwarning("警告", "选中行没有AI列需要处理")
            return
            
        # 如果只有一个AI列，直接处理
        if len(row_ai_columns) == 1:
            col_name = list(row_ai_columns.keys())[0]
            prompt_template = row_ai_columns[col_name]
            
            # 确认处理
            result = messagebox.askyesno("确认单元格处理", 
                                       f"即将处理第 {row_index+1} 行的AI列 '{col_name}'。\n\n"
                                       f"是否继续？")
            if not result:
                return
                
            try:
                self.update_status(f"正在处理单元格 {col_name}[{row_index+1}]...", "normal")
                
                success, result = self.ai_processor.process_single_cell(
                    df, row_index, col_name, prompt_template
                )
                
                if success:
                    self.update_table_display()
                    self.update_status(f"单元格 {col_name}[{row_index+1}] 处理完成", "success")
                    messagebox.showinfo("完成", f"单元格处理完成！\n列: {col_name}\n行: {row_index+1}")
                else:
                    self.update_status("单元格处理失败", "error")
                    messagebox.showerror("错误", "单元格处理失败")
                    
            except Exception as e:
                messagebox.showerror("错误", f"单元格处理时出错: {str(e)}")
                self.update_status("单元格处理失败", "error")
                
        else:
            # 多个AI列，让用户选择或处理全部
            dialog = tk.Toplevel(self.root)
            dialog.title("选择处理方式")
            dialog.geometry("450x350")
            dialog.transient(self.root)
            dialog.grab_set()
            
            # 居中显示
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
            y = (dialog.winfo_screenheight() // 2) - (350 // 2)
            dialog.geometry(f"450x350+{x}+{y}")
            
            ttk.Label(dialog, text=f"第 {row_index+1} 行有多个AI列，请选择处理方式:", 
                     style='Title.TLabel').pack(pady=10)
            
            # 选项框架
            option_frame = ttk.LabelFrame(dialog, text="处理选项", padding=10)
            option_frame.pack(fill=tk.X, padx=10, pady=10)
            
            process_option = tk.StringVar(value="all")
            
            ttk.Radiobutton(option_frame, text=f"处理该行所有AI列 ({len(row_ai_columns)}个)", 
                           variable=process_option, value="all").pack(anchor='w', pady=5)
            ttk.Radiobutton(option_frame, text="选择特定AI列处理", 
                           variable=process_option, value="select").pack(anchor='w', pady=5)
            
            # AI列列表框架
            list_frame = ttk.LabelFrame(dialog, text="AI列列表", padding=10)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE)
            listbox.pack(fill=tk.BOTH, expand=True)
            
            for col in row_ai_columns.keys():
                listbox.insert(tk.END, col)
                
            selected_result = [None]  # 存储选择结果
            
            def on_process():
                option = process_option.get()
                
                if option == "all":
                    # 处理所有AI列
                    selected_result[0] = ("all", list(row_ai_columns.keys()))
                else:
                    # 处理选中的AI列
                    selection = listbox.curselection()
                    if not selection:
                        messagebox.showwarning("警告", "请选择一个AI列")
                        return
                        
                    selected_col = list(row_ai_columns.keys())[selection[0]]
                    selected_result[0] = ("single", [selected_col])
                    
                dialog.destroy()
                
            def on_cancel():
                dialog.destroy()
                
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="开始处理", command=on_process).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)
            
            # 等待对话框关闭
            dialog.wait_window()
            
            if not selected_result[0]:
                return
                
            process_type, columns_to_process = selected_result[0]
            
            try:
                self.update_status(f"正在处理第 {row_index+1} 行的AI列...", "normal")
                
                success_count = 0
                total_count = len(columns_to_process)
                
                for col_name in columns_to_process:
                    try:
                        prompt_template = row_ai_columns[col_name]
                        success, result = self.ai_processor.process_single_cell(
                            df, row_index, col_name, prompt_template
                        )
                        
                        if success:
                            success_count += 1
                            
                        # 添加延迟
                        time.sleep(0.1)
                        
                    except Exception as e:
                        print(f"处理单元格 {col_name}[{row_index+1}] 时出错: {e}")
                        
                # 更新显示
                self.update_table_display()
                self.update_status(f"第 {row_index+1} 行处理完成 ({success_count}/{total_count})", "success")
                messagebox.showinfo("完成", f"第 {row_index+1} 行处理完成！\n成功: {success_count}/{total_count}")
                
            except Exception as e:
                messagebox.showerror("错误", f"单元格处理时出错: {str(e)}")
                self.update_status("单元格处理失败", "error")

    def _get_column_widths(self):
        """获取当前Treeview中所有列的宽度"""
        column_widths = {}
        if self.tree["columns"]:
            for col_id in self.tree["columns"]:
                column_widths[col_id] = self.tree.column(col_id, "width")
        return column_widths

    def _apply_column_widths(self, column_widths):
        """应用保存的列宽设置"""
        try:
            for col, width in column_widths.items():
                if self.tree.exists(col) or col in self.tree['columns']:
                    self.tree.column(col, width=width)
        except Exception as e:
            print(f"应用列宽设置失败: {e}")


        
    def sort_by_column(self, column, ascending=True):
        """按指定列排序"""
        try:
            df = self.table_manager.get_dataframe()
            if df is None or df.empty:
                return
                
            # 保存原始顺序（如果还没有保存的话）
            if self.sort_state['original_order'] is None:
                self.sort_state['original_order'] = df.index.tolist()
            
            # 更新排序状态
            self.sort_state['column'] = column
            self.sort_state['ascending'] = ascending
                
            # 创建排序后的数据框
            sorted_df = df.sort_values(by=column, ascending=ascending, na_position='last')
            
            # 更新表格管理器中的数据
            self.table_manager.dataframe = sorted_df.reset_index(drop=True)
            
            # 重新显示表格
            self.update_table_display()
            
            # 更新状态信息
            direction = "升序" if ascending else "降序"
            self.update_status(f"已按 {column} 列{direction}排序", "success")
            
        except Exception as e:
            print(f"排序失败: {e}")
            self.update_status(f"排序失败: {str(e)}", "error")
    
    def get_sort_indicator(self, column):
        """获取排序指示符"""
        if self.sort_state['column'] == column:
            return " ↑" if self.sort_state['ascending'] else " ↓"
        return ""
    
    def reset_sort(self):
        """重置排序到原始顺序"""
        if self.sort_state['original_order'] is not None:
            df = self.table_manager.get_dataframe()
            if df is not None:
                # 恢复原始顺序
                original_df = df.iloc[self.sort_state['original_order']].reset_index(drop=True)
                self.table_manager.dataframe = original_df
                
                # 重置排序状态
                self.sort_state = {
                    'column': None,
                    'ascending': True,
                    'original_order': None
                }
                
                # 更新显示
                self.update_table_display()
                self.update_status("已重置为原始顺序", "success")

    def on_closing(self):
        """处理窗口关闭事件，提示保存"""
        if self.table_manager.has_unsaved_changes():
            response = messagebox.askyesnocancel(
                "退出确认", 
                "有未保存的更改，您想保存吗？"
            )
            if response is True:  # 用户选择"是" (保存)
                success, _ = self.save_project() # 调用save_project
                if success: # 只有在保存成功后才退出
                    self.root.quit()
            elif response is False: # 用户选择"否" (不保存)
                self.root.quit()
            # 如果用户选择"取消"，则不执行任何操作，窗口不关闭
        else:
            self.root.quit()

    def toggle_full_prompt_display(self):
        """切换完整Prompt的显示/隐藏"""
        if self.show_full_prompt_var.get():
            self.prompt_text.config(height=10) # 展开高度
        else:
            self.prompt_text.config(height=3) # 默认高度

    def show_ai_column_progress(self, col_name):
        """显示AI列处理进度对话框"""
        df = self.table_manager.get_dataframe()
        if df is None:
            messagebox.showwarning("警告", "没有数据")
            return
            
        # 获取处理状态
        status = self.ai_processor.get_column_processing_status(df, col_name)
        if status is None:
            messagebox.showwarning("警告", f"列 '{col_name}' 不存在")
            return
            
        # 创建进度查询对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"AI列处理进度 - {col_name}")
        dialog.geometry("600x500")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"600x500+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text=f"AI列处理进度分析 - {col_name}", 
                               font=('Microsoft YaHei UI', 14, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # 统计信息框架
        stats_frame = ttk.LabelFrame(main_frame, text="处理统计", padding="10")
        stats_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 统计信息网格
        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack(fill=tk.X)
        
        # 第一行统计
        row1 = ttk.Frame(stats_grid)
        row1.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(row1, text="总行数:", font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(row1, text=str(status['total_rows']), foreground="blue").pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(row1, text="已处理:", font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(row1, text=str(status['processed_count']), foreground="green").pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(row1, text="完成率:", font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(row1, text=f"{status['completion_rate']}%", foreground="green").pack(side=tk.LEFT, padx=(5, 0))
        
        # 第二行统计
        row2 = ttk.Frame(stats_grid)
        row2.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(row2, text="失败:", font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(row2, text=str(status['failed_count']), foreground="red").pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(row2, text="空白:", font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Label(row2, text=str(status['empty_count']), foreground="orange").pack(side=tk.LEFT, padx=(5, 20))
        
        # 进度条
        progress_frame = ttk.Frame(stats_grid)
        progress_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(progress_frame, text="进度:", font=('Microsoft YaHei UI', 10, 'bold')).pack(side=tk.LEFT)
        progress_bar = ttk.Progressbar(progress_frame, length=300, mode='determinate')
        progress_bar.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
        progress_bar['value'] = status['completion_rate']
        
        # 任务选择框架
        task_frame = ttk.LabelFrame(main_frame, text="选择要处理的任务", padding="10")
        task_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 任务选择选项
        task_var = tk.StringVar(value="failed_and_empty")
        
        ttk.Radiobutton(task_frame, text=f"仅处理失败的任务 ({status['failed_count']}行)", 
                       variable=task_var, value="failed").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(task_frame, text=f"仅处理空白的任务 ({status['empty_count']}行)", 
                       variable=task_var, value="empty").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(task_frame, text=f"处理失败和空白的任务 ({status['failed_count'] + status['empty_count']}行)", 
                       variable=task_var, value="failed_and_empty").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(task_frame, text=f"重新处理所有任务 ({status['total_rows']}行)", 
                       variable=task_var, value="all").pack(anchor=tk.W, pady=2)
        
        # 详细信息框架（可选显示）
        details_frame = ttk.LabelFrame(main_frame, text="详细信息", padding="10")
        details_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 显示失败行和空白行的详细信息
        details_text = tk.Text(details_frame, height=6, wrap=tk.WORD)
        details_text.pack(fill=tk.X)
        
        details_content = ""
        if status['failed_rows']:
            details_content += f"失败行号: {', '.join(map(str, status['failed_rows'][:20]))}"
            if len(status['failed_rows']) > 20:
                details_content += f" (共{len(status['failed_rows'])}行，仅显示前20行)"
            details_content += "\n\n"
            
        if status['empty_rows']:
            details_content += f"空白行号: {', '.join(map(str, status['empty_rows'][:20]))}"
            if len(status['empty_rows']) > 20:
                details_content += f" (共{len(status['empty_rows'])}行，仅显示前20行)"
                
        details_text.insert("1.0", details_content)
        details_text.config(state=tk.DISABLED)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def start_processing():
            """开始处理选中的任务"""
            task_type = task_var.get()
            
            # 确定要处理的行
            rows_to_process = []
            if task_type == "failed":
                rows_to_process = [row - 1 for row in status['failed_rows']]  # 转为0-based索引
            elif task_type == "empty":
                rows_to_process = [row - 1 for row in status['empty_rows']]
            elif task_type == "failed_and_empty":
                rows_to_process = [row - 1 for row in status['failed_rows'] + status['empty_rows']]
            elif task_type == "all":
                rows_to_process = list(range(status['total_rows']))
                
            if not rows_to_process:
                messagebox.showinfo("提示", "没有需要处理的任务")
                return
                
            # 确认处理
            task_count = len(rows_to_process)
            result = messagebox.askyesno("确认处理", 
                                       f"即将处理 {task_count} 个任务。\n\n"
                                       f"是否继续？")
            if not result:
                return
                
            dialog.destroy()
            
            # 开始处理
            self.process_ai_column_concurrent(col_name, rows_to_process)
            
        def refresh_status():
            """刷新状态"""
            dialog.destroy()
            self.show_ai_column_progress(col_name)
            
        def close_dialog():
            """关闭对话框"""
            dialog.destroy()
            
        # 按钮
        ttk.Button(button_frame, text="🔄 刷新状态", command=refresh_status).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="▶️ 开始处理", command=start_processing).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="❌ 关闭", command=close_dialog).pack(side=tk.RIGHT)
        
        # 绑定快捷键
        dialog.bind('<Escape>', lambda e: close_dialog())
        dialog.bind('<F5>', lambda e: refresh_status())

    def process_ai_column_concurrent(self, col_name, rows_to_process=None):
        """并发处理AI列"""
        ai_columns = self.table_manager.get_ai_columns()
        if col_name not in ai_columns:
            messagebox.showwarning("警告", f"'{col_name}' 不是AI列")
            return
        
        df = self.table_manager.get_dataframe()
        if df is None:
            return
            
        # 获取AI列配置
        config = ai_columns[col_name]
        if isinstance(config, dict):
            prompt_template = config.get("prompt", "")
            model = config.get("model", "gpt-4.1")
            output_mode = config.get("output_mode", "single")
            output_fields = config.get("output_fields", [])
            processing_params = config.get("processing_params", {
                'max_workers': 3,
                'request_delay': 0.5,
                'max_retries': 2
            })
        else:
            # 向后兼容
            prompt_template = config
            model = "gpt-4.1"
            output_mode = "single"
            output_fields = []
            processing_params = {
                'max_workers': 3,
                'request_delay': 0.5,
                'max_retries': 2
            }
        
        # 如果未指定处理行，处理所有行（考虑筛选状态）
        if rows_to_process is None:
            if self.filter_state['active'] and self.filter_state['filtered_indices']:
                # 如果有筛选，只处理筛选后的行
                rows_to_process = self.filter_state['filtered_indices']
            else:
                # 否则处理所有行
                rows_to_process = list(range(len(df)))
            
        task_count = len(rows_to_process)
        if task_count == 0:
            messagebox.showinfo("提示", "没有需要处理的任务")
            return
            
        # 创建处理控制对话框
        control_dialog = tk.Toplevel(self.root)
        control_dialog.title(f"正在处理 - {col_name}")
        control_dialog.geometry("500x300")
        control_dialog.resizable(False, False)
        control_dialog.transient(self.root)
        control_dialog.grab_set()
        
        # 居中显示
        control_dialog.update_idletasks()
        x = (control_dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (control_dialog.winfo_screenheight() // 2) - (300 // 2)
        control_dialog.geometry(f"500x300+{x}+{y}")
        
        # 控制框架
        control_frame = ttk.Frame(control_dialog, padding="20")
        control_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(control_frame, text=f"正在处理AI列: {col_name}", 
                               font=('Microsoft YaHei UI', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # 配置信息
        config_frame = ttk.LabelFrame(control_frame, text="处理配置", padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        config_info = tk.Text(config_frame, height=4, wrap=tk.WORD)
        config_info.pack(fill=tk.X)
        config_text = f"模型: {model}\n"
        config_text += f"任务数: {task_count}\n"
        config_text += f"并发数: {processing_params['max_workers']}\n"
        config_text += f"请求延迟: {processing_params['request_delay']}秒"
        config_info.insert("1.0", config_text)
        config_info.config(state=tk.DISABLED)
        
        # 进度信息
        progress_frame = ttk.LabelFrame(control_frame, text="处理进度", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 进度标签
        progress_label = ttk.Label(progress_frame, text="准备开始处理...")
        progress_label.pack(pady=(0, 5))
        
        # 进度条
        progress_bar = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        # 统计标签
        stats_label = ttk.Label(progress_frame, text="成功: 0 | 失败: 0 | 进度: 0%")
        stats_label.pack()
        
        # 按钮框架
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X)
        
        # 停止按钮
        stop_button = ttk.Button(button_frame, text="⏹️ 停止处理", style='Accent.TButton')
        stop_button.pack(side=tk.LEFT)
        
        # 最小化按钮
        minimize_button = ttk.Button(button_frame, text="🔽 最小化", command=lambda: control_dialog.iconify())
        minimize_button.pack(side=tk.LEFT, padx=(10, 0))
        
        # 关闭按钮（初始状态禁用）
        close_button = ttk.Button(button_frame, text="❌ 关闭", state='disabled')
        close_button.pack(side=tk.RIGHT)
        
        # 处理状态变量
        processing_stopped = threading.Event()
        processing_finished = threading.Event()
        
        def stop_processing():
            """停止处理"""
            self.ai_processor.stop_current_processing()
            processing_stopped.set()
            stop_button.config(text="正在停止...", state='disabled')
            progress_label.config(text="正在停止处理，请稍候...")
            
        def close_dialog():
            """关闭对话框"""
            if not processing_finished.is_set():
                result = messagebox.askyesno("确认", "处理尚未完成，确定要关闭吗？")
                if not result:
                    return
                self.ai_processor.stop_current_processing()
            control_dialog.destroy()
        
        def update_progress_callback(completed, total, success_count, failed_count):
            """进度更新回调"""
            if control_dialog.winfo_exists():  # 检查窗口是否仍然存在
                try:
                    percentage = (completed / total * 100) if total > 0 else 0
                    progress_bar['value'] = percentage
                    progress_label.config(text=f"已处理 {completed}/{total} 个任务")
                    stats_label.config(text=f"成功: {success_count} | 失败: {failed_count} | 进度: {percentage:.1f}%")
                    control_dialog.update_idletasks()
                except tk.TclError:
                    # 窗口已关闭
                    pass
        
        def processing_thread():
            """处理线程"""
            try:
                # 设置AI处理器参数
                self.ai_processor.set_processing_params(
                    max_workers=processing_params['max_workers'],
                    request_delay=processing_params['request_delay'],
                    max_retries=processing_params['max_retries']
                )
                
                # 开始处理
                success, result = self.ai_processor.process_column_concurrent(
                    df, col_name, prompt_template, model, self.table_manager,
                    update_progress_callback, rows_to_process, 
                    output_fields if output_mode == "multi" else None
                )
                
                processing_finished.set()
                
                # 更新UI（在主线程中）
                def update_ui():
                    if control_dialog.winfo_exists():
                        try:
                            if success:
                                stats = result
                                final_text = f"处理完成！成功: {stats['success_count']}, 失败: {stats['failed_count']}"
                                if stats['stopped']:
                                    final_text = f"处理已停止。成功: {stats['success_count']}, 失败: {stats['failed_count']}"
                                progress_label.config(text=final_text)
                                stop_button.config(text="已完成", state='disabled')
                            else:
                                progress_label.config(text=f"处理失败: {result}")
                                stop_button.config(text="已失败", state='disabled')
                                
                            close_button.config(state='normal')
                            
                            # 更新表格显示
                            self.update_table_display()
                            
                        except tk.TclError:
                            pass  # 窗口已关闭
                            
                control_dialog.after(0, update_ui)
                
            except Exception as e:
                processing_finished.set()
                def show_error():
                    if control_dialog.winfo_exists():
                        try:
                            progress_label.config(text=f"处理出错: {str(e)}")
                            stop_button.config(text="已出错", state='disabled')
                            close_button.config(state='normal')
                        except tk.TclError:
                            pass
                control_dialog.after(0, show_error)
        
        # 绑定按钮事件
        stop_button.config(command=stop_processing)
        close_button.config(command=close_dialog)
        
        # 绑定窗口关闭事件
        control_dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        
        # 启动处理线程
        thread = threading.Thread(target=processing_thread, daemon=True)
        thread.start()
        
        self.update_status(f"开始并发处理列 {col_name}，任务数: {task_count}", "normal")

    def update_title(self):
        """更新窗口标题，显示文件状态"""
        title = "AI Excel Tool v2.0"
        
        # 如果有当前项目或文件
        if self.current_project_path:
            filename = os.path.basename(self.current_project_path)
            title += f" - {filename}"
        elif self.table_manager.file_path:
            filename = os.path.basename(self.table_manager.file_path)
            title += f" - {filename}"
        elif self.table_manager.get_dataframe() is not None:
            title += " - 未命名项目"
            
        # 如果有未保存的更改，添加星号
        if self.table_manager.has_unsaved_changes():
            title += " *"
            
        self.root.title(title)

    def check_unsaved_changes_before_action(self, action_name):
        """在执行操作前检查是否有未保存的更改"""
        if self.table_manager.has_unsaved_changes():
            response = messagebox.askyesnocancel(
                "未保存的更改", 
                f"当前有未保存的更改，在{action_name}前是否要保存？\n\n"
                "选择'是'保存后继续，'否'直接继续，'取消'停止操作。"
            )
            if response is True:  # 用户选择"是" (保存)
                success, _ = self.save_project()
                return success  # 只有保存成功才继续
            elif response is False:  # 用户选择"否" (不保存)
                return True  # 继续操作
            else:  # 用户选择"取消"
                return False  # 停止操作
        return True  # 没有未保存更改，继续操作
    
    def save_state(self, action_description):
        """保存当前状态到撤销历史"""
        try:
            if self.table_manager.get_dataframe() is not None:
                state = {
                    'action': action_description,
                    'timestamp': time.time(),
                    'dataframe': self.table_manager.get_dataframe().copy(),
                    'ai_columns': self.table_manager.get_ai_columns().copy(),
                    'long_text_columns': self.table_manager.get_long_text_columns().copy()
                }
                
                self.undo_history.append(state)
                print(f"DEBUG: Saved state for action: {action_description}, history length: {len(self.undo_history)}")
                
                # 限制历史记录数量
                if len(self.undo_history) > self.max_undo_steps:
                    self.undo_history.pop(0)
                    
        except Exception as e:
            print(f"保存状态失败: {e}")
    
    def undo_action(self):
        """撤销上一步操作"""
        try:
            print(f"DEBUG: undo_action called, history length: {len(self.undo_history)}")
            if not self.undo_history:
                self.update_status("没有可撤销的操作", "normal")
                print("DEBUG: No undo history available")
                return
                
            # 获取最后一个状态
            last_state = self.undo_history.pop()
            print(f"DEBUG: Undoing action: {last_state['action']}")
            
            # 恢复数据框
            self.table_manager.dataframe = last_state['dataframe'].copy()
            
            # 恢复AI列配置
            self.table_manager.ai_columns = last_state['ai_columns'].copy()
            
            # 恢复长文本列配置
            self.table_manager.long_text_columns = last_state['long_text_columns'].copy()
            
            # 更新界面显示
            self.update_table_display()
            
            # 设置为有未保存的更改
            self.table_manager.changes_made = True
            self.update_title()
            
            action = last_state['action']
            self.update_status(f"已撤销: {action}", "success")
            print(f"DEBUG: Successfully undone: {action}")
            
        except Exception as e:
            self.update_status(f"撤销失败: {str(e)}", "error")
            print(f"撤销操作失败: {e}")

def main():
    root = tk.Tk()
    
    # 设置简单对话框模块别名
    tk.simpledialog = tkinter.simpledialog
    
    app = AIExcelApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 