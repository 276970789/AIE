#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目管理器
支持保存和加载包含AI列配置的项目文件
"""

import json
import os
import zipfile
import tempfile
from datetime import datetime
import pandas as pd

class ProjectManager:
    def __init__(self):
        self.project_format_version = "1.0"
        
    def save_project(self, file_path, table_manager, ai_processor=None, column_widths=None):
        """
        保存项目到.aie文件（AI Excel Project）
        包含数据、AI列配置、界面状态等
        """
        try:
            # 准备项目数据
            project_data = {
                "format_version": self.project_format_version,
                "created_at": datetime.now().isoformat(),
                "app_version": "2.2",
                "project_info": {
                    "name": os.path.splitext(os.path.basename(file_path))[0],
                    "description": "AI Excel批量数据处理项目"
                }
            }
            
            # 获取数据框
            df = table_manager.get_dataframe()
            if df is not None:
                # 表格数据
                project_data["table_data"] = {
                    "columns": list(df.columns),
                    "data": df.to_dict('records'),
                    "row_count": len(df),
                    "col_count": len(df.columns)
                }
                
                # AI列配置
                ai_columns = table_manager.get_ai_columns()
                project_data["ai_config"] = {
                    "ai_columns": ai_columns,
                    "ai_column_count": len(ai_columns),
                    "prompt_templates": {}
                }
                
                # 长文本列配置
                long_text_columns = table_manager.get_long_text_columns()
                project_data["long_text_config"] = {
                    "long_text_columns": long_text_columns,
                    "long_text_column_count": len(long_text_columns)
                }
                
                # 保存每个AI列的详细配置
                for col_name, prompt in ai_columns.items():
                    # 如果是旧格式的prompt（字符串），将其转换为字典格式
                    if isinstance(prompt, str):
                        prompt_dict = {"prompt": prompt, "model": "gpt-4.1"}
                    else:
                        prompt_dict = prompt # 已经是字典格式
                    project_data["ai_config"]["prompt_templates"][col_name] = {
                        "prompt": prompt_dict["prompt"],
                        "column_type": "ai",
                        "model": prompt_dict.get("model", "gpt-4.1"), # 确保模型信息存在
                        "created_at": datetime.now().isoformat(),
                        "last_processed": None,
                        "processing_stats": {
                            "total_processed": 0,
                            "success_count": 0,
                            "error_count": 0
                        }
                    }
                
                # 普通列信息
                normal_columns = [col for col in df.columns if col not in ai_columns and col not in long_text_columns]
                project_data["normal_columns"] = normal_columns
                
                # 界面状态（可选）
                project_data["ui_state"] = {
                    "last_selected_column": None,
                    "last_selected_row": None,
                    "table_sorting": None,
                    "row_height_setting": "low"
                }
                # 保存列宽信息
                if column_widths is not None:
                    project_data["ui_state"]["column_widths"] = column_widths
                
            else:
                # 空项目
                project_data["table_data"] = None
                project_data["ai_config"] = {"ai_columns": {}, "ai_column_count": 0}
                project_data["long_text_config"] = {"long_text_columns": {}, "long_text_column_count": 0}
                project_data["normal_columns"] = []
                project_data["ui_state"] = {}
            
            # 保存项目文件
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
                
            return True, f"项目已保存到: {file_path}"
            
        except Exception as e:
            return False, f"保存项目失败: {str(e)}"
    
    def load_project(self, file_path, table_manager):
        """
        从.aie文件加载项目
        恢复数据、AI列配置等
        """
        try:
            # 读取项目文件
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            # 验证文件格式
            if project_data.get("format_version") != self.project_format_version:
                return False, f"不支持的项目文件格式版本: {project_data.get('format_version')}", {}
            
            # 恢复表格数据
            table_data = project_data.get("table_data")
            if table_data and table_data.get("data"):
                # 创建数据框
                df = pd.DataFrame(table_data["data"])
                if not df.empty:
                    # 确保列顺序正确
                    expected_columns = table_data.get("columns", [])
                    if expected_columns:
                        df = df.reindex(columns=expected_columns)
                    
                    # 设置到table_manager
                    table_manager.dataframe = df
                    
                    # 清空table_manager中现有的AI列和长文本列配置，避免重复
                    table_manager.ai_columns = {}
                    table_manager.long_text_columns = {}

                    # 恢复AI列配置
                    ai_config = project_data.get("ai_config", {})
                    # 优先从prompt_templates恢复完整的AI列配置
                    ai_columns = ai_config.get("prompt_templates", ai_config.get("ai_columns", {}))
                    for col_name, config in ai_columns.items():
                        if col_name in df.columns: # 确保列存在于数据框中
                            table_manager.add_ai_column(col_name, config["prompt"], config.get("model", "gpt-4.1"), preserve_data=True)
                    
                    # 恢复长文本列配置
                    long_text_config = project_data.get("long_text_config", {})
                    long_text_columns = long_text_config.get("long_text_columns", {})
                    for col_name, config in long_text_columns.items():
                        if col_name in df.columns: # 确保列存在于数据框中
                            table_manager.add_long_text_column(col_name, config["filename_field"], config["folder_path"], config.get("preview_length", 200))

                    # 刷新长文本列内容（如果存在）
                    for col_name in table_manager.long_text_columns:
                        table_manager.refresh_long_text_column(col_name)
                    
                    # 恢复界面状态（可选）
                    ui_state = project_data.get("ui_state", {})
                    column_widths = ui_state.get("column_widths", {})
                    
                    return True, f"项目加载成功: {len(df)}行 {len(df.columns)}列 (AI列: {len(table_manager.ai_columns)}, 长文本列: {len(table_manager.long_text_columns)})", column_widths
            else:
                # 空项目
                table_manager.create_blank_table()
                return True, "加载了空白项目", {}
                
        except FileNotFoundError:
            return False, f"项目文件不存在: {file_path}", {}
        except json.JSONDecodeError:
            return False, "项目文件格式错误，不是有效的JSON文件", {}
        except Exception as e:
            return False, f"加载项目失败: {str(e)}", {}
    
    def get_project_info(self, file_path):
        """
        获取项目文件的基本信息，不加载数据
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            info = {
                "name": project_data.get("project_info", {}).get("name", "未知项目"),
                "created_at": project_data.get("created_at", "未知时间"),
                "app_version": project_data.get("app_version", "未知版本"),
                "format_version": project_data.get("format_version", "未知格式"),
                "row_count": 0,
                "col_count": 0,
                "ai_column_count": 0
            }
            
            table_data = project_data.get("table_data")
            if table_data:
                info["row_count"] = table_data.get("row_count", 0)
                info["col_count"] = table_data.get("col_count", 0)
            
            ai_config = project_data.get("ai_config", {})
            info["ai_column_count"] = ai_config.get("ai_column_count", 0)
            
            return True, info
            
        except Exception as e:
            return False, f"读取项目信息失败: {str(e)}"
    
    def export_project_summary(self, file_path):
        """
        导出项目摘要（markdown格式）
        """
        try:
            success, info = self.get_project_info(file_path)
            if not success:
                return False, info
            
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            # 生成摘要内容
            summary = f"""# 项目摘要

## 基本信息
- **项目名称**: {info['name']}
- **创建时间**: {info['created_at']}
- **应用版本**: {info['app_version']}
- **数据规模**: {info['row_count']}行 {info['col_count']}列
- **AI列数量**: {info['ai_column_count']}个

## AI列配置

"""
            
            ai_config = project_data.get("ai_config", {})
            prompt_templates = ai_config.get("prompt_templates", {})
            
            if prompt_templates:
                for col_name, config in prompt_templates.items():
                    summary += f"### {col_name}\n"
                    summary += f"**提示词模板**:\n```\n{config['prompt']}\n```\n\n"
            else:
                summary += "无AI列配置\n\n"
            
            # 普通列
            normal_columns = project_data.get("normal_columns", [])
            if normal_columns:
                summary += f"## 普通列 ({len(normal_columns)}个)\n"
                for col in normal_columns:
                    summary += f"- {col}\n"
                summary += "\n"
            
            return True, summary
            
        except Exception as e:
            return False, f"生成项目摘要失败: {str(e)}"
    
    def validate_project_file(self, file_path):
        """
        验证项目文件的有效性
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            # 检查必要字段
            required_fields = ["format_version", "created_at"]
            for field in required_fields:
                if field not in project_data:
                    return False, f"缺少必要字段: {field}"
            
            # 检查格式版本
            if project_data["format_version"] != self.project_format_version:
                return False, f"格式版本不匹配: {project_data['format_version']} != {self.project_format_version}"
            
            return True, "项目文件有效"
            
        except json.JSONDecodeError:
            return False, "JSON格式错误"
        except Exception as e:
            return False, f"验证失败: {str(e)}" 