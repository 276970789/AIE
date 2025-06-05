#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格数据管理器
负责数据的加载、保存、AI列配置管理
"""

import pandas as pd
import os
from paper_processor import get_paper_processor

class TableManager:
    def __init__(self):
        self.dataframe = None
        self.ai_columns = {}  # {column_name: {"prompt": prompt_template, "model": model_name}}
        self.long_text_columns = {}  # {column_name: {"filename_field": field, "folder_path": path, "preview_length": length}}
        self.hidden_columns = set()  # 隐藏的列名集合
        self.file_path = None
        self.changes_made = False # 添加标志，追踪是否有未保存的更改
        
    def create_blank_table(self):
        """创建空白表格"""
        try:
            # 创建带示例数据的表格
            self.dataframe = pd.DataFrame({
                '文本内容': [
                    'Hello, how are you today?',
                    'Good morning, have a nice day!',
                    'Thank you for your help.',
                    ''  # 空行供用户添加
                ],
                '类别': [
                    '问候',
                    '祝福', 
                    '感谢',
                    ''
                ],
                '备注': [
                    '',
                    '',
                    '',
                    ''
                ]
            })
            
            self.file_path = None
            self.ai_columns = {}
            
            self.changes_made = True # 创建空白表格即视为有更改
            return True
            
        except Exception as e:
            print(f"创建空白表格错误: {e}")
            return False
        finally:
            self.changes_made = True # 创建空白表格即视为有更改
        
    def load_file(self, file_path):
        """加载文件"""
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext in ['.xlsx', '.xls']:
                self.dataframe = pd.read_excel(file_path)
            elif file_ext == '.csv':
                # 尝试不同编码
                encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
                for encoding in encodings:
                    try:
                        self.dataframe = pd.read_csv(file_path, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise Exception("无法识别文件编码")
            elif file_ext == '.jsonl':
                # 加载JSONL文件
                self.dataframe = self.load_jsonl_file(file_path)
            else:
                raise Exception("不支持的文件格式")
                
            self.file_path = file_path
            # 清空之前的AI列配置
            self.ai_columns = {}
            
            # 填充NaN值
            self.dataframe = self.dataframe.fillna('')
            
            self.changes_made = False # 成功加载文件，无未保存更改
            return True
            
        except Exception as e:
            print(f"加载文件错误: {e}")
            return False
            
    def load_jsonl_file(self, file_path):
        """加载JSONL文件"""
        import json
        
        data_list = []
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    for line_num, line in enumerate(f, 1):
                        line = line.strip()
                        if line:  # 跳过空行
                            try:
                                json_obj = json.loads(line)
                                data_list.append(json_obj)
                            except json.JSONDecodeError as e:
                                print(f"第{line_num}行JSON解析错误: {e}")
                                # 继续处理其他行，不中断整个加载过程
                                continue
                break  # 成功读取，跳出编码循环
            except UnicodeDecodeError:
                continue  # 尝试下一个编码
        else:
            raise Exception("无法识别JSONL文件编码")
            
        if not data_list:
            raise Exception("JSONL文件为空或没有有效的JSON行")
            
        # 转换为DataFrame
        df = pd.DataFrame(data_list)
        
        print(f"成功加载JSONL文件: {len(data_list)}行数据, {len(df.columns)}列")
        return df
        
    def get_dataframe(self):
        """获取数据框"""
        return self.dataframe
        
    def get_column_names(self):
        """获取列名列表"""
        if self.dataframe is not None:
            return list(self.dataframe.columns)
        return []
        
    def add_ai_column(self, column_name, prompt_template, model="gpt-4.1", processing_params=None, output_mode="single", output_fields=None, field_mode="predefined", preserve_data=False):
        """添加AI列"""
        if self.dataframe is not None:
            # 添加主列到数据框（如果不存在）
            if column_name not in self.dataframe.columns:
                self.dataframe[column_name] = ''
            elif not preserve_data:
                # 如果列已存在且不保留数据，则清空
                self.dataframe[column_name] = ''
            
            # 如果是多字段模式，创建对应的字段列（紧挨着主列）
            if output_mode == "multi" and output_fields:
                # 获取主列的位置
                columns = list(self.dataframe.columns)
                main_col_index = columns.index(column_name)
                
                # 准备要插入的新列
                new_columns_to_add = []
                for field_name in output_fields:
                    full_column_name = f"{column_name}_{field_name}"
                    if full_column_name not in self.dataframe.columns:
                        new_columns_to_add.append(full_column_name)
                    elif not preserve_data:
                        self.dataframe[full_column_name] = ''
                
                # 如果有新列需要添加，按顺序插入到主列右边
                if new_columns_to_add:
                    # 为新列添加空数据
                    for col_name in new_columns_to_add:
                        self.dataframe[col_name] = ''
                    
                    # 重新排列列的顺序，将新列放在主列右边
                    new_column_order = []
                    for i, col in enumerate(columns):
                        new_column_order.append(col)
                        if col == column_name:
                            # 在主列后面插入所有字段列
                            new_column_order.extend(new_columns_to_add)
                    
                    # 添加其他可能存在的列（如果有的话）
                    for col in self.dataframe.columns:
                        if col not in new_column_order:
                            new_column_order.append(col)
                    
                    # 重新排列DataFrame的列
                    self.dataframe = self.dataframe[new_column_order]
            
            # 默认处理参数
            default_params = {
                'max_workers': 3,
                'request_delay': 0.5,
                'max_retries': 2
            }
            
            if processing_params:
                default_params.update(processing_params)
            
            # 保存AI列配置（包含模型信息、处理参数和多字段配置）
            self.ai_columns[column_name] = {
                "prompt": prompt_template,
                "model": model,
                "processing_params": default_params,
                "output_mode": output_mode,
                "output_fields": output_fields or [],
                "field_mode": field_mode
            }
            self.changes_made = True # 添加AI列即视为有更改
            return True
        return False
            
    def add_normal_column(self, column_name, default_value=''):
        """添加普通列"""
        if self.dataframe is not None:
            self.dataframe[column_name] = default_value
            self.changes_made = True # 添加普通列即视为有更改
            
    def add_row(self):
        """添加新行"""
        if self.dataframe is not None:
            try:
                # 创建新行，所有列都设为空字符串
                new_row = {col: '' for col in self.dataframe.columns}
                
                # 使用concat而不是append（pandas 2.0+推荐）
                new_df = pd.DataFrame([new_row])
                self.dataframe = pd.concat([self.dataframe, new_df], ignore_index=True)
                
                self.changes_made = True # 添加行即视为有更改
                return True
            except Exception as e:
                print(f"添加行错误: {e}")
                return False
        return False
        
    def clear_all_data(self):
        """清空所有数据"""
        self.dataframe = None
        self.ai_columns = {}
        self.long_text_columns = {}
        self.file_path = None
        self.changes_made = True # 清空数据即视为有更改
        
    def get_ai_columns(self):
        """获取AI列配置"""
        return self.ai_columns
    
    def get_ai_column_prompt(self, column_name):
        """获取AI列的prompt模板"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            if isinstance(config, dict):
                return config.get("prompt", "")
            else:
                # 向后兼容：如果是旧格式（字符串），直接返回
                return config
        return ""
    
    def get_ai_column_model(self, column_name):
        """获取AI列使用的模型"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            if isinstance(config, dict):
                return config.get("model", "gpt-4.1")
            else:
                # 向后兼容：旧格式默认使用gpt-4.1
                return "gpt-4.1"
        return "gpt-4.1"
    
    def get_ai_column_processing_params(self, column_name):
        """获取AI列的处理参数"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            if isinstance(config, dict) and "processing_params" in config:
                return config["processing_params"]
            else:
                # 返回默认参数
                return {
                    'max_workers': 3,
                    'request_delay': 0.5,
                    'max_retries': 2
                }
        return {
            'max_workers': 3,
            'request_delay': 0.5,
            'max_retries': 2
        }
        
    def get_ai_column_output_mode(self, column_name):
        """获取AI列的输出模式"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            if isinstance(config, dict):
                return config.get("output_mode", "single")
        return "single"
        
    def get_ai_column_output_fields(self, column_name):
        """获取AI列的输出字段"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            if isinstance(config, dict):
                return config.get("output_fields", [])
        return []
        
    def get_ai_column_field_mode(self, column_name):
        """获取AI列的字段处理模式"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            if isinstance(config, dict):
                return config.get("field_mode", "predefined")
        return "predefined"
        
    def is_multi_field_ai_column(self, column_name):
        """检查是否为多字段AI列"""
        return (column_name in self.ai_columns and 
                self.get_ai_column_output_mode(column_name) == "multi" and
                len(self.get_ai_column_output_fields(column_name)) > 0)
        
    def get_ai_columns_simple(self):
        """获取简化的AI列配置（仅包含prompt，向后兼容）"""
        simple_config = {}
        for col_name, config in self.ai_columns.items():
            if isinstance(config, dict):
                simple_config[col_name] = config.get("prompt", "")
            else:
                simple_config[col_name] = config
        return simple_config
        
    def update_ai_column_value(self, column_name, row_index, value):
        """更新AI列的值"""
        if self.dataframe is not None and column_name in self.dataframe.columns:
            self.dataframe.at[row_index, column_name] = value
            self.changes_made = True # 更新单元格值即视为有更改
            
    def ensure_multi_field_columns_positioned(self, main_column_name, field_names):
        """确保多字段列在主列右边的正确位置"""
        if self.dataframe is None or main_column_name not in self.dataframe.columns:
            return
            
        columns = list(self.dataframe.columns)
        main_col_index = columns.index(main_column_name)
        
        # 检查哪些字段列需要创建或重新定位
        field_columns = [f"{main_column_name}_{field}" for field in field_names]
        columns_to_reposition = []
        
        for field_col in field_columns:
            if field_col not in self.dataframe.columns:
                # 创建新列
                self.dataframe[field_col] = ''
                columns_to_reposition.append(field_col)
            else:
                # 检查现有列是否在正确位置
                current_index = columns.index(field_col)
                expected_start = main_col_index + 1
                expected_end = main_col_index + len(field_columns)
                if not (expected_start <= current_index <= expected_end):
                    columns_to_reposition.append(field_col)
        
        # 如果需要重新定位列
        if columns_to_reposition:
            # 重新排列列的顺序
            new_column_order = []
            
            # 添加主列之前的所有列
            for col in columns:
                if col == main_column_name:
                    new_column_order.append(col)
                    # 在主列后面添加所有字段列
                    new_column_order.extend(field_columns)
                elif col not in field_columns:
                    new_column_order.append(col)
            
            # 添加任何新创建的列
            for col in self.dataframe.columns:
                if col not in new_column_order:
                    new_column_order.append(col)
            
            # 重新排列DataFrame
            self.dataframe = self.dataframe[new_column_order]
            self.changes_made = True
            
    def get_row_data(self, row_index):
        """获取指定行的数据"""
        if self.dataframe is not None:
            return self.dataframe.iloc[row_index].to_dict()
        return {}
        
    def get_row_count(self):
        """获取行数"""
        if self.dataframe is not None:
            return len(self.dataframe)
        return 0
        
    def export_excel(self, file_path):
        """导出Excel文件"""
        if self.dataframe is not None:
            self.dataframe.to_excel(file_path, index=False)
            
    def export_csv(self, file_path):
        """导出CSV文件"""
        if self.dataframe is not None:
            self.dataframe.to_csv(file_path, index=False, encoding='utf-8-sig')
            
    def export_jsonl(self, file_path):
        """导出JSONL文件"""
        if self.dataframe is not None:
            import json
            
            with open(file_path, 'w', encoding='utf-8') as f:
                for _, row in self.dataframe.iterrows():
                    # 将每行转换为字典，然后转换为JSON字符串
                    row_dict = row.to_dict()
                    json_line = json.dumps(row_dict, ensure_ascii=False)
                    f.write(json_line + '\n')
                    
            print(f"成功导出JSONL文件: {len(self.dataframe)}行数据")
            
    def delete_column(self, column_name):
        """删除指定列"""
        if self.dataframe is not None and column_name in self.dataframe.columns:
            print(f"删除列: {column_name}")
            
            # 使用drop方法删除列，并直接赋值
            self.dataframe = self.dataframe.drop(columns=[column_name])
            
            # 从AI列配置中移除
            if column_name in self.ai_columns:
                del self.ai_columns[column_name]
                print(f"删除AI列配置: {column_name}")
                
            # 从长文本列配置中移除
            if column_name in self.long_text_columns:
                del self.long_text_columns[column_name]
                print(f"删除长文本列配置: {column_name}")
                
            print(f"删除后列名: {list(self.dataframe.columns)}")
            self.changes_made = True # 删除列即视为有更改
            return True
        else:
            print(f"列不存在或数据框为空: {column_name}")
            return False
            
    def validate_prompt_template(self, prompt_template):
        """验证prompt模板中的字段引用是否有效"""
        if self.dataframe is None:
            return False, "没有加载数据"
            
        import re
        # 提取模板中的字段引用
        field_refs = re.findall(r'\{(\w+)\}', prompt_template)
        
        invalid_fields = []
        for field in field_refs:
            if field not in self.dataframe.columns:
                invalid_fields.append(field)
                
        if invalid_fields:
            return False, f"字段不存在: {', '.join(invalid_fields)}"
            
        return True, "模板有效"
        
    def rename_column(self, old_name, new_name):
        """重命名列"""
        if self.dataframe is not None and old_name in self.dataframe.columns:
            if new_name in self.dataframe.columns:
                return False, "新列名已存在"
            
            self.dataframe = self.dataframe.rename(columns={old_name: new_name})
            
            # 更新AI列配置
            if old_name in self.ai_columns:
                config = self.ai_columns.pop(old_name)
                self.ai_columns[new_name] = config
            
            # 更新长文本列配置
            if old_name in self.long_text_columns:
                config = self.long_text_columns.pop(old_name)
                self.long_text_columns[new_name] = config

            self.changes_made = True # 重命名列即视为有更改
            return True, ""
        return False, "列不存在"
            
    def update_ai_column_prompt(self, column_name, new_prompt):
        """更新AI列的提示词"""
        if column_name in self.ai_columns:
            if isinstance(self.ai_columns[column_name], dict):
                self.ai_columns[column_name]["prompt"] = new_prompt
            else:
                self.ai_columns[column_name] = new_prompt # 向后兼容旧格式
            self.changes_made = True # 更新AI列prompt即视为有更改
            print(f"AI提示词已更新: {column_name}")
            return True
        else:
            print(f"AI列不存在: {column_name}")
            return False
    
    def update_ai_column_config(self, column_name, new_prompt, new_model, processing_params=None, output_mode=None, output_fields=None, field_mode=None):
        """更新AI列配置（包含模型信息、处理参数、输出模式和字段）"""
        if column_name in self.ai_columns:
            config = self.ai_columns[column_name]
            
            # 保留现有的处理参数或使用新的
            if processing_params is None:
                if isinstance(config, dict) and "processing_params" in config:
                    processing_params = config["processing_params"]
                else:
                    processing_params = {
                        'max_workers': 3,
                        'request_delay': 0.5,
                        'max_retries': 2
                    }
            
            # 保留现有的输出模式和字段或使用新的
            if output_mode is None:
                if isinstance(config, dict):
                    output_mode = config.get("output_mode", "single")
                else:
                    output_mode = "single"
            
            if output_fields is None:
                if isinstance(config, dict):
                    output_fields = config.get("output_fields", [])
                else:
                    output_fields = []
            
            if field_mode is None:
                if isinstance(config, dict):
                    field_mode = config.get("field_mode", "predefined")
                else:
                    field_mode = "predefined"
            
            self.ai_columns[column_name] = {
                "prompt": new_prompt,
                "model": new_model,
                "processing_params": processing_params,
                "output_mode": output_mode,
                "output_fields": output_fields or [],
                "field_mode": field_mode
            }
            self.changes_made = True
            
    def convert_to_ai_column(self, column_name, prompt_template):
        """将普通列转换为AI列"""
        if self.dataframe is not None and column_name in self.dataframe.columns:
            if column_name not in self.ai_columns:
                self.ai_columns[column_name] = {"prompt": prompt_template, "model": "gpt-4.1"}
                # 如果是长文本列，需要从长文本列配置中移除
                if column_name in self.long_text_columns:
                    del self.long_text_columns[column_name]
                self.changes_made = True # 转换为AI列即视为有更改
                print(f"已转换为AI列: {column_name}")
                return True
        return False
            
    def convert_to_normal_column(self, column_name):
        """将AI列转换为普通列"""
        if column_name in self.dataframe.columns:
            if column_name in self.ai_columns:
                del self.ai_columns[column_name]
                self.changes_made = True # 转换为普通列即视为有更改
                print(f"已转换为普通列: {column_name}")
                return True
            if column_name in self.long_text_columns:
                del self.long_text_columns[column_name]
                self.changes_made = True # 转换为普通列即视为有更改
                print(f"已转换为普通列: {column_name}")
                return True
        return False
            
    def insert_column_at_position(self, position, column_name, prompt_template=None, is_ai_column=False, ai_model="gpt-4.1", processing_params=None, output_mode="single", output_fields=None, field_mode="predefined"):
        """在指定位置插入列"""
        if self.dataframe is not None:
            try:
                # 获取当前列名列表
                columns = list(self.dataframe.columns)
                
                # 确保位置在有效范围内
                position = max(0, min(position, len(columns)))
                
                # 创建新的列顺序
                new_columns = columns[:position] + [column_name] + columns[position:]
                
                # 为新列添加空值
                self.dataframe[column_name] = ''
                
                # 重新排列列的顺序
                self.dataframe = self.dataframe[new_columns]
                
                # 更新或添加AI列配置
                if is_ai_column:
                    # 默认处理参数
                    default_params = {
                        'max_workers': 3,
                        'request_delay': 0.5,
                        'max_retries': 2
                    }
                    
                    if processing_params:
                        default_params.update(processing_params)
                    
                    self.ai_columns[column_name] = {
                        "prompt": prompt_template, 
                        "model": ai_model,
                        "processing_params": default_params,
                        "output_mode": output_mode or "single",
                        "output_fields": output_fields or [],
                        "field_mode": field_mode
                    }
                    
                self.changes_made = True # 插入列即视为有更改
                print(f"已在位置{position}插入列: {column_name}")
                return True
                
            except Exception as e:
                print(f"插入列失败: {e}")
                return False
        else:
            print("数据框为空，无法插入列")
            return False
            
    def move_column(self, from_index, to_index):
        """移动列位置"""
        if self.dataframe is not None:
            try:
                columns = list(self.dataframe.columns)
                
                # 检查索引有效性
                if not (0 <= from_index < len(columns)):
                    print(f"无效的源索引: from={from_index}")
                    return False
                    
                if from_index == to_index:
                    return True  # 没有移动
                    
                # 调整目标索引，确保在有效范围内
                to_index = max(0, min(to_index, len(columns) - 1))
                    
                # 获取要移动的列名
                column_to_move = columns[from_index]
                
                # 创建新的列顺序
                new_columns = columns.copy()
                new_columns.pop(from_index)  # 移除原位置的列
                new_columns.insert(to_index, column_to_move)  # 在新位置插入
                
                # 重新排列DataFrame的列
                self.dataframe = self.dataframe[new_columns]
                
                self.changes_made = True # 移动列即视为有更改
                print(f"已移动列 '{column_to_move}' 从位置{from_index}到位置{to_index}")
                return True
                
            except Exception as e:
                print(f"移动列失败: {e}")
                return False
        else:
            print("数据框为空，无法移动列")
            return False
            
    def delete_row(self, row_index):
        """删除指定行"""
        if self.dataframe is not None:
            try:
                # 检查行索引有效性
                if not (0 <= row_index < len(self.dataframe)):
                    print(f"无效的行索引: {row_index}")
                    return False
                    
                # 删除指定行
                self.dataframe = self.dataframe.drop(self.dataframe.index[row_index]).reset_index(drop=True)
                
                self.changes_made = True # 删除行即视为有更改
                print(f"已删除第{row_index + 1}行")
                return True
                
            except Exception as e:
                print(f"删除行失败: {e}")
                return False
        else:
            print("数据框为空，无法删除行")
            return False
    
    def insert_row_at_position(self, position):
        """在指定位置插入空行"""
        if self.dataframe is not None:
            try:
                import pandas as pd
                
                # 确保位置在有效范围内
                position = max(0, min(position, len(self.dataframe)))
                
                # 创建新的空行，所有列都设为空字符串
                new_row = pd.Series([''] * len(self.dataframe.columns), index=self.dataframe.columns)
                
                # 分割数据框并插入新行
                if position == 0:
                    # 在开头插入
                    new_df = pd.concat([pd.DataFrame([new_row]), self.dataframe], ignore_index=True)
                elif position >= len(self.dataframe):
                    # 在末尾插入
                    new_df = pd.concat([self.dataframe, pd.DataFrame([new_row])], ignore_index=True)
                else:
                    # 在中间插入
                    before = self.dataframe.iloc[:position]
                    after = self.dataframe.iloc[position:]
                    new_df = pd.concat([before, pd.DataFrame([new_row]), after], ignore_index=True)
                
                self.dataframe = new_df
                self.changes_made = True # 插入行即视为有更改
                print(f"已在位置{position}插入新行")
                return True
                
            except Exception as e:
                print(f"插入行失败: {e}")
                return False
        else:
            print("数据框为空，无法插入行")
            return False 
    
    # ==================== 长文本列相关方法 ====================
    
    def add_long_text_column(self, column_name, filename_field, folder_path, preview_length=200):
        """添加长文本列"""
        if self.dataframe is not None:
            # 添加空列到数据框
            self.dataframe[column_name] = ''
            
            # 保存长文本列配置
            self.long_text_columns[column_name] = {
                "filename_field": filename_field,
                "folder_path": folder_path,
                "preview_length": preview_length
            }
            
            # 立即加载所有行的长文本内容
            self.refresh_long_text_column(column_name)
            
            self.changes_made = True # 添加长文本列即视为有更改
            print(f"已添加长文本列: {column_name}")
            return True
        return False
    
    def refresh_long_text_column(self, column_name):
        """刷新长文本列的内容"""
        print(f"DEBUG: refresh_long_text_column called for '{column_name}'")
        
        if (column_name not in self.long_text_columns or 
            self.dataframe is None or 
            column_name not in self.dataframe.columns):
            print(f"DEBUG: Early return - column_name in long_text_columns: {column_name in self.long_text_columns}")
            print(f"DEBUG: dataframe is not None: {self.dataframe is not None}")
            if self.dataframe is not None:
                print(f"DEBUG: column_name in dataframe.columns: {column_name in self.dataframe.columns}")
            return False
            
        config = self.long_text_columns[column_name]
        filename_field = config["filename_field"]
        folder_path = config["folder_path"]
        preview_length = config.get("preview_length", 200)
        
        print(f"DEBUG: Config - filename_field: {filename_field}, folder_path: {folder_path}, preview_length: {preview_length}")
        
        # 检查文件名字段是否存在
        if filename_field not in self.dataframe.columns:
            print(f"文件名字段不存在: {filename_field}")
            return False
        
        try:
            # 使用优化的处理器
            processor = get_paper_processor()
            
            # 获取所有文件名
            import pandas as pd
            filename_values = [str(row[filename_field]).strip() if pd.notna(row[filename_field]) else "" 
                              for _, row in self.dataframe.iterrows()]
            
            print(f"DEBUG: Found filename patterns: {filename_values}")
            
            # 批量加载预览
            previews = processor.batch_load_previews(filename_values, folder_path, preview_length)
            
            print(f"DEBUG: Generated {len(previews)} previews")
            
            # 更新列内容
            for index, row in self.dataframe.iterrows():
                filename_pattern = str(row[filename_field]).strip()
                if filename_pattern:
                    preview_content = previews.get(filename_pattern, f"未找到文件: {filename_pattern}")
                    self.dataframe.at[index, column_name] = preview_content
                    print(f"DEBUG: Row {index}, pattern '{filename_pattern}' -> preview: {preview_content[:50]}...")
                else:
                    self.dataframe.at[index, column_name] = "[文件名为空]"
                    print(f"DEBUG: Row {index}, empty filename -> '[文件名为空]'")
            
            print(f"已刷新长文本列: {column_name}")
            return True
            
        except Exception as e:
            print(f"刷新长文本列失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_long_text_columns(self):
        """获取长文本列配置"""
        return self.long_text_columns
    
    def get_long_text_content(self, column_name, row_index):
        """获取指定行的完整长文本内容"""
        if (column_name not in self.long_text_columns or 
            self.dataframe is None or 
            row_index >= len(self.dataframe)):
            return None
            
        config = self.long_text_columns[column_name]
        filename_field = config["filename_field"]
        folder_path = config["folder_path"]
        
        # 获取文件名
        filename_pattern = str(self.dataframe.iloc[row_index][filename_field]).strip()
        if not filename_pattern:
            return None
            
        try:
            # 使用优化的处理器
            processor = get_paper_processor()
            files_map = processor.find_files_in_folder(folder_path)
            
            file_path = files_map.get(filename_pattern)
            if file_path:
                return processor.get_file_content(file_path)
            else:
                return None
        except Exception as e:
            print(f"获取长文本内容失败: {e}")
            return None
    
    def is_long_text_column(self, column_name):
        """检查指定列是否为长文本列"""
        return column_name in self.long_text_columns
    
    def delete_long_text_column(self, column_name):
        """删除长文本列配置"""
        if column_name in self.long_text_columns:
            del self.long_text_columns[column_name]
            self.changes_made = True # 删除长文本列配置即视为有更改
            print(f"已删除长文本列配置: {column_name}")
            return True
        return False
    
    def update_long_text_column_config(self, column_name, filename_field, folder_path, preview_length=200):
        """更新长文本列配置"""
        if column_name in self.long_text_columns:
            self.long_text_columns[column_name] = {
                "filename_field": filename_field,
                "folder_path": folder_path,
                "preview_length": preview_length
            }
            self.changes_made = True # 更新长文本列配置即视为有更改
            # 刷新该列内容
            self.refresh_long_text_column(column_name)
            print(f"已更新长文本列配置: {column_name}")
            return True
        return False

    def reset_changes_flag(self):
        """重置changes_made标志为False"""
        self.changes_made = False
        
    def has_unsaved_changes(self):
        """检查是否有未保存的更改"""
        return self.changes_made
    
    def hide_column(self, column_name):
        """隐藏列"""
        if self.dataframe is not None and column_name in self.dataframe.columns:
            self.hidden_columns.add(column_name)
            self.changes_made = True
            return True
        return False
    
    def show_column(self, column_name):
        """显示列"""
        if column_name in self.hidden_columns:
            self.hidden_columns.remove(column_name)
            self.changes_made = True
            return True
        return False
    
    def is_column_hidden(self, column_name):
        """检查列是否被隐藏"""
        return column_name in self.hidden_columns
    
    def get_hidden_columns(self):
        """获取隐藏的列名列表"""
        return list(self.hidden_columns)
    
    def get_visible_columns(self):
        """获取可见的列名列表"""
        if self.dataframe is not None:
            return [col for col in self.dataframe.columns if col not in self.hidden_columns]
        return []
    
    def get_hidden_columns_count(self):
        """获取隐藏列的数量"""
        return len(self.hidden_columns)
    
    def show_all_columns(self):
        """显示所有列"""
        if self.hidden_columns:
            self.hidden_columns.clear()
            self.changes_made = True
            return True
        return False
        
    def import_from_jsonl(self, match_field, source_field, column_name, jsonl_df, position=None):
        """从JSONL数据导入到新列. 若提供position, 则在该位置插入,否则追加到末尾."""
        if self.dataframe is None:
            return False, "当前没有数据表格"
            
        try:
            # 检查匹配字段是否存在
            if match_field not in self.dataframe.columns:
                return False, f"当前表格中不存在匹配字段 '{match_field}'"
                
            if source_field not in jsonl_df.columns:
                return False, f"JSONL数据中不存在源字段 '{source_field}'"
                
            # 检查列名是否已存在
            if column_name in self.dataframe.columns:
                # This check might be redundant if JsonlImportDialog already ensures unique name,
                # but good for safety.
                return False, f"列名 '{column_name}' 已存在"
                
            #准备一个空的Series来填充数据，长度与当前dataframe一致
            import pandas as pd
            # Initialize with a default value (e.g., empty string) that matches expected type
            new_column_data = pd.Series([''] * len(self.dataframe), index=self.dataframe.index, dtype=object)

            # 执行匹配和导入
            matched_count = 0
            total_rows_in_df = len(self.dataframe)
            
            # Iterating over the current dataframe to fill the new_column_data Series
            for idx, row_data in self.dataframe.iterrows(): # Use self.dataframe.iterrows()
                match_value = row_data[match_field]
                
                # 在JSONL数据中查找匹配项
                # Ensure match_value type is compatible with jsonl_df[match_field] type
                # If types can mismatch (e.g. int vs str), explicit conversion might be needed.
                # For now, assume types are compatible or pandas handles it.
                jsonl_match = jsonl_df[jsonl_df[match_field] == match_value]
                
                if not jsonl_match.empty:
                    import_value = jsonl_match.iloc[0][source_field]
                    # Use .loc for assigning to Series to ensure alignment by index
                    new_column_data.loc[idx] = str(import_value) if pd.notna(import_value) else ''
                    matched_count += 1
            
            # 将填充好的Series插入到DataFrame
            if position is not None:
                # Ensure position is within valid bounds
                if not (0 <= position <= len(self.dataframe.columns)):
                    return False, f"提供的插入位置 {position} 无效."
                self.dataframe.insert(loc=position, column=column_name, value=new_column_data)
            else:
                # Append as before
                self.dataframe[column_name] = new_column_data
                    
            self.changes_made = True
            
            return True, f"成功导入 {matched_count}/{total_rows_in_df} 行数据到列 '{column_name}'"
            
        except Exception as e:
            import traceback
            print(f"Error during JSONL import: {e}")
            traceback.print_exc()
            return False, f"导入失败: {e}"