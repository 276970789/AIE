#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格数据管理器
负责数据的加载、保存、AI列配置管理
"""

import pandas as pd
import os

class TableManager:
    def __init__(self):
        self.dataframe = None
        self.ai_columns = {}  # {column_name: {"prompt": prompt_template, "model": model_name}}
        self.file_path = None
        
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
            
            return True
            
        except Exception as e:
            print(f"创建空白表格错误: {e}")
            return False
        
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
        
    def add_ai_column(self, column_name, prompt_template, model="gpt-4.1"):
        """添加AI列"""
        if self.dataframe is not None:
            # 添加空列到数据框
            self.dataframe[column_name] = ''
            # 保存AI列配置（包含模型信息）
            self.ai_columns[column_name] = {
                "prompt": prompt_template,
                "model": model
            }
            
    def add_normal_column(self, column_name, default_value=''):
        """添加普通列"""
        if self.dataframe is not None:
            self.dataframe[column_name] = default_value
            
    def add_row(self):
        """添加新行"""
        if self.dataframe is not None:
            try:
                # 创建新行，所有列都设为空字符串
                new_row = {col: '' for col in self.dataframe.columns}
                
                # 使用concat而不是append（pandas 2.0+推荐）
                new_df = pd.DataFrame([new_row])
                self.dataframe = pd.concat([self.dataframe, new_df], ignore_index=True)
                
                return True
            except Exception as e:
                print(f"添加行错误: {e}")
                return False
        return False
        
    def clear_all_data(self):
        """清空所有数据"""
        self.dataframe = None
        self.ai_columns = {}
        self.file_path = None
        
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
                # 向后兼容：如果是旧格式，默认返回gpt-4.1
                return "gpt-4.1"
        return "gpt-4.1"
    
    def get_ai_columns_simple(self):
        """获取简化的AI列配置（向后兼容）"""
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
        """删除列"""
        if self.dataframe is not None and column_name in self.dataframe.columns:
            print(f"删除列: {column_name}")
            
            # 使用drop方法删除列，并直接赋值
            self.dataframe = self.dataframe.drop(columns=[column_name])
            
            # 如果是AI列，也删除配置
            if column_name in self.ai_columns:
                del self.ai_columns[column_name]
                print(f"删除AI列配置: {column_name}")
                
            print(f"删除后列名: {list(self.dataframe.columns)}")
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
            try:
                # 重命名DataFrame中的列
                self.dataframe = self.dataframe.rename(columns={old_name: new_name})
                
                # 如果是AI列，也要更新AI配置
                if old_name in self.ai_columns:
                    prompt = self.ai_columns[old_name]
                    del self.ai_columns[old_name]
                    self.ai_columns[new_name] = prompt
                    print(f"AI列配置已更新: {old_name} → {new_name}")
                
                print(f"列重命名成功: {old_name} → {new_name}")
                return True
            except Exception as e:
                print(f"重命名列失败: {e}")
                return False
        else:
            print(f"列不存在: {old_name}")
            return False
            
    def update_ai_column_prompt(self, column_name, new_prompt):
        """更新AI列的提示词"""
        if column_name in self.ai_columns:
            self.ai_columns[column_name] = new_prompt
            print(f"AI提示词已更新: {column_name}")
            return True
        else:
            print(f"AI列不存在: {column_name}")
            return False
    
    def update_ai_column_config(self, column_name, new_prompt, new_model):
        """更新AI列的完整配置（包含模型）"""
        if column_name in self.ai_columns:
            self.ai_columns[column_name] = {
                "prompt": new_prompt,
                "model": new_model
            }
            print(f"AI列配置已更新: {column_name} (模型: {new_model})")
            return True
        else:
            print(f"AI列不存在: {column_name}")
            return False
            
    def convert_to_ai_column(self, column_name, prompt_template):
        """将普通列转换为AI列"""
        if self.dataframe is not None and column_name in self.dataframe.columns:
            # 添加到AI列配置
            self.ai_columns[column_name] = prompt_template
            print(f"已转换为AI列: {column_name}")
            return True
        else:
            print(f"列不存在: {column_name}")
            return False
            
    def convert_to_normal_column(self, column_name):
        """将AI列转换为普通列"""
        if column_name in self.ai_columns:
            # 从AI列配置中移除
            del self.ai_columns[column_name]
            print(f"已转换为普通列: {column_name}")
            return True
        else:
            print(f"AI列不存在: {column_name}")
            return False
            
    def insert_column_at_position(self, position, column_name, prompt_template=None, is_ai_column=False, ai_model="gpt-4.1"):
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
                
                # 如果是AI列，添加到AI配置
                if is_ai_column and prompt_template:
                    self.ai_columns[column_name] = {
                        "prompt": prompt_template,
                        "model": ai_model
                    }
                    
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
                print(f"已在位置{position}插入新行")
                return True
                
            except Exception as e:
                print(f"插入行失败: {e}")
                return False
        else:
            print("数据框为空，无法插入行")
            return False 