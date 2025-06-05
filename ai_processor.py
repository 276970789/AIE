#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI处理器
负责批量调用AI API处理数据
"""

import openai
import os
from dotenv import load_dotenv
import time
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue

class AIProcessor:
    def __init__(self):
        # 加载环境变量
        load_dotenv()
        
        # 获取配置，允许未注释的配置生效
        api_key = os.getenv('OPENAI_API_KEY')
        base_url = os.getenv('OPENAI_BASE_URL')
        model = os.getenv('OPENAI_MODEL', 'gpt-4.1')  # 默认使用gpt-4.1
        
        # 打印配置信息用于调试
        print(f"API Key: {api_key[:20] + '...' if api_key else 'None'}")
        print(f"Base URL: {base_url}")
        print(f"Model: {model}")
        
        # 配置OpenAI客户端
        self.client = openai.OpenAI(
            api_key=api_key,
            base_url=base_url
        )
        
        self.model = model
        
        # 添加处理控制相关属性
        self.stop_processing = False
        self.processing_stats = {}  # 记录处理统计信息
        
        # 默认参数配置
        self.max_workers = 3  # 默认并发数
        self.request_delay = 0.5  # 请求间延迟（秒）
        self.max_retries = 2  # 最大重试次数
        
    def set_processing_params(self, max_workers=3, request_delay=0.5, max_retries=2):
        """设置处理参数"""
        self.max_workers = max_workers
        self.request_delay = request_delay
        self.max_retries = max_retries
        
    def stop_current_processing(self):
        """停止当前处理"""
        self.stop_processing = True
        
    def reset_stop_flag(self):
        """重置停止标志"""
        self.stop_processing = False
        
    def get_column_processing_status(self, dataframe, column_name, table_manager=None):
        """获取指定列的处理状态统计"""
        if dataframe is None or column_name not in dataframe.columns:
            return None
        
        is_multi_field = False
        output_fields = []
        if table_manager:
            ai_config = table_manager.get_ai_columns().get(column_name)
            if isinstance(ai_config, dict):
                if ai_config.get("output_mode") == "multi" and ai_config.get("output_fields"):
                    is_multi_field = True
                    output_fields = ai_config.get("output_fields")

        total_rows = len(dataframe)
        # Counters for the sum total_rows = processed + empty + failed
        current_processed_count = 0
        current_failed_count = 0
        current_empty_count = 0 
        
        # Informational counters and row lists
        partially_processed_count_info = 0 
        failed_rows_info = []
        empty_rows_info = [] # Will include partially processed rows for reprocessing
        partially_processed_rows_info = []

        for i in range(total_rows):
            main_cell_value = str(dataframe.iloc[i].get(column_name, "")).strip()

            if main_cell_value.startswith('错误:') or 'API调用失败' in main_cell_value:
                current_failed_count += 1
                failed_rows_info.append(i + 1)
            elif not main_cell_value: # Main cell is completely empty
                if not is_multi_field:
                    current_empty_count += 1
                    empty_rows_info.append(i + 1)
                else: # Multi-field, main cell empty
                    all_sub_fields_empty_for_this_row = True
                    for field in output_fields:
                        sub_col_name = f"{column_name}_{field}"
                        if sub_col_name in dataframe.columns and str(dataframe.iloc[i].get(sub_col_name, "")).strip():
                            all_sub_fields_empty_for_this_row = False
                            break
                    if all_sub_fields_empty_for_this_row:
                        current_empty_count += 1
                        empty_rows_info.append(i + 1)
                    else:
                        # Main empty, but some sub-fields have data. Odd case, count as processed.
                        current_processed_count +=1
            else: # Main cell has content and is not an explicit error
                if not is_multi_field:
                    current_processed_count += 1
                else: # Multi-field: main cell has content
                    any_sub_field_populated = False
                    for field in output_fields:
                        sub_col_name = f"{column_name}_{field}"
                        if sub_col_name in dataframe.columns and str(dataframe.iloc[i].get(sub_col_name, "")).strip():
                            any_sub_field_populated = True
                            break
                    
                    if any_sub_field_populated:
                        current_processed_count += 1
                    else:
                        # Main has content, but all sub-fields are empty.
                        # This is "partially processed" and should be reprocessable via "empty".
                        current_empty_count += 1
                        empty_rows_info.append(i + 1)
                        # Informational tracking for this specific state:
                        partially_processed_count_info +=1
                        partially_processed_rows_info.append(i+1)
        
        return {
            'total_rows': total_rows,
            'processed_count': current_processed_count,
            'failed_count': current_failed_count,
            'empty_count': current_empty_count, 
            'partially_processed_count': partially_processed_count_info, # Informational
            'failed_rows': failed_rows_info,
            'empty_rows': empty_rows_info, # Includes rows from partially_processed_rows_info
            'partially_processed_rows': partially_processed_rows_info, # Informational
            'completion_rate': round((current_processed_count / total_rows * 100), 1) if total_rows > 0 else 0
        }
        
    def process_single_cell(self, dataframe, row_index, column_name, prompt_template, model=None, table_manager=None, output_fields=None):
        """处理单个单元格"""
        try:
            # 检查是否需要停止
            if self.stop_processing:
                return False, "处理已停止"
                
            # 获取行数据
            row_data = dataframe.iloc[row_index].to_dict()
            
            # 替换模板中的变量（包括长文本列）
            prompt = self.replace_template_variables(prompt_template, row_data, row_index, table_manager)
            
            # 检查prompt中是否包含空字段标记
            if "{字段为空:" in prompt:
                error_msg = f"Prompt包含空字段，无法处理: {prompt}"
                print(f"处理第{row_index+1}行失败: {error_msg}")
                dataframe.loc[row_index, column_name] = f"错误: {error_msg}"
                return False, error_msg
            
            # 检查prompt中是否包含未找到字段标记
            if "{未找到字段:" in prompt:
                error_msg = f"Prompt包含未找到的字段: {prompt}"
                print(f"处理第{row_index+1}行失败: {error_msg}")
                dataframe.loc[row_index, column_name] = f"错误: {error_msg}"
                return False, error_msg
            
            # 使用指定模型或默认模型
            use_model = model if model else self.model
            
            print(f"处理第{row_index+1}行，列：{column_name} (模型: {use_model})")
            print(f"Prompt: {prompt[:200]}..." if len(prompt) > 200 else f"Prompt: {prompt}")
            
            # 调用AI API（带重试机制）
            result = self.call_ai_api_with_retry(prompt, use_model)
            
            print(f"AI结果: {result}")
            
            # 处理多字段输出
            # 检查是否为多字段模式（可能是预定义字段或自动解析）
            is_multi_mode = False
            field_mode = "predefined"
            if table_manager:
                ai_config = table_manager.get_ai_columns().get(column_name, {})
                if isinstance(ai_config, dict):
                    output_mode = ai_config.get("output_mode", "single")
                    field_mode = ai_config.get("field_mode", "predefined")
                    is_multi_mode = (output_mode == "multi")
            
            # 如果是多字段模式（预定义字段模式需要有字段，自动解析模式不需要）
            if is_multi_mode and (output_fields or field_mode == "auto_parse"):
                
                success, parsed_results = self.parse_multi_field_response(result, output_fields, column_name, field_mode)
                
                # 添加调试信息
                print(f"DEBUG: Multi-field parsing - Column: {column_name}")
                print(f"DEBUG: Field mode: {field_mode}")
                print(f"DEBUG: Expected fields: {output_fields}")
                print(f"DEBUG: AI Response: {result[:500]}...")
                print(f"DEBUG: Parse success: {success}")
                print(f"DEBUG: Parsed results: {parsed_results}")
                
                if success:
                    # 对于自动解析模式，需要动态创建字段列
                    if field_mode == "auto_parse":
                        # 使用解析出的字段名作为实际字段
                        actual_fields = list(parsed_results.keys())
                        if table_manager:
                            table_manager.ensure_multi_field_columns_positioned(column_name, actual_fields)
                    else:
                        # 预定义模式，使用预设字段
                        if table_manager:
                            table_manager.ensure_multi_field_columns_positioned(column_name, output_fields)
                    
                    # 创建/更新多个列
                    print(f"DEBUG: Starting to update {len(parsed_results)} field columns")
                    for field_name, field_value in parsed_results.items():
                        full_column_name = f"{column_name}_{field_name}"
                        print(f"DEBUG: Processing field '{field_name}' -> column '{full_column_name}' with value '{field_value}'")
                        
                        # 确保列存在（如果table_manager不可用，回退到原方法）
                        if full_column_name not in dataframe.columns:
                            print(f"DEBUG: Creating new column '{full_column_name}'")
                            dataframe[full_column_name] = ""
                        else:
                            print(f"DEBUG: Column '{full_column_name}' already exists")
                            
                        # 更新单元格值
                        dataframe.loc[row_index, full_column_name] = field_value
                        print(f"DEBUG: Updated cell [{row_index}, {full_column_name}] = '{field_value}'")
                    
                    # 在主列中存储原始响应（可选）
                    dataframe.loc[row_index, column_name] = result
                    
                    return True, parsed_results
                else:
                    # 解析失败，将原始响应存储到主列并标记
                    error_msg = f"JSON解析失败: {parsed_results}"
                    dataframe.loc[row_index, column_name] = f"{result}\n\n[解析错误: {parsed_results}]"
                    return False, error_msg
            else:
                # 单字段模式，直接更新
                dataframe.loc[row_index, column_name] = result
                return True, result
            
        except Exception as e:
            error_msg = f"错误: {str(e)}"
            print(f"处理单元格时出错: {e}")
            dataframe.loc[row_index, column_name] = error_msg
            return False, error_msg
    
    def call_ai_api_with_retry(self, prompt, model=None, retries=None):
        """带重试机制的AI API调用"""
        if retries is None:
            retries = self.max_retries
            
        last_error = None
        
        for attempt in range(retries + 1):
            try:
                if attempt > 0:
                    print(f"重试第{attempt}次...")
                    time.sleep(1)  # 重试前等待
                    
                result = self.call_ai_api(prompt, model)
                return result
                
            except Exception as e:
                last_error = e
                if attempt < retries:
                    print(f"API调用失败，将重试: {str(e)}")
                    continue
                else:
                    print(f"API调用失败，已达最大重试次数: {str(e)}")
                    
        raise last_error
        
    def process_column_concurrent(self, dataframe, column_name, prompt_template, model=None, 
                                table_manager=None, progress_callback=None, rows_to_process=None, output_fields=None):
        """并发处理指定列的多行数据"""
        try:
            self.reset_stop_flag()
            
            # 如果未指定处理行，则处理所有行
            if rows_to_process is None:
                rows_to_process = list(range(len(dataframe)))
            
            total_tasks = len(rows_to_process)
            completed_tasks = 0
            success_count = 0
            failed_count = 0
            
            # 线程安全的结果收集
            results = {}
            results_lock = threading.Lock()
            
            def process_single_task(row_index):
                """处理单个任务的工作函数"""
                if self.stop_processing:
                    return row_index, False, "处理已停止"
                    
                try:
                    success, result = self.process_single_cell(
                        dataframe, row_index, column_name, prompt_template, model, table_manager, output_fields
                    )
                    
                    with results_lock:
                        results[row_index] = (success, result)
                        
                    # 添加延迟避免API限制
                    time.sleep(self.request_delay)
                    
                    return row_index, success, result
                    
                except Exception as e:
                    error_msg = f"错误: {str(e)}"
                    with results_lock:
                        results[row_index] = (False, error_msg)
                    return row_index, False, error_msg
            
            # 使用线程池进行并发处理
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # 提交任务
                future_to_row = {
                    executor.submit(process_single_task, row_index): row_index 
                    for row_index in rows_to_process
                }
                
                # 收集结果
                for future in as_completed(future_to_row):
                    if self.stop_processing:
                        break
                        
                    row_index = future_to_row[future]
                    try:
                        _, success, result = future.result()
                        
                        completed_tasks += 1
                        if success:
                            success_count += 1
                        else:
                            failed_count += 1
                            
                        # 更新进度
                        if progress_callback:
                            progress_callback(completed_tasks, total_tasks, success_count, failed_count)
                            
                    except Exception as e:
                        print(f"处理第{row_index+1}行时发生异常: {e}")
                        failed_count += 1
                        completed_tasks += 1
                        
                        if progress_callback:
                            progress_callback(completed_tasks, total_tasks, success_count, failed_count)
            
            # 保存处理统计
            self.processing_stats = {
                'total_tasks': total_tasks,
                'completed_tasks': completed_tasks,
                'success_count': success_count,
                'failed_count': failed_count,
                'stopped': self.stop_processing
            }
            
            return True, self.processing_stats
            
        except Exception as e:
            print(f"并发处理错误: {e}")
            return False, str(e)
        
    def process_batch(self, dataframe, ai_columns, progress_callback=None):
        """批量处理AI列"""
        try:
            total_tasks = len(dataframe) * len(ai_columns)
            current_task = 0
            
            # 遍历每一行
            for row_index in range(len(dataframe)):
                if self.stop_processing:
                    break
                    
                row_data = dataframe.iloc[row_index].to_dict()
                
                # 处理每个AI列
                for column_name, prompt_template in ai_columns.items():
                    if self.stop_processing:
                        break
                        
                    success, result = self.process_single_cell(dataframe, row_index, column_name, prompt_template)
                    
                    # 更新进度
                    current_task += 1
                    if progress_callback:
                        progress_callback(current_task, total_tasks)
                        
                    # 添加延迟避免API限制
                    time.sleep(self.request_delay)
                        
            return True
            
        except Exception as e:
            print(f"批量处理错误: {e}")
            return False
            
    def replace_template_variables(self, template, row_data, row_index=None, table_manager=None):
        """替换模板中的变量（包括长文本列）"""
        # 使用正则表达式找到所有 {变量名} 格式的占位符
        def replace_var(match):
            var_name = match.group(1)
            
            # 首先检查是否为长文本列
            if (table_manager and 
                table_manager.is_long_text_column(var_name) and 
                row_index is not None):
                # 获取完整的长文本内容
                long_text_content = table_manager.get_long_text_content(var_name, row_index)
                if long_text_content:
                    return long_text_content
                else:
                    return f"{{长文本列内容获取失败: {var_name}}}"
            
            # 普通字段替换
            field_value = row_data.get(var_name, f"{{未找到字段: {var_name}}}")
            
            # 检查字段值是否为空
            if not field_value or str(field_value).strip() == '':
                return f"{{字段为空: {var_name}}}"
            
            return str(field_value)
            
        return re.sub(r'\{(\w+)\}', replace_var, template)
        
    def call_ai_api(self, prompt, model=None):
        """调用AI API"""
        try:
            use_model = model if model else self.model
            
            # 根据模型调整参数
            if use_model == "o1":
                # o1模型的特殊配置
                response = self.client.chat.completions.create(
                    model=use_model,
                    messages=[
                        {"role": "user", "content": prompt}
                    ]
                    # o1模型不支持temperature和max_tokens参数
                )
            else:
                # 其他模型的标准配置
                response = self.client.chat.completions.create(
                    model=use_model,
                    messages=[
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=1000,
                    temperature=0.7
                )
            
            return response.choices[0].message.content.strip()
            
        except Exception as e:
            raise Exception(f"AI API调用失败: {str(e)}")
            
    def test_connection(self):
        """测试AI API连接"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "user", "content": "测试连接，请回复'连接成功'"}
                ],
                max_tokens=10
            )
            
            return True, response.choices[0].message.content.strip()
            
        except Exception as e:
            return False, str(e)

    def parse_multi_field_response(self, response, expected_fields, column_name, field_mode="predefined"):
        """解析多字段JSON响应"""
        try:
            import json
            import re
            
            print(f"DEBUG: parse_multi_field_response called")
            print(f"DEBUG: Original response: {response}")
            print(f"DEBUG: Expected fields: {expected_fields}")
            print(f"DEBUG: Field mode: {field_mode}")
            
            # 清理响应文本，提取JSON部分
            cleaned_response = self.extract_json_from_response(response)
            print(f"DEBUG: Cleaned response: {cleaned_response}")
            
            # 尝试解析JSON
            try:
                parsed_data = json.loads(cleaned_response)
                print(f"DEBUG: Parsed JSON data: {parsed_data}")
            except json.JSONDecodeError as e:
                print(f"DEBUG: JSON parse error: {str(e)}")
                return False, f"JSON格式错误: {str(e)}"
            
            # 检查是否为字典
            if not isinstance(parsed_data, dict):
                return False, "响应不是JSON对象格式"
            
            result = {}
            
            if field_mode == "auto_parse":
                # 自动解析模式：提取JSON中的所有字段
                print(f"DEBUG: Auto-parse mode, processing {len(parsed_data)} fields")
                for key, value in parsed_data.items():
                    # 过滤掉空值和无效字段
                    if value is not None and str(value).strip():
                        result[key] = str(value).strip()
                        print(f"DEBUG: Added field '{key}': '{result[key]}'")
                    else:
                        print(f"DEBUG: Skipped empty field '{key}': '{value}'")
                
                if not result:
                    print(f"DEBUG: No valid fields found in auto-parse mode")
                    return False, "JSON中没有找到有效字段"
                    
                print(f"DEBUG: Auto-parse result: {result}")
                return True, result
            else:
                # 预定义字段模式：只提取期望的字段
                print(f"DEBUG: Predefined mode, looking for fields: {expected_fields}")
                missing_fields = []
                
                for field in expected_fields:
                    if field in parsed_data:
                        result[field] = str(parsed_data[field]).strip()
                        print(f"DEBUG: Found field '{field}': '{result[field]}'")
                    else:
                        missing_fields.append(field)
                        print(f"DEBUG: Missing field '{field}'")
                
                if missing_fields:
                    # 部分字段缺失，但尝试使用可用的字段
                    print(f"警告: 缺失字段 {missing_fields}，但继续使用可用字段")
                    for field in missing_fields:
                        result[field] = "[字段缺失]"
                
                if not result:
                    print(f"DEBUG: No expected fields found in predefined mode")
                    return False, "没有找到任何期望的字段"
                    
                print(f"DEBUG: Predefined result: {result}")
                return True, result
            
        except Exception as e:
            return False, f"解析异常: {str(e)}"
    
    def extract_json_from_response(self, response):
        """从响应中提取JSON部分"""
        import re
        
        # 尝试找到JSON对象（以{开始，以}结束）
        json_pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
        
        # 查找所有可能的JSON对象
        json_matches = re.findall(json_pattern, response, re.DOTALL)
        
        if json_matches:
            # 返回第一个找到的JSON对象
            return json_matches[0].strip()
        
        # 如果没有找到完整的JSON对象，尝试清理并返回原文本
        # 移除多余的文本，只保留可能的JSON部分
        lines = response.split('\n')
        json_lines = []
        in_json = False
        
        for line in lines:
            line = line.strip()
            if line.startswith('{') or in_json:
                in_json = True
                json_lines.append(line)
                if line.endswith('}') and line.count('}') >= line.count('{'):
                    break
        
        if json_lines:
            return '\n'.join(json_lines)
        
        return response.strip()
    
    def retry_parse_single_cell(self, dataframe, row_index, column_name, expected_fields):
        """重试解析单个单元格的JSON响应"""
        try:
            # 获取当前单元格的原始响应
            current_value = str(dataframe.loc[row_index, column_name])
            
            # 检查是否包含解析错误标记
            if "[解析错误:" in current_value:
                # 提取原始响应（去掉错误标记）
                original_response = current_value.split("\n\n[解析错误:")[0]
            else:
                original_response = current_value
            
            # 重新尝试解析
            success, parsed_results = self.parse_multi_field_response(original_response, expected_fields, column_name)
            
            if success:
                # 更新多个字段列
                for field_name, field_value in parsed_results.items():
                    full_column_name = f"{column_name}_{field_name}"
                    if full_column_name not in dataframe.columns:
                        dataframe[full_column_name] = ""
                    dataframe.loc[row_index, full_column_name] = field_value
                
                # 更新主列（移除错误标记）
                dataframe.loc[row_index, column_name] = original_response
                
                return True, f"重新解析成功，提取了 {len(parsed_results)} 个字段"
            else:
                return False, f"重新解析仍然失败: {parsed_results}"
                
        except Exception as e:
            return False, f"重试解析时出错: {str(e)}"

    def build_full_prompt(self, row_index, col_name, table_manager):
        """构建发送给AI模型的完整Prompt"""
        prompt_template = table_manager.get_ai_column_prompt(col_name)
        row_data = table_manager.get_row_data(row_index)
        full_prompt = self.replace_template_variables(prompt_template, row_data, row_index, table_manager)
        return full_prompt 