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
        
    def process_single_cell(self, dataframe, row_index, column_name, prompt_template, model=None):
        """处理单个单元格"""
        try:
            # 获取行数据
            row_data = dataframe.iloc[row_index].to_dict()
            
            # 替换模板中的变量
            prompt = self.replace_template_variables(prompt_template, row_data)
            
            # 使用指定模型或默认模型
            use_model = model if model else self.model
            
            print(f"处理第{row_index+1}行，列：{column_name} (模型: {use_model})")
            print(f"Prompt: {prompt}")
            
            # 调用AI API
            result = self.call_ai_api(prompt, use_model)
            
            print(f"AI结果: {result}")
            
            # 更新数据框
            dataframe.loc[row_index, column_name] = result
            
            return True, result
            
        except Exception as e:
            error_msg = f"错误: {str(e)}"
            print(f"处理单元格时出错: {e}")
            dataframe.loc[row_index, column_name] = error_msg
            return False, error_msg
        
    def process_batch(self, dataframe, ai_columns, progress_callback=None):
        """批量处理AI列"""
        try:
            total_tasks = len(dataframe) * len(ai_columns)
            current_task = 0
            
            # 遍历每一行
            for row_index in range(len(dataframe)):
                row_data = dataframe.iloc[row_index].to_dict()
                
                # 处理每个AI列
                for column_name, prompt_template in ai_columns.items():
                    success, result = self.process_single_cell(dataframe, row_index, column_name, prompt_template)
                    
                    # 更新进度
                    current_task += 1
                    if progress_callback:
                        progress_callback(current_task, total_tasks)
                        
                    # 添加延迟避免API限制
                    time.sleep(0.1)
                        
            return True
            
        except Exception as e:
            print(f"批量处理错误: {e}")
            return False
            
    def replace_template_variables(self, template, row_data):
        """替换模板中的变量"""
        # 使用正则表达式找到所有 {变量名} 格式的占位符
        def replace_var(match):
            var_name = match.group(1)
            return str(row_data.get(var_name, f"{{未找到字段: {var_name}}}"))
            
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