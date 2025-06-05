import os
import openai
from dotenv import load_dotenv
import sys

# 调试模式开关
DEBUG = True  # 设置为True启用详细日志，False则仅显示关键信息

# print("--- API Test Script --- ")
# KEY = os.getenv("OPENAI_API_KEY")
# URL = os.getenv("OPENAI_BASE_URL")
# """
# 测试openai的api是否正常
# """

# def test_upload_pdf_with_url():
#     """
# (dataprocess-py310) D:\workspace\evalDRWeb>python test_anthropic_api.py
# [TextBlock(citations=None, text=" I'd be happy to analyze key findings from a document, but I don't see any document attached to your message. You can:\n\n1. Upload the document you'd like me to analyze\n2. Copy and paste the text content\n3. Provide a link if it's a publicly available document\n\nOnce you share the document, I can help identify the main points, key findings, and important insights for you.", type='text')]
#     """
#     client = openai.OpenAI(api_key=KEY, base_url=URL, timeout=60.0, max_retries=2)
#     message = client.messages.create(
#         model="claude-3-7-sonnet-20250219",
#         max_tokens=1024,
#         messages=[
#             {
#                 "role": "user",
#                 "content": [
#                     {
#                         "type": "document",
#                         "source": {
#                             "type": "url",
#                             "url": "https://assets.anthropic.com/m/1cd9d098ac3e6467/original/Claude-3-Model-Card-October-Addendum.pdf"
#                         }
#                     },
#                     {
#                         "type": "text",
#                         "text": "What are the key findings in this document?"
#                     }
#                 ]
#             }
#         ],
#     )
#     print(message.content)

# test_upload_pdf_with_url()





# --- Path Setup ---
# Calculate project root assuming this script is in '/03/tool_box/'
TOOL_BOX_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(TOOL_BOX_DIR)
if DEBUG:
    print(f"Project Root detected as: {PROJECT_ROOT}")

# --- Load Environment Variables from .env file in Project Root ---
dotenv_path = os.path.join(PROJECT_ROOT, '.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path=dotenv_path, override=True)
    if DEBUG:
        print(f"Loaded environment variables from: {dotenv_path}")
else:
    print("Warning: .env file not found in project root. Using system environment variables or defaults.")


# --- Get API Credentials ---
# KEY = os.getenv("OPENAI_API_KEY")
# URL = os.getenv("OPENAI_BASE_URL")
# MODEL = os.getenv("OPENAI_MODEL") # 从环境中获取model名称

# KEY="sk-CEvlVAAKY600R7oRXxg9ChZ6jSOE28tDoqK7XnnGbWrZIExd"
# URL="https://vip-api-global.aiearth.dev/v1"
KEY="sk-FastAPIHY1gN0kn1lDhB8fksC1ozs1Az0Juz3HxQ8vKhc6ZX"
# # # # OPENAI_BASE_URL=https://mtu.mtuopenai.xyz/v1/
URL="https://mtu2.mtuopenai.xyz/v1/"
# MODEL = "gpt-4.1" # 从环境中获取model名称
MODEL = "claude-3-7-sonnet-20250219"
if DEBUG:
    print(f"Using Base URL: {URL}")
    print(f"API Key Loaded: {'Yes' if KEY else 'No'}")
print(f"使用模型: {MODEL}")

# --- Check Credentials and Initialize Client ---
client = None
if not KEY or not URL:
    print("错误: OPENAI_API_KEY 或 OPENAI_BASE_URL 未在环境变量或.env文件中找到。")
    print("请确保它们在.env文件中正确设置。")
else:
    try:
        if DEBUG:
            print("\n初始化 OpenAI 客户端...")
        client = openai.OpenAI(api_key=KEY, base_url=URL, timeout=60.0, max_retries=2) # Shorter timeout for test
        if DEBUG:
            print("OpenAI 客户端初始化成功。")
    except Exception as e:
        print(f"错误: 初始化 OpenAI 客户端失败: {e}")

# --- Perform Test API Call --- 
if client:
    try:
        if DEBUG:
            print(f"\n尝试调用模型 '{MODEL}'...")
        test_prompt = "介绍一下你自己"
        
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": "You are a simple test assistant."}, 
                {"role": "user", "content": test_prompt}
            ],
            max_tokens=500, # Small max tokens for a simple test
            # temperature=1,
        )
        
        api_response = response.choices[0].message.content.strip()
        print("\n--- API调用成功！---")
        print(f"模型回应: {api_response}")
        
    except openai.AuthenticationError as e:
         print(f"\n--- API调用失败: 认证错误 ---")
         print(f"错误详情: {e}")
         print("请检查您的API密钥是否正确且有效。")
    except openai.APIConnectionError as e:
         print(f"\n--- API调用失败: 连接错误 ---")
         print(f"错误详情: {e}")
         print(f"无法连接到API端点 ({URL})。请检查URL和网络连接。")
    except openai.RateLimitError as e:
        print(f"\n--- API调用失败: 速率限制错误 ---")
        print(f"错误详情: {e}")
        print("您已超出API使用配额。请检查您的计划或稍后再试。")
    except openai.APITimeoutError as e:
        print(f"\n--- API调用失败: 超时错误 ---")
        print(f"错误详情: {e}")
        print("请求花费的时间过长。")
    except Exception as e:
        print(f"\n--- API调用失败: 意外错误 ---")
        print(f"错误类型: {type(e).__name__}")
        print(f"错误详情: {e}")
else:
    print("\n跳过API调用，因为客户端无法初始化。")

print("\n--- 测试脚本完成 ---") 