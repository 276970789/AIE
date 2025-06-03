# AI Excel 批量数据处理工具

一个功能强大的AI驱动的Excel数据处理工具，集成了智能对话、表格管理和项目管理功能，支持批量数据处理和AI自动化分析。

## ✨ 核心特性

### 🤖 AI智能处理
- **多模型支持**: 支持GPT-4.1和O1模型，满足不同场景需求
- **智能列处理**: 创建AI列，使用自定义prompt模板批量处理数据
- **模板变量**: 支持`{列名}`语法引用其他字段值，实现动态数据处理
- **批量处理**: 一键处理整列或选定行的AI任务

### 📊 表格管理
- **多格式支持**: 支持Excel (.xlsx/.xls)、CSV、JSONL文件格式
- **智能编码**: 自动识别文件编码（UTF-8、GBK、GB2312等）
- **实时编辑**: 支持单元格编辑、行列操作、数据排序
- **列管理**: 支持列的增删改、重命名、类型转换

### 💼 项目管理
- **项目保存**: 以.aie格式保存项目，包含数据、AI配置和界面状态
- **配置管理**: 保存AI列的prompt模板和模型配置
- **版本控制**: 支持项目版本管理和兼容性检查

### 🎨 现代化界面
- **响应式设计**: 支持窗口缩放和自适应布局
- **主题样式**: 现代化的UI设计，支持多种视觉主题
- **交互优化**: 右键菜单、拖拽操作、快捷键支持
- **进度显示**: 实时显示AI处理进度和状态

## 🏗️ 技术架构

```
ai_excel_tool/
├── main.py                 # 主程序入口，GUI界面
├── ai_processor.py         # AI处理核心，OpenAI API集成
├── table_manager.py        # 表格数据管理，文件I/O
├── project_manager.py      # 项目管理，配置保存/加载
├── ai_column_dialog.py     # AI列配置对话框
├── requirements.txt        # 项目依赖
├── start_ai_excel.bat     # Windows启动脚本
├── data/                  # 数据文件目录
├── docs/                  # 文档目录
├── tests/                 # 测试文件
└── 论文/                  # 学术资料目录
    ├── bio/              # 生物学相关
    ├── chemistry/        # 化学相关
    ├── economy/          # 经济学相关
    ├── law/              # 法学相关
    └── philosophy/       # 哲学相关
```

## 🚀 快速开始

### 环境要求
- Python 3.7+
- Windows/macOS/Linux

### 安装步骤

1. **克隆项目**
```bash
git clone <repository-url>
cd ai_excel_tool
```

2. **安装依赖**
```bash
pip install -r requirements.txt
```

3. **配置AI服务**
创建`.env`文件并配置OpenAI API：
```env
OPENAI_API_KEY=your_api_key_here
OPENAI_BASE_URL=https://api.openai.com/v1
```

4. **启动应用**
```bash
# 方式1: 直接运行Python
python main.py

# 方式2: 使用批处理文件 (Windows)
start_ai_excel.bat
```

## 📖 使用指南

### 基础操作

1. **导入数据**
   - 点击"导入数据"按钮
   - 支持Excel、CSV、JSONL格式
   - 自动识别文件编码

2. **创建AI列**
   - 点击"新建列" → 选择"AI处理列"
   - 输入列名和prompt模板
   - 选择AI模型（GPT-4.1或O1）

3. **配置Prompt模板**
```
示例模板：
请将以下{类别}类的英文内容翻译成中文：{文本内容}

可用变量：
- {列名}: 引用其他列的值
- 支持多变量组合使用
```

4. **批量处理**
   - 选择AI列 → 右键 → "处理整列"
   - 或选择特定行进行处理
   - 实时查看处理进度

### 高级功能

#### 项目管理
- **保存项目**: 文件 → 保存项目 (.aie格式)
- **加载项目**: 文件 → 打开项目
- **项目包含**: 数据、AI配置、界面状态

#### 数据导出
- **Excel导出**: 保持格式和样式
- **CSV导出**: 兼容各种编码
- **JSONL导出**: 适合机器学习数据集
- **选择性导出**: 仅导出指定列

#### 界面定制
- **行高调节**: 低/中/高三种模式
- **列宽调整**: 拖拽调整或自动适应
- **排序功能**: 点击列头进行排序

## ⚙️ 配置说明

### AI模型配置

| 模型 | 特点 | 适用场景 |
|------|------|----------|
| GPT-4.1 | 快速响应，平衡性能 | 日常文本处理、翻译、分类 |
| O1 | 深度推理，复杂分析 | 复杂逻辑推理、学术分析 |

### 环境变量

```env
# OpenAI API配置
OPENAI_API_KEY=sk-xxx...           # API密钥
OPENAI_BASE_URL=https://api.openai.com/v1  # API基础URL
OPENAI_MODEL=gpt-4.1               # 默认模型

# 可选配置
AI_TIMEOUT=30                      # API超时时间(秒)
MAX_RETRIES=3                      # 最大重试次数
```

### 项目文件格式(.aie)

```json
{
  "format_version": "1.0",
  "created_at": "2024-01-01T00:00:00",
  "app_version": "2.2",
  "table_data": {
    "columns": ["列1", "列2", "AI列"],
    "data": [...],
    "row_count": 100,
    "col_count": 3
  },
  "ai_config": {
    "ai_columns": {
      "AI列": {
        "prompt": "处理模板",
        "model": "gpt-4.1"
      }
    }
  },
  "ui_state": {
    "column_widths": {...},
    "row_height_setting": "low"
  }
}
```

## 🔧 开发指南

### 核心模块

1. **AIProcessor**: AI处理核心
   - OpenAI API集成
   - 模板变量替换
   - 批量处理逻辑

2. **TableManager**: 数据管理
   - 文件I/O操作
   - 数据框操作
   - AI列配置管理

3. **ProjectManager**: 项目管理
   - 项目序列化/反序列化
   - 版本兼容性处理
   - 配置管理

### 扩展开发

添加新的AI模型支持：
```python
# 在ai_processor.py中添加
def call_ai_api(self, prompt, model=None):
    if model == "new_model":
        # 新模型的API调用逻辑
        pass
```

添加新的文件格式支持：
```python
# 在table_manager.py中添加
def load_file(self, file_path):
    if file_ext == '.new_format':
        # 新格式的加载逻辑
        pass
```

## 🐛 故障排除

### 常见问题

1. **API连接失败**
   - 检查网络连接
   - 验证API密钥和URL配置
   - 确认API服务可用性

2. **文件编码问题**
   - 工具会自动尝试多种编码
   - 建议使用UTF-8编码保存文件

3. **内存不足**
   - 大文件建议分批处理
   - 关闭不必要的应用程序

4. **界面显示异常**
   - 检查系统字体支持
   - 尝试调整界面缩放比例

### 日志调试

程序运行时会在控制台输出详细日志：
```
API Key: sk-xxx...
Base URL: https://api.openai.com/v1
Model: gpt-4.1
处理第1行，列：翻译结果 (模型: gpt-4.1)
Prompt: 请翻译：Hello World
AI结果: 你好世界
```

## 📄 许可证

本项目采用MIT许可证，详见LICENSE文件。

## 🤝 贡献指南

欢迎提交Issue和Pull Request！

1. Fork本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启Pull Request

## 📞 支持与反馈

- 🐛 Bug报告: [Issues](https://github.com/your-repo/issues)
- 💡 功能建议: [Discussions](https://github.com/your-repo/discussions)
- 📧 邮件联系: your-email@example.com

---

**AI Excel 批量数据处理工具** - 让数据处理更智能，让工作更高效！ 