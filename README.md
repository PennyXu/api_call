# P_chat_to_excel 智能处理工具

一个基于 RAG（检索增强生成）和 OpenAI API 的 Excel 数据处理工具，能够智能分析表格内容并生成文本结果。

## 功能特性

- 📊 **Excel 智能处理** - 自动解析和处理 Excel 表格数据
- 🤖 **AI 对话集成** - 集成 OpenAI 兼容 API，支持多种大语言模型
- 🔍 **RAG 知识库** - 支持 PDF 文档构建知识库，增强回答准确性
- ⚡ **并行处理** - 多线程并发处理大量数据，提升效率
- 💾 **持久化配置** - 自动保存和加载提示词配置
- 📈 **进度显示** - 实时显示处理进度和状态

## 安装依赖

```bash
pip install openai pandas tqdm sentence-transformers faiss-cpu langchain pypdf2
```

## 快速开始

### 1. 基础使用

```python
# 初始化工具
chat_tool = P_chat_to_excel(
    name="my_project",
    api_key="your-api-key",
    base_url="https://api.openai.com/v1"
)

# 加载 Excel 文件
chat_tool.excel_info("data.xlsx", "content_column")

# 配置处理参数
chat_tool.info_collect()
# 按照提示输入：
# prompt: 你是一个数据分析助手
# inquiry: 分析这段文本的情感倾向
# column: content
# result_column_name: 情感分析结果
# file_path: result.xlsx

# 执行处理
chat_tool.chat()
```

### 2. 使用 RAG 知识库

```python
# 构建知识库
kb = P_RAGKnowledgeBase()
kb.build_from_pdf("reference.pdf")

# 使用知识库增强处理
chat_tool.chat(kb=kb, top_k=3)
```

## 核心组件

### P_chat_to_excel 主类

**初始化参数：**
- `name`: 项目名称，用于配置持久化
- `api_key`: OpenAI API 密钥
- `base_url`: API 基础地址

**主要方法：**

- `excel_info(path, column)`: 加载 Excel 文件并指定目标列
- `data_parsing(column)`: 解析 JSON 格式的文本数据
- `concat(columns)`: 合并多个列的内容
- `data_sample(num)`: 随机采样指定数量的行
- `info_collect()`: 交互式配置处理参数
- `chat()`: 执行 AI 处理

### P_RAGKnowledgeBase 知识库类

```python
kb = P_RAGKnowledgeBase(
    chunk_size=500,
    chunk_overlap=50,
    embedding_model_name='all-MiniLM-L6-v2'
)
kb.build_from_pdf("document.pdf")
```

### P_RAGRetriever 检索器类

```python
retriever = P_RAGRetriever(knowledge_base=kb, openai_client=client, top_k=3)
```

## 配置说明

### API 配置

支持所有 OpenAI 兼容的 API：
- 深度求索 (DeepSeek)
- 通义千问 (Qwen)
- OpenAI
- 其他兼容服务

```python
client = P_chat_to_excel(
    name="project",
    api_key="your-token",
    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"  # 通义千问
)
```

### 处理模式

1. **单列处理** (`axis=None`)
   - 对单个目标列进行处理
   - 结果保存到新列

2. **多列并行处理** (`axis=1`)
   - 对每列独立处理
   - 每列后插入结果列

3. **多列合并处理** (`axis=0`)
   - 合并所有列内容
   - 统一处理并保存结果

## 高级用法

### 批量处理配置

```python
# 直接设置参数（避免交互）
chat_tool.prompt = "你是一个专业的产品评论分析师"
chat_tool.inquiry = "提取这段评论中的产品特征和用户情感"
chat_tool.column = "review_content"
chat_tool.result_column_name = "分析结果"
chat_tool.file_path = "analysis_result.xlsx"
```

### 采样测试

```python
# 先采样测试
chat_tool.data_sample(10)  # 随机取10行
chat_tool.chat(sample=True)  # 在采样数据上测试模型效果

# 确认效果后处理全量数据
chat_tool.chat(sample=False)
```

### 自定义模型参数

```python
chat_tool.chat(
    temperature=0.7,      # 创造性 (0-1)
    top_p=0.9,           # 核采样参数
    frequency_penalty=0.2, # 频率惩罚
    max_workers=10        # 并发线程数
)
```

## 示例场景

### 1. 客户评论分析

```python
chat_tool.prompt = "你是客户体验分析师，擅长从评论中提取关键信息"
chat_tool.inquiry = "从评论中提取：1.提到的问题 2.情感倾向 3.改进建议"
chat_tool.chat()
```

### 2. 产品描述生成

```python
chat_tool.prompt = "你是电商文案写手，擅长撰写吸引人的产品描述"
chat_tool.inquiry = "根据产品特性生成一段营销文案"
chat_tool.chat(kb=kb)  # 使用产品手册作为知识库
```

### 3. 数据分类标注

```python
chat_tool.prompt = "你是数据标注专家"
chat_tool.inquiry = "将文本分类为：正面、负面、中性"
chat_tool.chat()
```

## 注意事项

1. **API 限制**：注意调用频率和配额限制
2. **数据安全**：敏感数据请使用本地部署的模型
3. **内存管理**：处理大数据集时注意内存使用
4. **错误处理**：网络异常时会自动重试并记录空结果

## 故障排除

### 常见问题

1. **API 连接失败**
   - 检查 API 密钥和基础地址
   - 验证网络连接

2. **内存不足**
   - 减少 `max_workers` 参数
   - 使用 `data_sample()` 分批处理

3. **处理速度慢**
   - 调整 `chunk_size` 参数
   - 检查网络延迟
