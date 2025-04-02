# PDF转Word工具

这是一个简单的PDF转Word转换工具，使用pdf2docx实现高质量的PDF到Word文档转换。

## 项目目标

- 提供简单易用的PDF到Word的转换服务
- 保持原PDF中的文本、图片和格式
- 提供良好的用户体验和转换状态反馈

## 技术架构

### 前端
- Vue3
- Element Plus
- Axios
- FileSaver.js

### 后端
- Python 3.10+
- FastAPI
- pdf2docx
- python-docx

## 功能特点

1. **文件上传**
   - 支持拖拽/点击上传（最大20MB）
   - 格式验证（仅支持.pdf文件）
   - 即时显示文件信息

2. **PDF转Word功能**
   - 文本提取与格式保留
   - 图片提取与插入
   - 表格结构保留
   - 文本框多段落支持
   - 转换进度实时显示

3. **安全与性能**
   - 文件类型验证
   - 异步任务处理
   - 自动清理临时文件
   - 独立日志管理

## 使用方法

1. 选择或拖拽PDF文件到上传区域
2. 点击"开始转换"按钮
3. 等待转换完成后，点击"下载文件"按钮获取转换后的Word文档
4. 如需进行新的转换，点击"刷新"按钮重置界面

## 安装与部署

### 环境要求
- Python 3.10+
- Node.js 16+
- Conda (推荐)

### 后端安装
```bash
# 使用pip安装
cd backend
pip install -r requirements.txt

# 或使用conda创建新环境（推荐）
conda create -n text2text python=3.12
conda activate text2text
cd backend
pip install -r requirements.txt
```

### 前端安装
```bash
cd frontend
npm install
```

### 启动服务
```bash
# 启动后端（使用conda环境）
conda activate pdf2word  # 如果使用conda环境
cd backend
python -m uvicorn main:app --reload

# 启动前端
cd frontend
npm run dev
```

## 代码结构

### 后端
- `main.py` - FastAPI应用主文件，处理HTTP请求
- `utils/pdf_converter.py` - PDF转Word转换核心功能
- `utils/file_manager.py` - 文件管理功能

### 前端
- `src/views/HomeView.vue` - 主页面组件
- `src/components/` - 各种UI组件

## 优化特点

1. **高质量转换**
   - 使用pdf2docx实现高质量转换
   - 优化文本框处理，防止文本框被错误识别为表格
   - 保持原始格式和布局

2. **简化依赖**
   - 只使用必要的依赖库
   - 减小部署体积
   - 提高安装和部署速度

3. **用户体验**
   - 实时进度反馈
   - 自动清理临时文件
   - 简洁直观的界面

## 许可证

本项目采用MIT许可证。详情请参阅LICENSE文件。



