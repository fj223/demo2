# HTML Slides 转换工具

## 功能
将 HTML 幻灯片文件转换为 PDF 和 PPTX 格式。

## 环境设置

### 1. 创建虚拟环境
```powershell
# Windows
python -m venv venv

# 激活虚拟环境
.\venv\Scripts\Activate.ps1
```

### 2. 安装依赖
```powershell
pip install -r requirements.txt
playwright install chromium
```

### 3. 运行项目
```powershell
# 处理目录中的所有 HTML 文件
python convert_slides.py 01042026\01042026

# 处理单个文件
python convert_slides.py 01042026\01042026\ch1.html
```

## 输出
转换后的 PDF 和 PPTX 文件会保存在 `output` 目录中。