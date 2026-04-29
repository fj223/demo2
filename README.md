# Инструмент для преобразования HTML-слайдов
## Функциональность
Преобразование файлов HTML-слайдов в форматы PDF и PPTX. Поддерживает как图片模式 (将整个页面渲染为图片)，也支持**可编辑模式** (提取文本内容，生成可编辑的PPTX)。

## 快速开始 - 从零开始安装

### 前置需求
你需要先安装 **Python 3.8 或更高版本**。

检查是否已安装：
```bash
python --version
```

### 1. 克隆或下载项目
```bash
# 如果使用git克隆
git clone https://github.com/fj223/demo2.git
cd demo2
```

### 2. 创建虚拟环境 (推荐)
```powershell
# Windows
python -m venv venv

# 激活虚拟环境
.\venv\Scripts\Activate.ps1
```

如果在 PowerShell 上遇到执行策略限制，请先运行：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. 安装依赖
```powershell
# 安装Python库
pip install -r requirements.txt

# 安装 Playwright 浏览器驱动 (Chromium)
playwright install chromium
```

### 4. 运行项目
```powershell
# 处理所有HTML文件在目录
python convert_slides.py 01042026\01042026

# 处理单个文件
python convert_slides.py 01042026\01042026\ch1.html

# 生成可编辑的 PPTX (推荐)
python convert_slides.py --editable 01042026\01042026\GPU.html
```

## 详细使用说明

### 命令行参数
```powershell
python convert_slides.py [选项] <文件或目录>
```

#### 主要选项
| 选项 | 说明 | 示例 |
|------|------|------|
| `--editable` | 生成可编辑的 PPTX 而不是图片模式 | `python convert_slides.py --editable GPU.html` |
| `--no-pdf` | 跳过 PDF 生成 | `python convert_slides.py --no-pdf ch1.html` |
| `--text "文本"` | 为所有幻灯片添加文本 | `python convert_slides.py --text "备注" GPU.html` |
| `--text-file 文件.txt` | 从文本文件读取内容添加到幻灯片 | `python convert_slides.py --text-file notes.txt GPU.html` |
| `--clipboard` | 从剪贴板读取内容添加到幻灯片 | `python convert_slides.py --clipboard GPU.html` |
| `--position top/bottom` | 设置添加文本的位置 | `python convert_slides.py --text "备注" --position top GPU.html` |

### 使用示例

#### 1. 可编辑 PPTX 模式 (推荐)
这是最有用的模式，会保留文本格式、布局、标题样式等，可以在 PowerPoint 中继续编辑：

```powershell
# 转换 GPU.html 为可编辑 PPTX
python convert_slides.py --editable 01042026\01042026\GPU.html
```

#### 2. 图片模式 + PDF
将 HTML 渲染为图片，然后生成 PDF 和图片 PPTX：

```powershell
# 单个文件
python convert_slides.py 01042026\01042026\ch1.html

# 整个目录
python convert_slides.py 01042026\01042026
```

#### 3. 只生成可编辑 PPTX (跳过 PDF)
```powershell
python convert_slides.py --editable --no-pdf 01042026\01042026\GPU.html
```

## 项目结构说明
```
demo2/
├── convert_slides.py        # 主转换脚本
├── requirements.txt         # Python依赖列表
├── README.md               # 说明文档
├── 01042026/              # HTML幻灯片目录
│   └── 01042026/
│       ├── GPU.html       # Q&A风格演示文稿
│       ├── ch1.html       # 第1章
│       ├── ch2.html       # 第2章
│       ├── Logic_programming.html
│       └── ...
└── output/                 # 输出目录（自动创建）
    ├── GPU.pptx
    ├── ch1.pptx
    └── ...
```

## Вывод
Преобразованные файлы PDF и PPTX будут сохранены в каталоге `output`.

## 常见问题

### Q: 在 Windows 上激活虚拟环境失败
A: 运行 `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`，然后重试。

### Q: playwright install 失败
A: 确保有稳定的网络连接，或者尝试使用代理。第一次运行时 Playwright 会下载 Chromium 浏览器 (~200MB)。

### Q: 输出的 PPTX 只有标题，没有内容
A: 使用 `--editable` 参数，该模式会强制展开所有隐藏的内容块。我们已经修复了 Q&A 风格的 HTML (如 GPU.html) 的内容提取问题。

### Q: 可以转换在线网页吗？
A: 可以！直接传入网址：
```powershell
python convert_slides.py --editable https://example.com/your-slide-page
```

## 添加文本到幻灯片
Вы можете добавить текст на каждую страницу при генерации:

```powershell
python convert_slides.py --text "这是要添加到幻灯片的文本" 01042026\01042026\ch1.html
python convert_slides.py --text-file notes.txt 01042026\01042026\ch1.html
python convert_slides.py --clipboard 01042026\01042026\ch1.html
```

Дополнительно можно задать позицию текста:

```powershell
python convert_slides.py --text "文本" --position top 01042026\01042026\ch1.html
```

## 技术说明

本项目使用以下技术:
- **Playwright**: 自动化浏览器，加载和渲染 HTML
- **Pillow**: 图像处理
- **python-pptx**: PowerPoint 文件生成

## 更新日志

### 最新版本
- ✅ 修复 Q&A 风格 HTML 的内容提取问题
- ✅ 添加 `force-open` 类来强制展开隐藏内容
- ✅ 支持 55+ 张幻灯片的完整内容提取
