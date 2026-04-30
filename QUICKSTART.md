# 🚀 快速开始指南

## 问题描述

你有一个 HTML 演示文稿（GPU.html），其中包含大量隐藏内容，需要点击按钮才能展开。你想将它转换为**可编辑文字**的 PPTX 文件。

## 解决方案

使用 `convert_slides_text.py` - 它会自动点击所有按钮，提取完整内容。

## 3 步完成转换

### 步骤 1: 安装依赖（首次使用）

```powershell
# 安装 Python 库
pip install -r requirements.txt

# 安装浏览器驱动
playwright install chromium
```

⏱️ 预计时间：2-3 分钟

### 步骤 2: 运行转换

```powershell
# 转换 GPU.html
python convert_slides_text.py 01042026\01042026\GPU.html
```

⏱️ 预计时间：30-60 秒

### 步骤 3: 查看结果

打开 `output/GPU_editable.pptx`，你会看到：
- ✅ 5 个演示标题页
- ✅ 25 个问题内容页
- ✅ 每个问题包含：理论、任务、解决方案
- ✅ 所有文字都可以编辑

## 🎯 你会得到什么

### 原始 HTML（GPU.html）
```
演示 1: 标题
  问题 1: [点击展开] 📖 理论
         [点击展开] 📝 任务
         [点击展开] ✅ 解决方案
  问题 2: ...
  ...
```

### 生成的 PPTX（GPU_editable.pptx）
```
幻灯片 1: 演示 1 标题页
幻灯片 2: 问题 1
  📖 理论: [完整文本内容]
  📝 任务: [完整文本内容]
  ✅ 解决方案: [完整文本内容]
幻灯片 3: 问题 2
  ...
```

## 📊 对比

| 特性 | 旧方法 (截图) | 新方法 (文本提取) |
|------|--------------|------------------|
| 可编辑 | ❌ | ✅ |
| 隐藏内容 | ❌ 只有导航栏 | ✅ 完整内容 |
| 文件大小 | 50MB | 500KB |
| 提取内容 | ~500 字符 | ~50,000 字符 |

## 🔧 故障排除

### 问题：pip install 失败
```powershell
# 升级 pip
python -m pip install --upgrade pip

# 重试
pip install -r requirements.txt
```

### 问题：playwright install 失败
```powershell
# 使用国内镜像（如果在中国）
set PLAYWRIGHT_DOWNLOAD_HOST=https://npmmirror.com/mirrors/playwright/
playwright install chromium
```

### 问题：找不到文件
```powershell
# 检查当前目录
dir 01042026\01042026\GPU.html

# 如果不存在，调整路径
python convert_slides_text.py <你的实际路径>\GPU.html
```

## 🎓 进阶用法

### 批量转换整个目录
```powershell
python convert_slides_text.py 01042026\01042026
```

### 查看提取统计
```powershell
python compare_methods.py
```

### 测试功能
```powershell
python test_extraction.py
```

## 📞 需要帮助？

1. 查看详细文档：`README_TEXT_EXTRACTION.md`
2. 查看解决方案总结：`SOLUTION_SUMMARY.md`
3. 运行对比演示：`python compare_methods.py`

## ✨ 核心优势

1. **自动化**：无需手动点击按钮
2. **完整性**：提取所有隐藏内容
3. **可编辑**：生成真实文本，不是图片
4. **高效**：处理速度快，文件小

---

**就这么简单！** 🎉

现在你可以：
- ✏️ 编辑 PPTX 中的任何文字
- 🌍 翻译成其他语言
- 📝 添加笔记和注释
- 📤 轻松分享（文件小）
