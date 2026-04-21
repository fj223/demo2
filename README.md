# Инструмент для преобразования HTML-слайдов
## Функциональность
Преобразование файлов HTML-слайдов в форматы PDF и PPTX.

## Настройка окружения

### 1. Создание виртуального окружения
```powershell
# Windows
python -m venv venv

# Активация виртуального окружения
.\venv\Scripts\Activate.ps1
```

### 2. Установка зависимостей
```powershell
pip install -r requirements.txt
playwright install chromium
```

### 3. Запуск проекта
```powershell
# Обработка всех HTML-файлов в каталоге
python convert_slides.py 01042026\01042026

# Обработка отдельного файла
python convert_slides.py 01042026\01042026\ch1.html

# 生成可编辑 PPTX 而不是图片 PPTX
python convert_slides.py --editable 01042026\01042026\ch1.html

# 从在线网页转换成可编辑 PPTX
python convert_slides.py --editable https://example.com/your-slide-page
```

> 可编辑 PPTX 现在更好地保留原始布局、标题、文本样式和高亮格式，适合在 PowerPoint 中继续编辑。
## Вывод
Преобразованные файлы PDF и PPTX будут сохранены в каталоге `output`.

## Добавление текста на слайды
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
