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
```

## Вывод
Преобразованные файлы PDF и PPTX будут сохранены в каталоге `output`.