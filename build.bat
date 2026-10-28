@echo off
rem Сборка akt.exe. Нужен установленный Python 3.10+ (python.org, галочка "Add to PATH")
pip install python-docx pyinstaller
pyinstaller --noconsole --onefile --name akt --add-data "template.docx;." app.py
copy materials.json dist\materials.json
echo Готово: папка dist (akt.exe + materials.json)
pause
