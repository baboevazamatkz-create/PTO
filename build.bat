@echo off
rem Сборка akt.exe. Нужен установленный Python 3.10+ (python.org, галочка "Add to PATH")
pip install python-docx pyinstaller
pyinstaller --noconsole --onefile --name akt --add-data "template.docx;." app.py
echo Готово: папка dist (akt.exe)
pause
