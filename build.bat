@echo off
rem Сборка akt.exe. Нужен установленный Python 3.10+ (python.org, галочка "Add to PATH")
pip install python-docx nuitka
python -m nuitka --onefile --assume-yes-for-downloads --windows-console-mode=disable --enable-plugin=tk-inter --windows-icon-from-ico=akt.ico --include-package-data=docx --include-data-files=template.docx=template.docx --include-data-files=akt.ico=akt.ico --output-dir=dist --output-filename=akt.exe app.py
echo Готово: dist\akt.exe
pause
