@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt || goto :err
python -m PyInstaller --onefile --noconsole --name BookCatalogChecker ^
    --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes ^
    app.py || goto :err
echo.
echo Built: dist\BookCatalogChecker.exe
exit /b 0

:err
echo.
echo Build failed.
exit /b 1
