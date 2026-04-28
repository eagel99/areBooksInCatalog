@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt || goto :err
pyinstaller --onefile --noconsole --name BookCatalogChecker app.py || goto :err
echo.
echo Built: dist\BookCatalogChecker.exe
exit /b 0

:err
echo.
echo Build failed.
exit /b 1
