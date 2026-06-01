@echo off
echo ========================================
echo  PDF to DOCX Converter - Build Script
echo ========================================
echo.

:: Install dependencies
echo [1/3] Cai dat thu vien...
pip install -r requirements.txt

echo.
echo [2/3] Dong goi thanh file .exe...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "PDF_to_DOCX" ^
    --add-data "%VIRTUAL_ENV%\Lib\site-packages\tkinterdnd2;tkinterdnd2/" ^
    converter.py

echo.
echo [3/3] Hoan thanh!
echo File .exe nam tai: dist\PDF_to_DOCX.exe
echo.
pause
