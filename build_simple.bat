@echo off
echo ========================================
echo  PDF to DOCX Converter - Build Script
echo ========================================
echo.

:: Install dependencies
echo [1/3] Cai dat thu vien...
pip install pdf2docx tkinterdnd2 pyinstaller

echo.
echo [2/3] Dong goi thanh file .exe...

:: Try build with tkinterdnd2 drag-and-drop support
python -c "import tkinterdnd2; import os; p=os.path.dirname(tkinterdnd2.__file__); print(p)" > tmp_path.txt 2>nul
set /p DND_PATH=<tmp_path.txt
del tmp_path.txt 2>nul

if exist "%DND_PATH%\tkdnd" (
    echo   -> Build voi ho tro keo tha file...
    pyinstaller --onefile --windowed --name "PDF_to_DOCX" --add-data "%DND_PATH%;tkinterdnd2" converter.py
) else (
    echo   -> Build khong co keo tha (van dung duoc nut chon file)...
    pyinstaller --onefile --windowed --name "PDF_to_DOCX" converter.py
)

echo.
if exist "dist\PDF_to_DOCX.exe" (
    echo [3/3] THANH CONG!
    echo File .exe nam tai: dist\PDF_to_DOCX.exe
    echo Co the copy file nay chia se cho nguoi khac dung.
) else (
    echo [3/3] Co loi xay ra, kiem tra output o tren.
)
echo.
pause
