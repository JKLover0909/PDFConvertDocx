# PDF → DOCX Converter

Tool chuyển đổi file PDF sang DOCX, giao diện GUI đơn giản trên Windows.

## Cách dùng file `.exe` (dành cho người nhận)

1. Double-click `PDF_to_DOCX.exe`
2. Bấm vùng chọn file (hoặc kéo thả PDF vào)
3. Bấm **Convert**
4. File `.docx` sẽ được lưu cùng thư mục với file PDF gốc

> Không cần cài Python hay bất kỳ phần mềm nào.

---

## Cách build file `.exe` (dành cho developer)

### Yêu cầu
- Python 3.8+ đã cài trên máy
- Kết nối internet (để tải thư viện)

### Các bước

```bat
build_simple.bat
```

Script sẽ tự động:
1. Cài các thư viện cần thiết (`pdf2docx`, `tkinterdnd2`, `pyinstaller`)
2. Đóng gói thành `dist\PDF_to_DOCX.exe`

---

## Cấu trúc project

```
PDFConvertDocx/
├── converter.py        # Source code chính
├── requirements.txt    # Danh sách thư viện
├── build_simple.bat    # Script build .exe
└── README.md
```
