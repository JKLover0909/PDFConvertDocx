# PDF → DOCX Converter

Tool chuyển đổi file PDF sang DOCX, giao diện GUI đơn giản trên Windows.

## ⬇️ Tải về dùng ngay (dành cho người dùng Windows)

1. Vào trang **[Releases](../../releases/latest)**
2. Tải file `PDF_to_DOCX.exe`
3. Double-click để chạy — **không cần cài Python hay phần mềm nào**

## Cách dùng

1. Bấm vùng chọn file (hoặc kéo thả PDF vào)
2. Bấm **Convert**
3. File `.docx` sẽ được lưu cùng thư mục với file PDF gốc

---

## 🔨 Build tự động (GitHub Actions)

Mỗi khi push code lên `main`, GitHub Actions sẽ tự build `.exe` trên Windows runner.

**Để tạo Release mới:**
```bash
git tag v1.0.0
git push origin v1.0.0
```
File `PDF_to_DOCX.exe` sẽ tự động đính kèm vào trang Releases.

---

## Cách build thủ công trên Windows (tuỳ chọn)

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
├── .github/workflows/build.yml  # CI: tự build .exe trên GitHub Actions
├── converter.py                 # Source code chính
├── requirements.txt             # Danh sách thư viện
├── build_simple.bat             # Script build thủ công trên Windows
└── README.md
```
