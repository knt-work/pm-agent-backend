# Hướng dẫn chạy `test_local_reader.py` ở local

## 1) Yêu cầu
- Python 3.10+ (đã kiểm tra với 3.13.5)
- Hai file `lambda_function.py` và `test_local_reader.py` nằm cùng thư mục
- Một file `.pptx` để phân tích (ví dụ: `mock-proposal.pptx`)

## 2) Cài môi trường

### Windows (PowerShell)
- `py -m venv .venv`
- `.\.venv\Scripts\Activate.ps1`
- Nếu bị chặn, chạy: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`
- `python -m pip install --upgrade pip`
- `pip install python-pptx openpyxl boto3`

### macOS/Linux
- `python3 -m venv .venv`
- `source .venv/bin/activate`
- `python -m pip install --upgrade pip`
- `pip install python-pptx openpyxl boto3`

## 3) Chỉ định file PPTX cần đọc
- Mở `test_local_reader.py` và cập nhật biến `LOCAL_FILE_PATH` (gần đầu file), ví dụ:
  - `LOCAL_FILE_PATH = "your-presentation.pptx"`
- Đảm bảo file `.pptx` nằm cùng thư mục với script, hoặc dùng đường dẫn đầy đủ.
- Kiểm tra tồn tại file:
  - PowerShell: `Test-Path .\your-presentation.pptx`
  - macOS/Linux: `[ -f your-presentation.pptx ] && echo OK || echo MISSING`

## 4) Chạy script
- Windows: `python .\test_local_reader.py`
- macOS/Linux: `python3 ./test_local_reader.py`

Kết quả sẽ in ra JSON tóm tắt (số slide, text, charts, hình ảnh) theo từng slide.

## 5) Lỗi thường gặp & cách xử lý
- `ModuleNotFoundError: No module named 'pptx'` (hoặc `openpyxl`, `boto3`):
  - Chưa cài đủ thư viện. Chạy: `pip install python-pptx openpyxl boto3`.
- `LỖI: Không tìm thấy file tại đường dẫn ...`:
  - Kiểm tra lại `LOCAL_FILE_PATH` và vị trí file `.pptx`.
- `Lỗi khi in JSON (có thể do encoding)`:
  - Script đã có nhánh in ra đối tượng thô; có thể bỏ qua hoặc lưu output bằng redirect: `python test_local_reader.py > output.json`.
- `Error reading chart data:`
  - Một số chart có thể không có Excel nhúng; script đã fallback trả về `title` và `series` (nếu có).

## 6) Gợi ý thêm
- Nên luôn chạy trong virtualenv (`.venv`) để cô lập phụ thuộc.
- Nếu muốn tái sử dụng về sau, có thể lưu phụ thuộc:
  - `pip freeze > requirements.txt`
  - Cài lại bằng: `pip install -r requirements.txt`

