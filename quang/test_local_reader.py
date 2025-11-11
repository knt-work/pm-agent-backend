# test_local_reader.py
# -*- coding: utf-8 -*-

import json
import os
import io
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ----- IMPORT TỪ LAMBDA FILE -----
try:
    from lambda_function import (
        extract_chart_data,
        extract_text_shape,
        extract_table_shape,
        extract_picture_shape,
        _analyze_presentation_stream
    )
except ImportError:
    print("LỖI: Không tìm thấy 'lambda_function.py' hoặc thiếu các hàm cần thiết.")
    exit(1)

# -----------------------------------------------------------------
# 📍 ĐƯỜNG DẪN
# -----------------------------------------------------------------
LOCAL_FILE_PATH = os.path.join(os.path.dirname(__file__), "mock-proposal.pptx")
OUTPUT_FILE_PATH = os.path.join(os.path.dirname(__file__), "analysis_output.json")
# -----------------------------------------------------------------


def analyze_local_pptx(file_path):
    if not os.path.exists(file_path):
        print(f"LỖI: Không tìm thấy file tại: {file_path}")
        return None

    print(f"--- Bắt đầu phân tích file local: {file_path} ---")

    try:
        with open(file_path, "rb") as f:
            file_stream = io.BytesIO(f.read())

        result = _analyze_presentation_stream(file_stream, os.path.basename(file_path))
        print("--- ✅ Phân tích file local THÀNH CÔNG ---")
        return result
    except Exception as e:
        print("❌ LỖI: Không thể mở/đọc file PPTX.")
        print(f"Chi tiết lỗi: {e}")
        return None


def save_to_json_file(data, path):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"--- 💾 Đã lưu kết quả phân tích ra file: {path} ---")
    except Exception as e:
        print(f"❌ Lỗi khi ghi file JSON: {e}")


# ----- RUN -----
if __name__ == "__main__":
    results = analyze_local_pptx(LOCAL_FILE_PATH)

    if results:
        save_to_json_file(results, OUTPUT_FILE_PATH)

        print("\n\n--- 🔍 XEM TRƯỚC (TÓM TẮT) ---")
        try:
            short_preview = json.dumps(results, indent=2, ensure_ascii=False)[:1000]
            print(short_preview + "\n...\n(đã cắt bớt, xem full trong analysis_output.json)")
        except Exception as e:
            print(f"Lỗi khi in JSON: {e}")
            print(str(results)[:1000] + "...")
    else:
        print("\nKhông có kết quả để hiển thị do lỗi ở trên.")
