import json
import io
import os.path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ----- IMPORT CÁC HÀM TỪ CODE LAMBDA CỦA BẠN -----
try:
    from lambda_function import extract_chart_data, extract_shape_text
except ImportError:
    print("LỖI: Không tìm thấy file 'lambda_function.py'.")
    print("Hãy đảm bảo 'test_local_reader.py' và 'lambda_function.py' ở chung thư mục.")
    exit()

# -----------------------------------------------------------------
# 📍 HÃY THAY ĐỔI ĐƯỜNG DẪN NÀY
# -----------------------------------------------------------------
LOCAL_FILE_PATH = "mock-proposal.pptx" # (Hoặc tên file pptx của bạn)
# -----------------------------------------------------------------


def analyze_local_pptx(file_path):
    if not os.path.exists(file_path):
        print(f"LỖI: Không tìm thấy file tại đường dẫn: {file_path}")
        print("Hãy kiểm tra lại biến 'LOCAL_FILE_PATH'.")
        return None

    print(f"--- Bắt đầu phân tích file local: {file_path} ---")

    try:
        prs = Presentation(file_path)
    except Exception as e:
        print(f"LỖI: Không thể mở file. File có thể bị hỏng hoặc không phải PPTX.")
        print(f"Chi tiết lỗi: {e}")
        return None

    file_results = {
        "file_name": os.path.basename(file_path),
        "slide_count": len(prs.slides),
        "slides": []
    }

    for i, slide in enumerate(prs.slides):
        slide_data = {
            "slide_number": i + 1,
            "text": [],
            "charts": [],
            "image_count": 0
        }
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.CHART:
                chart_data = extract_chart_data(shape.chart)
                if chart_data:
                    slide_data["charts"].append(chart_data)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                slide_data["image_count"] += 1
            else:
                extracted_texts = extract_shape_text(shape)
                if extracted_texts:
                    slide_data["text"].extend(extracted_texts)

        slide_data["text"] = list(filter(None, [t.strip() for t in slide_data["text"]]))
        file_results["slides"].append(slide_data)

    print("--- Phân tích file local THÀNH CÔNG ---")
    return file_results

# ----- PHẦN CHẠY CHÍNH (Rất quan trọng) -----
if __name__ == "__main__":

    full_analysis = analyze_local_pptx(LOCAL_FILE_PATH)

    if full_analysis:
        print("\n\n--- KẾT QUẢ PHÂN TÍCH (JSON) ---")
        try:
            print(json.dumps(full_analysis, indent=2, ensure_ascii=False))
        except Exception as e:
            print(f"Lỗi khi in JSON (có thể do encoding): {e}")
            print(full_analysis)
    else:
        print("\nKhông có kết quả để hiển thị do lỗi ở trên.")