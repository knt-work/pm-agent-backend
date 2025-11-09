import json
import boto3
import os
import io

# ----- CÁC THƯ VIỆN CẦN CÀI BẰNG 'pip' -----
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import openpyxl # Để đọc file Excel nhúng từ chart
# --------------------------------------

# Khai báo S3 client
s3_client = boto3.client('s3')

# --- HÀM 1: Đọc Chart Data ---
def extract_chart_data(chart):
    try:
        chart_part = chart.part
        embedded_excel_blob = chart_part.chart_workbook.xlsx_part.blob
        workbook = openpyxl.load_workbook(io.BytesIO(embedded_excel_blob), data_only=True)
        sheet = workbook.active 
        data = []
        for row in sheet.iter_rows():
            row_data = [cell.value for cell in row]
            data.append(row_data)

        return {
            "title": chart.chart_title.text_frame.text if chart.has_title else "Untitled Chart",
            "excel_data": data
        }
    except Exception as e:
        print(f"Error reading chart data: {e}")
        try:
            return {
                "title": chart.chart_title.text_frame.text if chart.has_title else "Untitled Chart",
                "series": [s.name for s in chart.series]
            }
        except Exception:
            return None

# --- HÀM 2: Đọc Text ---
def extract_shape_text(shape):
    text_runs = []
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            text_runs.extend(extract_shape_text(s))
    elif shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_runs.append(run.text)
    elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        for row in shape.table.rows:
            for cell in row.cells:
                text_runs.append(cell.text_frame.text)
    return text_runs

# --- HÀM XỬ LÝ CHÍNH (CHO LAMBDA) ---
def lambda_handler(event, context):
    BUCKET_NAME = os.environ.get('BUCKET_NAME', 'default-bucket')

    # ----- KHỐI 'TRY' BÊN NGOÀI -----
    try:
        body = json.loads(event.get('body', '{}'))
        file_keys = body.get('fileKeys')

        if not file_keys:
            return {'statusCode': 400, 'body': json.dumps({'message': "'fileKeys' (array) is required."})}

        full_analysis = {}
        processed_files = []
        failed_files = []

        for file_key in file_keys:
            if not file_key.endswith(('.pptx')):
                print(f"Skipping non-pptx file: {file_key}")
                continue

            # ----- KHỐI 'TRY' BÊN TRONG -----
            try:
                response = s3_client.get_object(Bucket=BUCKET_NAME, Key=file_key)
                file_stream = io.BytesIO(response['Body'].read())
                prs = Presentation(file_stream)
                file_results = {
                    "file_name": file_key,
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

                full_analysis[file_key] = file_results
                processed_files.append(file_key)

            # ----- KHỐI 'EXCEPT' BỊ THIẾU -----
            except Exception as e:
                print(f"Error processing file {file_key}: {e}")
                failed_files.append(file_key)

        # Tóm tắt
        summary_text = f"Processed {len(processed_files)} files. "
        for key, value in full_analysis.items():
            total_charts = sum(len(s['charts']) for s in value['slides'])
            total_text_snippets = sum(len(s['text']) for s in value['slides'])
            summary_text += f"File '{key}' ({value['slide_count']} slides) contains {total_text_snippets} text snippets and {total_charts} charts. "

        return {
            'statusCode': 200,
            'headers': {'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*'},
            'body': json.dumps({ 
                'message': 'Analysis complete.',
                'summary': summary_text,
                'full_analysis_snippet': str(full_analysis)[:1000] + "...",
                'processed': processed_files,
                'failed': failed_files
            })
        }

    # ----- KHỐI 'EXCEPT' CỦA BÊN NGOÀI -----
    except Exception as e:
        print(f"Handler error: {e}")
        return {'statusCode': 500, 'body': json.dumps({'message': 'Error processing files'})}