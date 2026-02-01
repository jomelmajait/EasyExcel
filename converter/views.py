import cv2
import numpy as np
import pytesseract
import pandas as pd
from django.shortcuts import render
from django.http import HttpResponse
from io import BytesIO

# Make sure this path is correct for your machine
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def upload_view(request):
    return render(request, "upload.html")

def convert_image(request):
    if request.method == "POST" and request.FILES.get('image'):
        image_file = request.FILES['image']

        # 1. Read image from memory
        file_bytes = np.frombuffer(image_file.read(), np.uint8)
        img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)

        if img is None:
            return HttpResponse("Invalid image format", status=400)

        # 2. Preprocessing for better OCR accuracy
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # Thresholding helps remove the gray grid lines and makes text pop
        gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

        # 3. Get OCR data with Page Segmentation Mode 6 (treat as uniform block)
        custom_config = r'--oem 3 --psm 6'
        data = pytesseract.image_to_data(gray, config=custom_config, output_type=pytesseract.Output.DATAFRAME)

        # Filter out empty text and noise
        data = data[data['text'].notnull() & (data['text'].str.strip() != '')].copy()

        if data.empty:
            return HttpResponse("No text detected in the image.", status=400)

        # 4. Logic to group words into rows
        # Sort by vertical 'top' position first
        data = data.sort_values(['top', 'left'])
        
        rows = []
        current_row_data = []
        
        last_top = data.iloc[0]['top']
        tolerance = 15  # Pixels. Adjust if rows are very close together.

        for _, row in data.iterrows():
            # Check if current word is on the same line as the previous word
            if abs(row['top'] - last_top) <= tolerance:
                current_row_data.append(row)
            else:
                # Row is finished. Sort by 'left' to keep columns in order.
                current_row_data.sort(key=lambda x: x['left'])
                rows.append([str(r['text']) for r in current_row_data])
                
                # Start new row
                current_row_data = [row]
                last_top = row['top']
        
        # Add the final row
        if current_row_data:
            current_row_data.sort(key=lambda x: x['left'])
            rows.append([str(r['text']) for r in current_row_data])

        # 5. Convert to DataFrame
        df = pd.DataFrame(rows)

        # 6. Save to Excel
        output = BytesIO()
        try:
            # We use xlsxwriter as the engine (requires: pip install xlsxwriter)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False)
        except Exception as e:
            return HttpResponse(f"Error generating Excel: {e}", status=500)
            
        output.seek(0)

        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="converted.xlsx"'
        return response

    return render(request, 'upload.html')