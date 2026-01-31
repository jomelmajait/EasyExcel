import cv2
import numpy as np
import pytesseract
from PIL import Image
import pandas as pd
from django.shortcuts import render
from django.http import HttpResponse
import io
# Tandaan: Kung Windows gamit mo, baka kailangan mong i-point ang path ng tesseract.exe
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def upload_view(request):
    return render(request, 'upload.html')

def convert_image(request):
    if request.method == 'POST' and request.FILES.get('image'):
        image_file = request.FILES['image']
        
        # Basahin ang image gamit ang OpenCV para sa better table detection
        file_bytes = np.asarray(bytearray(image_file.read()), dtype=np.uint8)
        img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
        
        # Convert to grayscale at threshold para luminaw ang text at lines
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

        # OCR with Table-specific configuration (--psm 6)
        # Ito ay para pilitin ang OCR na hanapin ang table structure
        custom_config = r'--oem 3 --psm 6'
        text_data = pytesseract.image_to_string(thresh, config=custom_config)
        
        # Linisin ang data: Gamitin ang tabs o multiple spaces para mag-split ng columns
        rows = []
        for line in text_data.split('\n'):
            if line.strip():
                # Sinusubukan nating i-split base sa malalaking space sa pagitan ng words
                columns = [col.strip() for col in line.split('  ') if col.strip()]
                rows.append(columns)

        # Gawing DataFrame. Gagamit tayo ng empty columns para sa alignment.
        df = pd.DataFrame(rows)

        export_format = request.POST.get('format')
        if export_format == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False)
            output.seek(0)
            
            response = HttpResponse(output, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="better_converted_data.xlsx"'
            return response
            
    return render(request, 'upload.html')