import os
import io
import numpy as np
import fitz # PyMuPDF
from PIL import Image
import threading
import shutil

# Global reader cache as requested
ocr_reader = None
import_error_occurred = False

class OCREngine:
    _tesseract_path = None

    @staticmethod
    def get_ocr_reader():
        global ocr_reader, import_error_occurred
        if import_error_occurred:
            return None
            
        if ocr_reader is None:
            try:
                import easyocr
                # Load reader once and store in models folder
                ocr_reader = easyocr.Reader(
                    ['en'], 
                    gpu=False, 
                    verbose=False,
                    model_storage_directory='./ocr_models'
                )
            except (ImportError, RuntimeError, ValueError, Exception):
                import_error_occurred = True
                return None
        return ocr_reader

    @classmethod
    def find_tesseract(cls):
        """Deeply searches for tesseract.exe, prioritizing local bundled versions."""
        if cls._tesseract_path and os.path.exists(cls._tesseract_path):
            return cls._tesseract_path

        import pytesseract
        import sys

        # Candidates folder search list
        # 1. Check if we're running in a PyInstaller bundle (_MEIPASS)
        # 2. Check the local app directory (for portable usage)
        # 3. Check common system installation paths
        
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        local_app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        candidates = [
            # Bundle / Local Folder (This is the "Portable" solution)
            os.path.join(base_dir, "tesseract", "tesseract.exe"),
            os.path.join(local_app_dir, "tesseract", "tesseract.exe"),
            os.path.join(os.getcwd(), "tesseract", "tesseract.exe"),
            
            # Common system installation paths
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
            os.path.expandvars(r"%APPDATA%\Tesseract-OCR\tesseract.exe"),
        ]
        
        # Also check system PATH using shutil
        path_found = shutil.which("tesseract")
        if path_found:
            candidates.insert(0, path_found)

        for path in candidates:
            if os.path.exists(path):
                cls._tesseract_path = path
                pytesseract.pytesseract.tesseract_cmd = path
                return path
        return None

    @staticmethod
    def _raise_ocr_error():
        msg = (
            "OCR Engine Error: Tesseract not found.\n\n"
            "HOW TO FIX (ZERO INSTALLATION):\n"
            "1. Create a folder named 'tesseract' in your app directory.\n"
            "2. Copy all Tesseract files into it.\n"
            "3. Restart the app. It will now work without any setup!"
        )
        raise RuntimeError(msg)

    @staticmethod
    def extract_text_from_image(image_path: str) -> str:
        """Extracts text from a single image trying EasyOCR then Tesseract."""
        reader = OCREngine.get_ocr_reader()
        if reader:
            try:
                results = reader.readtext(image_path, detail=0, paragraph=True)
                return '\n'.join(results)
            except Exception:
                pass # Fallback to Tesseract
        
        # Tesseract fallback
        if not OCREngine.find_tesseract():
            OCREngine._raise_ocr_error()
        
        import pytesseract
        img = Image.open(image_path)
        return pytesseract.image_to_string(img)

    @staticmethod
    def extract_text_from_pdf(pdf_path: str, progress_callback=None):
        """Standard text extraction from scanned PDF by converting pages to images."""
        import fitz
        import io
        import tempfile
        
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        all_text = []

        for i in range(total_pages):
            page = doc.load_page(i)
            # Render page at 300 DPI for high OCR accuracy
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
            
            # Create a temporary file for the page image
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp.write(pix.tobytes("png"))
                tmp_path = tmp.name
                
            try:
                # Use the robust image extract method (EasyOCR -> Tesseract)
                page_text = OCREngine.extract_text_from_image(tmp_path)
                if page_text.strip():
                    all_text.append(f"--- Page {i+1} ---\n{page_text}")
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            
            if progress_callback:
                progress_callback(int(((i + 1) / total_pages) * 100))
                
        doc.close()
        return "\n\n".join(all_text)

    @staticmethod
    def get_text_with_bboxes(image_source, page_width=None, page_height=None) -> list:
        """
        Extracts text with bounding boxes from an image or PIL Image object.
        Returns: list of {"text": str, "bbox": [x0, y0, x1, y1]}
        """
        results = []
        reader = OCREngine.get_ocr_reader()
        
        # Convert source to numpy array for EasyOCR
        if isinstance(image_source, Image.Image):
            img_array = np.array(image_source)
        elif isinstance(image_source, (str, bytes)):
            img = Image.open(io.BytesIO(image_source) if isinstance(image_source, bytes) else image_source)
            img_array = np.array(img)
        else:
            return []

        if reader:
            try:
                # detail=1 returns list of ([[x,y], [x,y], [x,y], [x,y]], text, prob)
                ocr_data = reader.readtext(img_array, detail=1, paragraph=False)
                for (coords, text, prob) in ocr_data:
                    # coords: [top_left, top_right, bottom_right, bottom_left]
                    x_coords = [p[0] for p in coords]
                    y_coords = [p[1] for p in coords]
                    bbox = [min(x_coords), min(y_coords), max(x_coords), max(y_coords)]
                    results.append({"text": text, "bbox": bbox})
                return results
            except Exception:
                pass # Fallback to Tesseract

        # Tesseract fallback
        if not OCREngine.find_tesseract():
            return [] # Silent fail or log

        import pytesseract
        # image_to_data returns tsv data with bboxes
        data = pytesseract.image_to_data(image_source, output_type=pytesseract.Output.DICT)
        for i in range(len(data['text'])):
            if int(data['conf'][i]) > 0 and data['text'][i].strip():
                x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                results.append({
                    "text": data['text'][i],
                    "bbox": [float(x), float(y), float(x+w), float(y+h)]
                })
        return results

