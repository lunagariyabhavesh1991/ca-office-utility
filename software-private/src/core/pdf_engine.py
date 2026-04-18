import os
from io import BytesIO
from typing import List, Tuple, Dict, Optional
from pypdf import PdfReader, PdfWriter

class PDFEngine:
    @staticmethod
    def check_ghostscript() -> bool:
        """Checks if Ghostscript is available on the system."""
        import shutil
        for cmd in ['gswin64c', 'gswin32c', 'gs']:
            if shutil.which(cmd):
                return True
        return False

    @staticmethod
    def merge_pdfs(pdf_paths: List[str], output_path: str) -> None:
        """Merges multiple PDFs into a single file."""
        writer = PdfWriter()
        for path in pdf_paths:
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
        
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

    @staticmethod
    def split_pdf(pdf_path: str, output_folder: str, page_ranges: Optional[List[Tuple[int, int]]] = None) -> List[str]:
        """
        Splits a PDF by page ranges or into individual pages if ranges not provided.
        Returns a list of generated file paths.
        """
        reader = PdfReader(pdf_path)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_files = []

        if not page_ranges:
            # Split into individual pages
            for i in range(len(reader.pages)):
                writer = PdfWriter()
                writer.add_page(reader.pages[i])
                
                out_path = os.path.join(output_folder, f"{base_name}_page_{i+1}.pdf")
                with open(out_path, "wb") as out_file:
                    writer.write(out_file)
                output_files.append(out_path)
        else:
            # Split by ranges (1-indexed for logic, 0-indexed for pypdf)
            for i, (start, end) in enumerate(page_ranges):
                writer = PdfWriter()
                # Ensure valid range
                start_idx = max(0, start - 1)
                end_idx = min(len(reader.pages), end)
                
                for page_num in range(start_idx, end_idx):
                    writer.add_page(reader.pages[page_num])
                    
                out_path = os.path.join(output_folder, f"{base_name}_part_{i+1}.pdf")
                with open(out_path, "wb") as out_file:
                    writer.write(out_file)
                output_files.append(out_path)
                
        return output_files

    @staticmethod
    def extract_pages(pdf_path: str, output_path: str, pages_to_keep: List[int]) -> None:
        """Extracts specific pages (1-indexed) and saves to a new PDF."""
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        for page_num in pages_to_keep:
            idx = page_num - 1
            if 0 <= idx < len(reader.pages):
                writer.add_page(reader.pages[idx])
                
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

    @staticmethod
    def rotate_pages(pdf_path: str, output_path: str, rotations: Dict[int, int]) -> None:
        """
        Rotates specific pages. 
        rotations: Dictionary of {page_number (1-indexed): angle (90, 180, 270)}
        """
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        for i, page in enumerate(reader.pages):
            page_num = i + 1
            if page_num in rotations:
                page.rotate(rotations[page_num])
            writer.add_page(page)
            
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

    @staticmethod
    def remove_pages(pdf_path: str, output_path: str, pages_to_remove: List[int]) -> None:
        """
        Removes specific pages (1-indexed) and saves to a new PDF.
        """
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        remove_idx = set(p - 1 for p in pages_to_remove)
        
        for i, page in enumerate(reader.pages):
            if i not in remove_idx:
                writer.add_page(page)
                
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

    @staticmethod
    def compress_pdf_target(pdf_path: str, output_path: str, mode: str = "default", 
                            target_kb: Optional[int] = None, progress_callback=None,
                            current_file_idx: int = 0, total_files: int = 1) -> bool:
        """
        Compress PDF by recompressing all embedded images inside the PDF.
        This is the ONLY way to significantly reduce PDF file size.
        """
        import fitz  # PyMuPDF
        from PIL import Image
        import io
        import os

        def get_quality_for_mode(mode):
            if mode == "default":
                return 45   # aggressive compression, ~40-70% size reduction
            elif mode == "100":
                return 20
            elif mode == "200":
                return 25
            elif mode == "500":
                return 35
            elif mode in ("1024", "1mb"):
                return 45
            else:
                return 45

        def try_compress_at_quality(input_path, out_path, quality, current_file_idx=0, total_files=1, 
                                    iteration_idx=0, total_iterations=1):
            doc = fitz.open(input_path)
            processed_images = 0
            total_pages = len(doc)
            
            # Use a set to avoid processing the same image twice (PDF optimization)
            processed_xrefs = set()
            
            for page_index in range(len(doc)):
                page = doc[page_index]
                image_list = page.get_images(full=True)
                for img_info in image_list:
                    xref = img_info[0]
                    if xref in processed_xrefs:
                        continue
                    processed_xrefs.add(xref)
                    
                    try:
                        # Try extracting first
                        pix = fitz.Pixmap(doc, xref)
                        
                        # If it's a mask or has transparency, convert to RGB
                        if pix.n >= 4:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        
                        img_bytes = pix.tobytes("jpeg")
                        old_size = len(img_bytes)
                        
                        pil_img = Image.open(io.BytesIO(img_bytes))
                        
                        # Downscale if huge
                        max_dim = 2000
                        if pil_img.width > max_dim or pil_img.height > max_dim:
                            pil_img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
                            
                        out_buf = io.BytesIO()
                        pil_img.save(out_buf, format="JPEG", quality=quality, optimize=True)
                        compressed_data = out_buf.getvalue()
                        new_size = len(compressed_data)
                        
                        # Only update if we actually saved space
                        if new_size < old_size or quality < 50:
                            page.replace_image(xref, stream=compressed_data)
                            processed_images += 1
                        
                        pix = None # free memory
                    except Exception:
                        continue
                
                # Report progress per page
                if progress_callback:
                    # Base progress for current file completion
                    file_prog = (page_index + 1) / total_pages
                    
                    # Scaled progress: 
                    # sub_file_prog is progress within the current file (0 to 1) 
                    # taking into account multiple iterations
                    sub_file_prog = (iteration_idx + file_prog) / total_iterations
                    overall_prog = (current_file_idx + sub_file_prog) / total_files
                    progress_callback(overall_prog)
            
            doc.save(out_path, garbage=4, deflate=True, clean=True, deflate_images=True)
            doc.close()
            return os.path.getsize(out_path)

        # For fixed size targets, use binary search on quality
        if mode in ("100", "200", "500", "1024", "custom"):
            target_map = {
                "100": 100 * 1024,
                "200": 200 * 1024,
                "500": 500 * 1024,
                "1024": 1024 * 1024,
            }
            if mode == "custom" and target_kb:
                target_bytes = int(target_kb) * 1024
            else:
                target_bytes = target_map.get(mode, 500 * 1024)

            low, high = 5, 85
            for i in range(8):  # max 8 iterations
                mid = (low + high) // 2
                result_size = try_compress_at_quality(
                    pdf_path, output_path, mid, current_file_idx, total_files, 
                    iteration_idx=i, total_iterations=9) 
                if result_size <= target_bytes:
                    low = mid + 1
                else:
                    high = mid - 1
            # Final compress at best quality found (iteration 8 of 9)
            try_compress_at_quality(pdf_path, output_path, max(5, high), current_file_idx, total_files, 
                                   iteration_idx=8, total_iterations=9)
        else:
            # Default Auto Best — quality 35 for strong compression
            try_compress_at_quality(pdf_path, output_path, 35, current_file_idx, total_files)
        
        return os.path.exists(output_path)

    @staticmethod
    def add_watermark(pdf_path: str, watermark_text: str, output_path: str) -> None:
        """
        Adds a diagonal text watermark. Creates a temporary PDF with the text and merges it.
        """
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import inch
        from reportlab.lib.colors import Color
        
        # Create watermark PDF in memory
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=A4)
        c.translate(inch, inch)
        c.setFont("Helvetica-Bold", 60)
        c.setFillColor(Color(0, 0, 0, alpha=0.3)) # 30% opacity
        c.rotate(45)
        # Draw string roughly in the middle
        c.drawString(2 * inch, 0, watermark_text)
        c.save()
        
        packet.seek(0)
        watermark_pdf = PdfReader(packet)
        watermark_page = watermark_pdf.pages[0]
        
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        for page in reader.pages:
            page.merge_page(watermark_page)
            writer.add_page(page)
            
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

    @staticmethod
    def encrypt_pdf(pdf_path: str, output_path: str, password: str) -> None:
        """Adds a password to the PDF."""
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        for page in reader.pages:
            writer.add_page(page)
            
        writer.encrypt(password)
        with open(output_path, "wb") as out_file:
            writer.write(out_file)
            
    @staticmethod
    def decrypt_pdf(pdf_path: str, output_path: str, password: str) -> bool:
        """Removes password from a PDF if known. Returns True if successful."""
        reader = PdfReader(pdf_path)
        if reader.is_encrypted:
            try:
                # Decrypt returns 1 for user pw, 2 for owner pw, 0 if failed
                success = reader.decrypt(password)
                if not success:
                    return False
            except Exception:
                return False
                
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
            
        with open(output_path, "wb") as out_file:
            writer.write(out_file)
            
        return True
