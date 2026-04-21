import os
import re
import fitz # PyMuPDF
import numpy as np
from PIL import Image, ImageOps, ImageEnhance
from core.ocr_engine import OCREngine
from core.invoice_field_extractors import InvoiceFieldExtractors
from core.invoice_postprocessor import InvoicePostprocessor
from core.invoice_validation import InvoiceValidation

class InvoiceParserEngine:
    """
    Main orchestration engine for extracting structured data from invoice PDFs/Images.
    """
    
    @staticmethod
    def parse_invoice(file_path: str, detect_multi: bool = True) -> list:
        """
        Segments a PDF into units and parses each.
        Returns a list of dictionaries (one per invoice).
        """
        if not os.path.exists(file_path):
            return [{
                "Filename": os.path.basename(file_path),
                "Status": "Failed",
                "Remarks": "File not found"
            }]

        ext = os.path.splitext(file_path)[1].lower()
        if ext not in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp']:
            return [{
                "Filename": os.path.basename(file_path),
                "Status": "Failed",
                "Remarks": f"Unsupported: {ext}"
            }]

        # 1. SEGMENTATION (One unit for images, Multiple for segmented PDFs)
        units = []
        if ext == '.pdf':
            if detect_multi:
                units = InvoiceParserEngine.segment_pdf(file_path)
            else:
                # LEGACY / STANDARD MODE: Treat whole PDF as one unit
                raw_text = InvoiceParserEngine._extract_pdf_hybrid(file_path)
                units = [{"text": raw_text, "pages": "1", "ocr": False}]
        else:
            # Single Image
            raw_text = InvoiceParserEngine._extract_with_preprocessed_ocr(file_path)
            units = [{"text": raw_text, "pages": "1", "ocr": True}]

        all_results = []
        
        for unit in units:
            raw_text = unit["text"]
            page_range = unit.get("pages", "1")
            used_ocr = unit.get("ocr", False)
            
            result = {
                "Filename": os.path.basename(file_path),
                "Page Range": page_range,
                "Status": "Failed",
                "Invoice No": "",
                "Date": "",
                "Party Name": "",
                "Buyer GSTIN": "",
                "Taxable Value": "",
                "CGST": "", "CGST %": "", 
                "SGST": "", "SGST %": "", 
                "IGST": "", "IGST %": "",
                "Grand Total": "",
                "Confidence": "Low",
                "Remarks": ""
            }

            if not raw_text.strip():
                result["Remarks"] = "No text extracted"
                all_results.append(result)
                continue

            try:
                clean_text = InvoicePostprocessor.normalize_text(raw_text)
                
                # 2. Field Extraction
                b_gstin = InvoiceFieldExtractors.extract_buyer_gstin(clean_text)
                inv_no = InvoiceFieldExtractors.extract_invoice_number(clean_text)
                inv_date = InvoiceFieldExtractors.extract_date(clean_text)
                amounts = InvoiceFieldExtractors.extract_amounts(clean_text)
                party_name = InvoiceFieldExtractors.extract_party_name(clean_text)
                
                # 3. Formatting & Mapping
                result["Buyer GSTIN"] = b_gstin.upper()
                result["Invoice No"] = InvoicePostprocessor.clean_invoice_number(inv_no)
                result["Date"] = InvoicePostprocessor.normalize_date(inv_date)
                result["Party Name"] = InvoicePostprocessor.clean_party_name(party_name)
                
                result["Taxable Value"] = amounts["taxable"]
                result["CGST"] = amounts["cgst"]
                result["CGST %"] = amounts["cgst_rate"]
                result["SGST"] = amounts["sgst"]
                result["SGST %"] = amounts["sgst_rate"]
                result["IGST"] = amounts["igst"]
                result["IGST %"] = amounts["igst_rate"]
                result["Grand Total"] = amounts["grand_total"]
                
                # 4. Validation
                guard_status, guard_remark = InvoiceValidation.validate_field_guards(result, clean_text)
                math_ok, math_remark = InvoiceValidation.validate_totals(
                    result["Taxable Value"], result["CGST"], result["SGST"], result["IGST"], result["Grand Total"]
                )
                
                valid_fields = sum([
                    bool(result["Invoice No"]), 
                    bool(result["Date"]), 
                    bool(result["Grand Total"]),
                    bool(result["Party Name"])
                ])
                
                remarks = []
                if used_ocr: remarks.append("OCR Mode")
                if math_remark: remarks.append(math_remark)
                if guard_remark: remarks.append(guard_remark)
                result["Remarks"] = " | ".join([r for r in remarks if r])
                
                if valid_fields >= 4 and math_ok and guard_status:
                    result["Status"] = "Parsed"
                    result["Confidence"] = "High"
                elif valid_fields >= 2:
                    result["Status"] = "Needs Review"
                    result["Confidence"] = "Medium"
                else:
                    result["Status"] = "Partial"
                    result["Confidence"] = "Low"
                    
                all_results.append(result)

            except Exception as e:
                result["Remarks"] = f"Unit Error: {str(e)}"
                all_results.append(result)

        return all_results

    @staticmethod
    def segment_pdf(file_path: str) -> list:
        """Segments a PDF into logical invoice units."""
        doc = fitz.open(file_path)
        units = []
        current_pages = []
        current_text = ""
        last_inv_no = None
        
        for i in range(len(doc)):
            page = doc[i]
            p_text = page.get_text("text", sort=True)
            used_ocr = False
            
            if len(p_text.strip()) < 100:
                # Try OCR for this page
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                temp = os.path.join(os.environ.get('TEMP', '/tmp'), f"page_{i}.png")
                InvoiceParserEngine._preprocess_image(img).save(temp)
                p_text = OCREngine.extract_text_from_image(temp)
                used_ocr = True
                try: os.remove(temp)
                except: pass
            
            norm_text = InvoicePostprocessor.normalize_text(p_text)
            
            # Strict boundary detection for segmentation
            # Use finditer to find the actual invoice number, skipping template noise
            inv_no = None
            candidates = []
            for m in re.finditer(r'(?i)(?:\bInvoice No|\bInv\.? No|\bBill No|\bNo\b|#)\.?\s*[\.\:\-\s]*([A-Z0-9\-\/]{3,25})', norm_text):
                cand = m.group(1).strip()
                # Check line prefix for address-specific labels
                start_p = max(0, m.start()-30)
                prefix = norm_text[start_p:m.start()].upper()
                label = m.group(0).upper()
                
                # REJECT if it's explicitly part of an address or template labels
                if any(x in prefix for x in ['PLOT', 'SHED', 'BLOCK', 'GIDC', 'ACCOUNT', 'A/C', 'PHONE', 'MO.', 'ACK', 'IRN', 'GENERATED']): continue
                
                # REJECT typical template words that land in 'No.' (e.g. No. Original, No. Date)
                junk = ['ORIGINAL', 'DATE', 'GUJARAT', 'STATE', 'DETAILS', 'DESCRIPTION', 'PLACE', 'TRANSPORT', 'PAN', 'PIN', 'HSN', 'SAC', 'QTY', 'RATE', 'AMOUNT', 'TAKEN BACK', 'PAGE', 'INVOICE']
                if any(j in cand.upper() for j in junk): continue
                
                score = 0
                if 'INVOICE' in label or 'INV' in label or 'BILL' in label: score += 2
                if '/' in cand or '-' in cand: score += 1
                if cand.isdigit(): score += 0.5
                candidates.append((cand, score, label))
                print(f"  Page {i+1} cand={cand} score={score} label={label}")
            
            if candidates:
                candidates.sort(key=lambda x: x[1], reverse=True)
                inv_no = candidates[0][0]
            
            # 2. Markers
            has_bf = any(x in norm_text.upper() for x in ["B/F", "BROUGHT FORWARD", "CONTINUED FROM"])
            has_cf = any(x in norm_text.upper() for x in ["C/F", "CARRY FORWARD", "CONTINUED TO"])
            is_new_header = any(x in norm_text.upper() for x in ["TAX INVOICE", "DETAILS OF RECEIVER", "DETAILS OF RECIPIENT", "BILL TO", "BUYER", "DETAILS OF BUYER"])
            
            is_start = False
            
            if inv_no:
                clean_inv = InvoicePostprocessor.clean_invoice_number(inv_no)
                if last_inv_no is None:
                    is_start = True
                elif clean_inv != last_inv_no:
                    # New number found. 
                    # Only split if it's not a continuation (B/F check)
                    if not has_bf:
                        is_start = True
                last_inv_no = clean_inv
            elif is_new_header and last_inv_no is None:
                # First page of document, start first segment
                is_start = True
            # REMOVED: splitting on header alone when last_inv_no is set.
            # This prevents pages with 'Tax Invoice' logo but no new number (footers/continuations) from starting new segments.

            if is_start and current_pages:
                units.append({
                    "text": current_text,
                    "pages": InvoiceParserEngine._format_page_range(current_pages),
                    "ocr": used_ocr
                })
                current_pages = []
                current_text = ""
                current_text = ""
            
            current_pages.append(i + 1)
            current_text += "\n" + p_text
        
        if current_pages:
            units.append({
                "text": current_text,
                "pages": InvoiceParserEngine._format_page_range(current_pages),
                "ocr": used_ocr
            })
            
        doc.close()
        return units

    @staticmethod
    def _format_page_range(pages: list) -> str:
        if not pages: return ""
        if len(pages) == 1: return str(pages[0])
        return f"{pages[0]}-{pages[-1]}"



    @staticmethod
    def _extract_with_preprocessed_ocr(file_path: str) -> str:
        """Enhances image before OCR."""
        try:
            if file_path.lower().endswith('.pdf'):
                doc = fitz.open(file_path)
                full_text = ""
                for i in range(len(doc)):
                    page = doc[i]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    enhanced_img = InvoiceParserEngine._preprocess_image(img)
                    temp_p = os.path.join(os.environ.get('TEMP', '/tmp'), f"ocr_pre_{i}.png")
                    enhanced_img.save(temp_p)
                    page_text = OCREngine.extract_text_from_image(temp_p)
                    full_text += f"\n{page_text}"
                    try: os.remove(temp_p)
                    except: pass
                doc.close()
                return full_text
            else:
                img = Image.open(file_path)
                enhanced_img = InvoiceParserEngine._preprocess_image(img)
                temp_p = os.path.join(os.environ.get('TEMP', '/tmp'), "ocr_pre_img.png")
                enhanced_img.save(temp_p)
                raw_text = OCREngine.extract_text_from_image(temp_p)
                try: os.remove(temp_p)
                except: pass
                return raw_text
        except Exception as e:
            print(f"OCR Preproc error: {e}")
            return OCREngine.extract_text_from_pdf(file_path) if file_path.lower().endswith('.pdf') else OCREngine.extract_text_from_image(file_path)

    @staticmethod
    def _preprocess_image(img: Image.Image) -> Image.Image:
        """Grayscale, Contrast, Sharpness."""
        img = ImageOps.grayscale(img)
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        return img

    @staticmethod
    def _extract_pdf_hybrid(file_path: str) -> str:
        """Native PDF Text extraction."""
        try:
            doc = fitz.open(file_path)
            all_text = ""
            for i in range(len(doc)):
                all_text += f"\n{doc[i].get_text('text', sort=True)}"
            doc.close()
            return all_text
        except:
            return ""
