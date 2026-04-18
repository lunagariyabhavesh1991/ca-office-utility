import fitz  # PyMuPDF
import os
import glob
from PIL import Image

class PDFEditorEngine:
    def __init__(self):
        self.doc = None
        self.ocr_page_data = {} # page_idx -> list of spans from OCR
        self.current_page_cache = None
        self.cached_page_index = -1
        self.current_path = None
        self.temp_path = None
        self._system_font_cache = {}  # font_name -> file_path cache

    def open_pdf(self, path):
        """Opens a PDF and creates a temporary working copy."""
        self.current_path = path
        self.doc = fitz.open(path)
        return len(self.doc)

    def get_page_image(self, page_num, zoom=1.0):
        """Renders a page to a PIL Image."""
        if not self.doc:
            return None
        
        page = self.doc[page_num]
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def get_thumbnails(self, size=(150, 200)):
        """Returns a list of thumbnail images for all pages."""
        thumbs = []
        if not self.doc:
            return thumbs
            
        for page in self.doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail(size)
            thumbs.append(img)
        return thumbs

    def get_page_objects(self, page_num):
        """Returns a list of text spans and images with coordinates. 
        Enriches text spans with underline detection. Cached for performance."""
        if not self.doc:
            return []
            
        if self.cached_page_index == page_num and self.current_page_cache is not None:
            return self.current_page_cache

        page = self.doc[page_num]
        dict_data = page.get_text("dict")
        objects = []
        
        # 1. Native PDF Objects
        for b in dict_data["blocks"]:
            if b["type"] == 0:  # Text block
                for line in b["lines"]:
                    for span in line["spans"]:
                        span["type"] = "text"
                        # Detect underline for this span
                        span["is_underlined"] = self._detect_underline(page, span["bbox"])
                        objects.append(span)
            elif b["type"] == 1:  # Image block
                b["type"] = "image"
                objects.append(b)
        
        # 2. Synthetic OCR Objects (if any)
        if page_num in self.ocr_page_data:
            for span in self.ocr_page_data[page_num]:
                span["synthetic"] = True
                objects.append(span)

        # 3. Form Widgets (Signatures)
        for widget in page.widgets():
            if widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
                objects.append({
                    "type": "signature",
                    "bbox": list(widget.rect),
                    "xref": widget.xref,
                    "field_name": widget.field_name
                })

        self.current_page_cache = objects
        self.cached_page_index = page_num
        return objects

    def invalidate_cache(self):
        """Clears the object cache. Call after edits or page changes."""
        self.current_page_cache = None
        self.cached_page_index = -1

    def run_ocr_on_page(self, page_num):
        """Performs OCR on the specified page and stores synthetic text objects."""
        if not self.doc: return False
        try:
            from .ocr_engine import OCREngine
            page = self.doc[page_num]
            # Get high-res pixmap for better OCR
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            
            results = OCREngine.get_text_with_bboxes(img_data)
            if results:
                formatted_results = []
                for res in results:
                    formatted_results.append({
                        "type": "text",
                        "text": res["text"],
                        "bbox": res["bbox"],
                        "size": 11,
                        "font": "helv",
                        "color": 0,
                        "synthetic": True
                    })
                self.ocr_page_data[page_num] = formatted_results
                return True
        except Exception:
            pass
        return False

    def replace_image(self, page_num, block, image_path):
        """Replaces an image at the specified location."""
        if not self.doc:
            return False
            
        page = self.doc[page_num]
        bbox = block["bbox"]
        
        page.add_redact_annot(bbox)
        page.apply_redactions()
        
        page.insert_image(bbox, filename=image_path)
        return True

    def delete_object(self, page_num, block):
        """Deletes (redacts) or removes (widget) the specified object from the PDF."""
        if not self.doc: return False
        page = self.doc[page_num]
        
        if block.get("type") == "signature" and "xref" in block:
            # For signatures, we delete the actual widget
            try:
                for widget in page.widgets():
                    if widget.xref == block["xref"]:
                        page.delete_widget(widget)
                        break
                self.invalidate_cache()
                return True
            except Exception:
                pass

        # Fallback to redaction for text/images
        bbox = fitz.Rect(block.get("bbox", [0,0,0,0]))
        page.add_redact_annot(bbox + (-0.5, -0.5, 0.5, 0.5), fill=(1, 1, 1))
        page.apply_redactions()
        self.invalidate_cache()
        return True

    # ─── TEXT EDITING WITH FORMAT PRESERVATION ───────────────────────────

    def update_text(self, page_num, block, new_text, force_bold=None, force_underline=None):
        """Redacts old text and inserts new text preserving original formatting.
        
        Strategy priority:
          A) Use the original embedded font from the PDF (best fidelity)
          B) Find matching system font on disk (good fidelity)
          C) Map to Base14 built-in font (fallback)
        """
        if not self.doc:
            return False, "No document loaded"

        page = self.doc[page_num]
        bbox = fitz.Rect(block.get("bbox", [0, 0, 0, 0]))
        if bbox.width == 0 or bbox.height == 0:
            return False, "Invalid selection"

        # ── 1. Capture original underline drawings BEFORE redacting ──
        orig_underline_info = self._get_underline_drawings(page, bbox)
        
        # Early underline detection for redaction sizing
        has_underline = (force_underline if force_underline is not None 
                         else block.get("is_underlined", False) 
                         or orig_underline_info.get("found", False))

        # ── 2. Redact the area with white fill ──
        # If underlined, expand redaction area downward to also erase old underline drawings
        if has_underline:
            redact_rect = bbox + (-0.5, -0.5, 0.5, 4.0)  # Extra 4pt below to cover underline
        else:
            redact_rect = bbox + (-0.5, -0.5, 0.5, 0.5)
        page.add_redact_annot(redact_rect, fill=(1, 1, 1))
        page.apply_redactions()

        # ── 3. Extract original text properties ──
        fontname_orig = block.get("font", "helv")
        fontname_lower = fontname_orig.lower()
        fontsize = block.get("size", 11)
        color = self._hex_to_rgb(block.get("color", 0))
        flags = block.get("flags", 0)
        origin = block.get("origin", (bbox.x0, bbox.y1 - 2.0))

        # Bold detection
        heavier_terms = ["bold", "heavy", "black", "semibold", "medium",
                         "500", "600", "700", "800", "demi"]
        is_originally_bold = (flags & 16) or any(t in fontname_lower for t in heavier_terms)
        is_bold = force_bold if force_bold is not None else is_originally_bold

        # Italic detection
        is_italic = (flags & 2) or "italic" in fontname_lower or "oblique" in fontname_lower

        # Underline: use explicit override, else detected value from span
        is_underlined = force_underline if force_underline is not None else block.get("is_underlined", False)

        warning = None

        # ── 4. Try to resolve the best font ──
        font_obj, font_label, strategy_used = self._resolve_font(
            page, fontname_orig, is_bold, is_italic
        )

        # ── 5. Calculate width adjustment ──
        # Measure original text width from bbox
        orig_text_width = bbox.width
        adjusted_fontsize = fontsize

        try:
            if font_obj is not None:
                # Measure new text width with the resolved font
                new_text_width = font_obj.text_length(new_text, fontsize=fontsize)
                if new_text_width > 0 and orig_text_width > 0 and "\n" not in new_text:
                    ratio = orig_text_width / new_text_width
                    # Only scale DOWN if too wide (ratio < 0.95)
                    # Do NOT scale UP if shorter (ratio >= 1.0)
                    if ratio < 0.95:
                        adjusted_fontsize = fontsize * ratio
                        # Clamp: don't shrink below 60% of original
                        adjusted_fontsize = max(fontsize * 0.6, adjusted_fontsize)
                    else:
                        # Keep original size if it fits easily
                        adjusted_fontsize = fontsize
        except Exception:
            pass  # text_length not available for all font types

        # ── 6. Insert new text ──
        try:
            if "\n" not in new_text:
                if font_obj is not None and strategy_used in ("embedded", "system"):
                    # Strategy A/B: use the resolved Font object
                    page.insert_text(
                        origin, new_text,
                        fontname=font_label,
                        fontfile=None,  # already registered
                        fontsize=adjusted_fontsize,
                        color=color,
                    )
                    if strategy_used == "system":
                        warning = f"Using system font match for '{fontname_orig}'"
                else:
                    # Strategy C: Base14 fallback
                    base14_name = self._map_to_base14(fontname_lower, is_bold, is_italic, flags)
                    page.insert_text(
                        origin, new_text,
                        fontname=base14_name,
                        fontsize=adjusted_fontsize,
                        color=color,
                    )
                    if strategy_used == "base14_mapped":
                        warning = f"Approximate font match for '{fontname_orig}'"
            else:
                # Multi-line text
                if font_obj is not None and strategy_used in ("embedded", "system"):
                    page.insert_textbox(
                        bbox, new_text,
                        fontname=font_label,
                        fontsize=fontsize,
                        color=color,
                        align=block.get("align", 0),
                    )
                else:
                    base14_name = self._map_to_base14(fontname_lower, is_bold, is_italic, flags)
                    page.insert_textbox(
                        bbox, new_text,
                        fontname=base14_name,
                        fontsize=fontsize,
                        color=color,
                        align=block.get("align", 0),
                    )

            # ── 7. Restore underline under actual new text ──
            if is_underlined:
                # Calculate the actual width of the new text for precise underline
                actual_new_width = None
                try:
                    if font_obj is not None:
                        actual_new_width = font_obj.text_length(new_text, fontsize=adjusted_fontsize)
                except Exception:
                    pass
                if actual_new_width is None:
                    actual_new_width = len(new_text) * adjusted_fontsize * 0.5
                
                # Build a rect that covers exactly the new text for underline detection
                # We use origin[1] (baseline) as the primary anchor for the new underline
                new_text_rect = fitz.Rect(
                    origin[0],                        # start x = same as text origin
                    origin[1] - adjusted_fontsize,     # top y
                    origin[0] + actual_new_width,      # end x = actual text end
                    origin[1] + adjusted_fontsize * 0.2 # bottom y (slightly below baseline)
                )
                self._restore_underline(page, origin, new_text, adjusted_fontsize,
                                        color, font_obj, orig_underline_info, bbox=new_text_rect)

            self.invalidate_cache()
            return True, warning

        except Exception as e:
            # Absolute last-resort fallback
            try:
                page.insert_text(
                    (bbox.x0, bbox.y1 - 2.5), new_text,
                    fontname="helv", fontsize=fontsize, color=color,
                )
                warning = f"Fallback insert used: {str(e)}"
            except Exception as e2:
                warning = f"Critical error: {str(e2)}"

        self.invalidate_cache()
        return True, warning

    # ─── FONT RESOLUTION STRATEGIES ─────────────────────────────────────

    def _resolve_font(self, page, fontname_orig, is_bold, is_italic):
        """Tries to resolve the best matching font. Returns (Font_obj, label, strategy)."""

        # Strategy A: Try to extract the embedded font from the PDF
        try:
            font_obj, label = self._get_embedded_font(page, fontname_orig)
            if font_obj is not None:
                return font_obj, label, "embedded"
        except Exception:
            pass

        # Strategy B: Try to find a matching system font file
        try:
            font_obj, label = self._get_system_font(page, fontname_orig, is_bold, is_italic)
            if font_obj is not None:
                return font_obj, label, "system"
        except Exception:
            pass

        # Strategy C: Base14 fallback (no Font object, just a name string)
        return None, None, "base14_mapped"

    def _get_embedded_font(self, page, fontname_orig):
        """Extracts the embedded font from the PDF and registers it on the page.
        Returns (fitz.Font, registered_name) or (None, None)."""
        if not self.doc:
            return None, None

        # Iterate page fonts and find the matching one
        font_list = page.get_fonts(full=True)  # [(xref, ext, type, basefont, name, encoding), ...]
        
        best_match = None
        fontname_lower = fontname_orig.lower().replace(" ", "").replace("-", "")

        for xref, ext, ftype, basefont, name, encoding in font_list:
            # Match by basefont or name
            check_names = [
                basefont.lower().replace(" ", "").replace("-", ""),
                name.lower().replace(" ", "").replace("-", ""),
            ]
            for cn in check_names:
                if cn == fontname_lower or fontname_lower in cn or cn in fontname_lower:
                    best_match = (xref, ext, ftype, basefont, name)
                    break
            if best_match:
                break

        if best_match is None:
            return None, None

        xref = best_match[0]
        try:
            # Extract font binary data from PDF
            font_data = self.doc.extract_font(xref)
            # font_data = (basename, ext, subtype, content_bytes)
            if font_data and len(font_data) >= 4 and font_data[3]:
                font_buffer = font_data[3]
                if len(font_buffer) > 100:  # Valid font data
                    font_obj = fitz.Font(fontbuffer=font_buffer)
                    # Register on page with a unique label
                    label = f"f{xref}"
                    page.insert_font(fontname=label, fontbuffer=font_buffer)
                    return font_obj, label
        except Exception:
            pass

        return None, None

    def _get_system_font(self, page, fontname_orig, is_bold, is_italic):
        """Finds a matching system font file and registers it on the page.
        Returns (fitz.Font, registered_name) or (None, None)."""
        font_path = self._find_system_font(fontname_orig, is_bold, is_italic)
        if font_path is None:
            return None, None

        try:
            font_obj = fitz.Font(fontfile=font_path)
            # Create a unique label from the file name
            label = "sys_" + os.path.splitext(os.path.basename(font_path))[0][:20]
            label = label.replace(" ", "").replace("-", "_")
            page.insert_font(fontname=label, fontfile=font_path)
            return font_obj, label
        except Exception:
            return None, None

    def _find_system_font(self, fontname_orig, is_bold=False, is_italic=False):
        """Maps a PDF font name to a system font file path on Windows.
        Uses a curated mapping + directory scan with fuzzy matching."""

        # Check cache first
        cache_key = f"{fontname_orig}_{is_bold}_{is_italic}"
        if cache_key in self._system_font_cache:
            return self._system_font_cache[cache_key]

        font_dirs = [
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
            os.path.expanduser("~\\AppData\\Local\\Microsoft\\Windows\\Fonts"),
        ]

        # Clean font name for matching
        fontname_clean = fontname_orig.lower()
        # Remove common prefixes added by PDF subset embedding (e.g., "ABCDEF+Arial-Bold")
        if "+" in fontname_clean:
            fontname_clean = fontname_clean.split("+", 1)[1]
        # Remove style suffixes for base name matching
        for suffix in ["-bold", "-italic", "-bolditalic", "-oblique", 
                       ",bold", ",italic", ",bolditalic",
                       "-regular", ",regular", "-roman"]:
            fontname_clean = fontname_clean.replace(suffix, "")
        fontname_clean = fontname_clean.replace("-", "").replace(" ", "").strip()

        # Curated mapping of common font families → Windows font file basenames
        FONT_MAP = {
            "arial": {"reg": "arial", "bold": "arialbd", "italic": "ariali", "bi": "arialbi"},
            "calibri": {"reg": "calibri", "bold": "calibrib", "italic": "calibrii", "bi": "calibriz"},
            "timesnewroman": {"reg": "times", "bold": "timesbd", "italic": "timesi", "bi": "timesbi"},
            "times": {"reg": "times", "bold": "timesbd", "italic": "timesi", "bi": "timesbi"},
            "cambria": {"reg": "cambria", "bold": "cambriab", "italic": "cambriai", "bi": "cambriaz"},
            "georgia": {"reg": "georgia", "bold": "georgiab", "italic": "georgiai", "bi": "georgiaz"},
            "verdana": {"reg": "verdana", "bold": "verdanab", "italic": "verdanai", "bi": "verdanaz"},
            "tahoma": {"reg": "tahoma", "bold": "tahomabd", "italic": "tahoma", "bi": "tahomabd"},
            "trebuchetms": {"reg": "trebuc", "bold": "trebucbd", "italic": "trebucit", "bi": "trebucbi"},
            "segoeui": {"reg": "segoeui", "bold": "segoeuib", "italic": "segoeuii", "bi": "segoeuiz"},
            "consolas": {"reg": "consola", "bold": "consolab", "italic": "consolai", "bi": "consolaz"},
            "couriernew": {"reg": "cour", "bold": "courbd", "italic": "couri", "bi": "courbi"},
            "courier": {"reg": "cour", "bold": "courbd", "italic": "couri", "bi": "courbi"},
            "roboto": {"reg": "roboto", "bold": "roboto", "italic": "roboto", "bi": "roboto"},
            "garamond": {"reg": "garamond", "bold": "garamond", "italic": "garamond", "bi": "garamond"},
            "bookantiqua": {"reg": "bkant", "bold": "bkant", "italic": "bkant", "bi": "bkant"},
            "palatino": {"reg": "pala", "bold": "palab", "italic": "palai", "bi": "palabi"},
            "comicsansms": {"reg": "comic", "bold": "comicbd", "italic": "comici", "bi": "comicz"},
            "impact": {"reg": "impact", "bold": "impact", "italic": "impact", "bi": "impact"},
            "lucidaconsole": {"reg": "lucon", "bold": "lucon", "italic": "lucon", "bi": "lucon"},
            "robotoslab": {"reg": "RobotoSlab-Regular", "bold": "RobotoSlab-Bold", "italic": "RobotoSlab-Regular", "bi": "RobotoSlab-Bold"},
        }

        # Determine variant key
        if is_bold and is_italic:
            variant = "bi"
        elif is_bold:
            variant = "bold"
        elif is_italic:
            variant = "italic"
        else:
            variant = "reg"

        # Try curated mapping first
        for map_key, variants in FONT_MAP.items():
            if map_key in fontname_clean or fontname_clean in map_key:
                target_basename = variants.get(variant, variants["reg"])
                for font_dir in font_dirs:
                    for ext in ["ttf", "otf", "TTF", "OTF"]:
                        candidate = os.path.join(font_dir, f"{target_basename}.{ext}")
                        if os.path.isfile(candidate):
                            self._system_font_cache[cache_key] = candidate
                            return candidate

        # Fuzzy scan: look for font files whose name contains the cleaned font name
        for font_dir in font_dirs:
            if not os.path.isdir(font_dir):
                continue
            try:
                for fname in os.listdir(font_dir):
                    fname_lower = fname.lower().replace("-", "").replace("_", "").replace(" ", "")
                    if fontname_clean in fname_lower and fname_lower.endswith((".ttf", ".otf")):
                        # Check for bold/italic variant match
                        if is_bold and "bold" not in fname_lower and "bd" not in fname_lower:
                            continue
                        if is_italic and "italic" not in fname_lower and "it" not in fname_lower and "i." not in fname_lower:
                            continue
                        full_path = os.path.join(font_dir, fname)
                        self._system_font_cache[cache_key] = full_path
                        return full_path
            except OSError:
                continue

        # Second pass fuzzy scan without style filtering (better to have wrong weight than wrong family)
        for font_dir in font_dirs:
            if not os.path.isdir(font_dir):
                continue
            try:
                for fname in os.listdir(font_dir):
                    fname_lower = fname.lower().replace("-", "").replace("_", "").replace(" ", "")
                    if fontname_clean in fname_lower and fname_lower.endswith((".ttf", ".otf")):
                        full_path = os.path.join(font_dir, fname)
                        self._system_font_cache[cache_key] = full_path
                        return full_path
            except OSError:
                continue

        self._system_font_cache[cache_key] = None
        return None

    def _map_to_base14(self, fontname_lower, is_bold, is_italic, flags):
        """Maps to the closest Base14 built-in font (last resort)."""
        serif_names = ["times", "serif", "bookman", "georgia", "palatino",
                       "garamond", "century", "cambria", "didot"]
        is_serif = (flags & 4) or any(x in fontname_lower for x in serif_names)

        mono_names = ["courier", "mono", "consolas", "lucida", "code", "fixed"]
        is_mono = (flags & 8) or any(x in fontname_lower for x in mono_names)

        if is_serif:
            if is_bold and is_italic: return "biit"
            elif is_bold: return "tibo"
            elif is_italic: return "tiit"
            else: return "tiro"
        elif is_mono:
            if is_bold and is_italic: return "boit"
            elif is_bold: return "cobo"
            elif is_italic: return "coit"
            else: return "cour"
        else:
            if is_bold and is_italic: return "hbit"
            elif is_bold: return "hebo"
            elif is_italic: return "heit"
            else: return "helv"

    # ─── UNDERLINE DETECTION & RESTORATION ───────────────────────────────

    def _detect_underline(self, page, span_bbox):
        """Detects if a text span is underlined by checking for horizontal lines
        drawn just below the text baseline in the page's drawing commands and annotations."""
        try:
            rect = fitz.Rect(span_bbox)
            baseline_y = rect.y1
            # Underline typically appears within a few points below the baseline
            tolerance_y = 4.0

            # Check page drawings (vector graphics / stroked lines)
            drawings = page.get_drawings()
            for d in drawings:
                if d.get("type") == "s" or d.get("items"):  # stroke or path items
                    for item in d.get("items", []):
                        if item[0] == "l":  # line segment
                            p1, p2 = item[1], item[2]
                            # Horizontal line near text bottom?
                            if (abs(p1.y - p2.y) < 1.5 and  # nearly horizontal
                                abs(p1.y - baseline_y) < tolerance_y and
                                p1.x <= rect.x1 + 2 and p2.x >= rect.x0 - 2 and
                                abs(p2.x - p1.x) > rect.width * 0.3):  # spans significant width
                                return True

            # Check annotations (underline annotations)
            for annot in page.annots():
                if annot.type[0] == 9:  # Underline annotation type
                    annot_rect = annot.rect
                    # Check overlap with span bbox
                    if (annot_rect.x0 <= rect.x1 + 2 and annot_rect.x1 >= rect.x0 - 2 and
                        annot_rect.y0 <= rect.y1 + tolerance_y and annot_rect.y1 >= rect.y0):
                        return True

        except Exception:
            pass
        return False

    def _get_underline_drawings(self, page, text_bbox):
        """Before redacting, capture underline line info (position, thickness, color)."""
        info = {"found": False, "thickness": 1.0, "color": (0, 0, 0)}
        try:
            rect = fitz.Rect(text_bbox)
            baseline_y = rect.y1
            tolerance_y = 4.0

            drawings = page.get_drawings()
            for d in drawings:
                stroke_color = d.get("color", (0, 0, 0))
                line_width = d.get("width", 1.0)
                for item in d.get("items", []):
                    if item[0] == "l":
                        p1, p2 = item[1], item[2]
                        if (abs(p1.y - p2.y) < 1.5 and
                            abs(p1.y - baseline_y) < tolerance_y and
                            p1.x <= rect.x1 + 2 and p2.x >= rect.x0 - 2 and
                            abs(p2.x - p1.x) > rect.width * 0.3):
                            info = {"found": True, "thickness": line_width or 1.0,
                                    "color": stroke_color or (0, 0, 0),
                                    "y_offset": p1.y - baseline_y}
                            return info

            # Check underline annotations
            for annot in page.annots():
                if annot.type[0] == 9:
                    annot_rect = annot.rect
                    if (annot_rect.x0 <= rect.x1 + 2 and annot_rect.x1 >= rect.x0 - 2 and
                        annot_rect.y0 <= rect.y1 + tolerance_y and annot_rect.y1 >= rect.y0):
                        a_colors = annot.colors
                        stroke = a_colors.get("stroke", (0, 0, 0)) if a_colors else (0, 0, 0)
                        info = {"found": True, "thickness": 1.0, "color": stroke, "y_offset": 1.5}
                        return info
        except Exception:
            pass
        return info

    def _restore_underline(self, page, origin, text, fontsize, color, font_obj, orig_info, bbox=None):
        """Draws an underline below newly inserted text, matching original style."""
        try:
            # Use origin[1] as baseline if possible, it's more reliable than bbox
            x_start = bbox.x0 if bbox is not None else origin[0]
            
            if bbox is not None:
                x_end = bbox.x1
            else:
                # Calculate text width
                if font_obj is not None:
                    try:
                        text_width = font_obj.text_length(text, fontsize=fontsize)
                    except Exception:
                        text_width = len(text) * fontsize * 0.5
                else:
                    text_width = len(text) * fontsize * 0.5
                x_end = x_start + text_width

            # Underline Y: use baseline origin[1] + a small percentage of fontsize
            # Gap of 10-12% of fontsize is standard
            underline_y = origin[1] + (fontsize * 0.12)
            
            # Determine thickness from original info if found, else default
            thickness = orig_info.get("thickness", max(0.8, fontsize / 16.0)) if orig_info.get("found") else max(0.8, fontsize / 16.0)
            u_color = orig_info.get("color", color) if orig_info.get("found") else color

            # Normalize color to tuple of floats
            if isinstance(u_color, (list, tuple)):
                u_color = tuple(float(c) for c in u_color)
            else:
                u_color = color

            start_point = fitz.Point(x_start, underline_y)
            end_point = fitz.Point(x_end, underline_y)

            # Draw the underline
            page.draw_line(start_point, end_point, color=u_color, width=thickness)
        except Exception:
            # Fallback: annotation-based underline
            try:
                if bbox is not None:
                    u_rect = fitz.Rect(bbox.x0, bbox.y1 - 2, bbox.x1, bbox.y1)
                else:
                    text_w = len(text) * fontsize * 0.5
                    u_rect = fitz.Rect(origin[0], origin[1] + 0.5,
                                       origin[0] + text_w, origin[1] + 2.5)
                annot = page.add_underline_annot(u_rect)
                annot.set_colors(stroke=color)
                annot.update()
            except Exception:
                pass

    # ─── OTHER PAGE OPERATIONS ───────────────────────────────────────────

    def rotate_page(self, page_num, angle=90):
        """Rotates the specified page clockwise."""
        if not self.doc:
            return False
        page = self.doc[page_num]
        page.set_rotation((page.rotation + angle) % 360)
        return True

    def delete_page(self, page_num):
        """Deletes the specified page."""
        if not self.doc or len(self.doc) <= 1:
            return False
        self.doc.delete_page(page_num)
        return True

    def duplicate_page(self, page_num):
        """Duplicates the specified page and inserts it after."""
        if not self.doc:
            return False
        self.doc.fullcopy_page(page_num, page_num + 1)
        return True

    def apply_overlays(self, all_overlays):
        """Bakes all session overlays into the PDF document."""
        if not self.doc or not all_overlays:
            return
            
        for page_num, overlays in all_overlays.items():
            if page_num >= len(self.doc): continue
            page = self.doc[page_num]
            
            for ov in overlays:
                bbox = fitz.Rect(ov["bbox"])
                rotation = ov.get("rotation", 0)
                
                if ov["type"] == "image":
                    if ov.get("preserve_aspect") and "aspect" in ov:
                        orig_aspect = ov["aspect"]
                        box_w = bbox.width
                        box_h = bbox.height
                        box_aspect = box_w / box_h if box_h != 0 else 1.0
                        
                        if orig_aspect > box_aspect:
                            # Image wider than box relatively
                            new_w = box_w
                            new_h = box_w / orig_aspect
                            offset_y = (box_h - new_h) / 2
                            bbox = fitz.Rect(bbox.x0, bbox.y0 + offset_y, bbox.x1, bbox.y0 + offset_y + new_h)
                        else:
                            # Image taller than box relatively
                            new_h = box_h
                            new_w = box_h * orig_aspect
                            offset_x = (box_w - new_w) / 2
                            bbox = fitz.Rect(bbox.x0 + offset_x, bbox.y0, bbox.x0 + offset_x + new_w, bbox.y1)
                    
                    page.insert_image(bbox, filename=ov["path"], rotate=rotation)
                elif ov["type"] == "text":
                    color = self._hex_to_rgb(ov.get("color", "#000000"))
                    fontsize = ov.get("fontsize", 12)
                    text = ov["text"]
                    is_bold = ov.get("bold", False)
                    
                    fontname = "hebo" if is_bold else "helv"
                    align = ov.get("align", 1)
                    
                    page.insert_textbox(bbox, text, 
                                        fontname=fontname,
                                        fontsize=fontsize, 
                                        color=color, 
                                        rotate=rotation,
                                        align=align)
        return True

    def save_pdf(self, path, page_overlays=None):
        """Saves the modified PDF, applying any pending overlays first.
        Handles overwriting the currently open file on Windows by using a temp file."""
        if not self.doc:
            return False
            
        if page_overlays:
            self.apply_overlays(page_overlays)
            
        try:
            target_abs = os.path.abspath(path)
            current_abs = os.path.abspath(self.doc.name) if self.doc.name else ""
            
            if target_abs == current_abs:
                # 1. Save to a temporary file first
                temp_path = path + ".tmp_save"
                self.doc.save(temp_path, incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP)
                
                # 2. Close the current document to release file lock
                self.doc.close()
                self.doc = None
                
                # 3. Replace the original file
                import shutil
                shutil.move(temp_path, path)
                
                # 4. Re-open the document
                self.doc = fitz.open(path)
            else:
                self.doc.save(path, incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP)
            return True
        except Exception:
            # Last resort fallback save if something went wrong with the move
            try:
                if self.doc and not self.doc.is_closed:
                    self.doc.save(path + ".failed_over.pdf", incremental=False)
            except: pass
            return False

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None

    def _hex_to_rgb(self, color_int):
        """Converts integer color to RGB tuple (0-1 range)."""
        if isinstance(color_int, str):
            if color_int.startswith("#"):
                color_int = int(color_int[1:], 16)
            else:
                color_int = int(color_int, 16)
                
        r = ((color_int >> 16) & 0xFF) / 255.0
        g = ((color_int >> 8) & 0xFF) / 255.0
        b = (color_int & 0xFF) / 255.0
        return (r, g, b)
