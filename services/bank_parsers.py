"""
Specialized parser implementations for specific banks.
"""
import os
import pdfplumber
from typing import List
from services.bank_parser_base import BaseBankParser, StatementTransaction

class KotakParser(BaseBankParser):
    """Refined Coordinate-Based Parser for Kotak Mahindra Bank."""
    
    def parse(self, pdf_path: str, pages: List[int], debug_callback=None) -> List[StatementTransaction]:
        self.custom_headers = ["Sr No", "Date", "Time", "Value Date", "Narration", "Ref / Chq No", "Signed Amount", "Debit", "Credit", "Balance"]
        
        with pdfplumber.open(pdf_path) as pdf:
            current_txn = None
            
            # X-Ranges for Column Mapping (Kotak 7 Columns)
            REGIONS = {
                "sr_no": (0, 45),
                "date": (45, 115),
                "val_date": (115, 185),
                "narration": (185, 395),
                "ref": (395, 505),
                "amt_signed": (505, 615),
                "balance": (615, 750)
            }

            for page_idx in pages:
                if page_idx >= len(pdf.pages): continue
                page = pdf.pages[page_idx]
                words = page.extract_words()
                if not words: continue
                
                # Group words into Lines (Vertical tolerance: 3pts)
                lines = []
                words.sort(key=lambda x: (x['top'], x['x0']))
                current_line = [words[0]]
                for w in words[1:]:
                    if abs(w['top'] - current_line[0]['top']) < 3:
                        current_line.append(w)
                    else:
                        lines.append(current_line)
                        current_line = [w]
                lines.append(current_line)

                for line in lines:
                    col_data = {k: [] for k in REGIONS.keys()}
                    for w in line:
                        x_mid = (w['x0'] + w['x1']) / 2
                        for col_name, (x_start, x_end) in REGIONS.items():
                            if x_start <= x_mid < x_end:
                                col_data[col_name].append(w['text'])
                                break
                    
                    cells = {k: " ".join(v).strip() for k, v in col_data.items()}
                    
                    # --- NEW ROW DETECTION ---
                    sr_no = cells["sr_no"]
                    date_raw = cells["date"]
                    parsed_date = self._parse_date(date_raw, self.profile["date_formats"])
                    
                    if sr_no.isdigit() and parsed_date:
                        if current_txn: self.transactions.append(current_txn)
                        current_txn = StatementTransaction(sr_no=sr_no, date=parsed_date)
                        current_txn.value_date = cells["val_date"]
                        current_txn.narration = cells["narration"]
                        current_txn.ref_no = cells["ref"]
                        
                        raw_amt = self._parse_amount(cells["amt_signed"])
                        if raw_amt is not None:
                            current_txn.signed_amount = raw_amt
                            if raw_amt < 0: current_txn.debit = abs(raw_amt)
                            else: current_txn.credit = raw_amt
                        
                        current_txn.balance = self._parse_amount(cells["balance"])
                            
                    elif current_txn:
                        # Wrapped Data (Time in Date col, or wrapped Narration/Ref)
                        if any(x in date_raw for x in [":", "AM", "PM"]):
                            current_txn.time = date_raw
                        if cells["narration"]:
                            current_txn.narration += " " + cells["narration"]
                        if cells["ref"]:
                            current_txn.ref_no += " " + cells["ref"]

            if current_txn: self.transactions.append(current_txn)
        return self.transactions

class AUParser(BaseBankParser):
    """Refined Coordinate-Based Parser for AU Small Finance Bank."""

    def parse(self, pdf_path: str, pages: List[int], debug_callback=None) -> List[StatementTransaction]:
        self.custom_headers = ["Trans Date", "Value Date", "Description / Narration", "Chq / Ref No", "Debit", "Credit", "Balance"]
        
        with pdfplumber.open(pdf_path) as pdf:
            current_txn = None
            
            REGIONS = {
                "trans_date": (0, 75),
                "val_date": (75, 150),
                "narration": (150, 380),
                "ref": (380, 480),
                "debit": (480, 545),
                "credit": (545, 610),
                "balance": (610, 800)
            }

            for page_idx in pages:
                if page_idx >= len(pdf.pages): continue
                page = pdf.pages[page_idx]
                words = page.extract_words()
                if not words: continue
                
                # Group words into Lines
                lines = []
                words.sort(key=lambda x: (x['top'], x['x0']))
                current_line = [words[0]]
                for w in words[1:]:
                    if abs(w['top'] - current_line[0]['top']) < 3:
                        current_line.append(w)
                    else:
                        lines.append(current_line)
                        current_line = [w]
                lines.append(current_line)

                for line in lines:
                    col_data = {k: [] for k in REGIONS.keys()}
                    for w in line:
                        x_mid = (w['x0'] + w['x1']) / 2
                        for col_name, (x_start, x_end) in REGIONS.items():
                            if x_start <= x_mid < x_end:
                                col_data[col_name].append(w['text'])
                                break
                    
                    cells = {k: " ".join(v).strip() for k, v in col_data.items()}
                    if not any(cells.values()): continue
                    
                    trans_date_raw = cells["trans_date"]
                    parsed_date = self._parse_date(trans_date_raw, self.profile["date_formats"])
                    
                    if parsed_date:
                        if current_txn: self.transactions.append(current_txn)
                        current_txn = StatementTransaction(date=parsed_date)
                        current_txn.value_date = cells["val_date"]
                        current_txn.narration = cells["narration"]
                        current_txn.ref_no = cells["ref"]
                        current_txn.debit = self._parse_amount(cells["debit"], strict=True) or 0.0
                        current_txn.credit = self._parse_amount(cells["credit"], strict=True) or 0.0
                        current_txn.balance = self._parse_amount(cells["balance"], strict=True)
                            
                    elif current_txn:
                        if cells["narration"]:
                            current_txn.narration += " " + cells["narration"]
                        if cells["ref"]:
                             current_txn.ref_no += " " + cells["ref"]

            if current_txn: self.transactions.append(current_txn)
        return self.transactions

class ICICIParser(BaseBankParser):
    """Robust parser for ICICI Bank with multi-line narration merging."""
    
    def parse(self, pdf_path: str, pages: List[int], debug_callback=None) -> List[StatementTransaction]:
        from services.pdf_table_extractor import PDFTableExtractor
        current_txn = None
        
        for page_idx in pages:
            extraction = PDFTableExtractor.extract(pdf_path, mode="accurate", pages=[page_idx])
            if not extraction.rows: continue
            
            # Map headers (ICICI: Date, Particulars, Chq/Ref, Withdrawals, Deposits, Balance)
            headers = [h.lower() for h in extraction.headers] if extraction.headers else []
            
            for row in extraction.rows:
                cells = [str(c).strip() if c else "" for c in row]
                if not any(cells): continue
                
                # Check for date in first or second column
                date_str = self._parse_date(cells[0], self.profile["date_formats"])
                if not date_str and len(cells) > 1:
                    date_str = self._parse_date(cells[1], self.profile["date_formats"])
                
                if date_str:
                    if current_txn: self.transactions.append(current_txn)
                    
                    # New Txn
                    current_txn = StatementTransaction(date=date_str)
                    
                    # Map other columns (rough based on typical ICICI layout)
                    if len(cells) >= 6:
                        # 0:Date, 1:ValueDate(maybe), 2:Particulars, 3:Ref, 4:Dr, 5:Cr, 6:Bal
                        # But we use a simpler fallback if headers are missing
                        current_txn.narration = cells[2] if len(cells) > 2 else ""
                        current_txn.ref_no = cells[3] if len(cells) > 3 else ""
                        current_txn.debit = self._parse_amount(cells[-3]) or 0.0
                        current_txn.credit = self._parse_amount(cells[-2]) or 0.0
                        current_txn.balance = self._parse_amount(cells[-1])
                    else:
                        # Minimal
                        current_txn.narration = " ".join(cells[1:-1])
                        current_txn.balance = self._parse_amount(cells[-1])
                elif current_txn:
                    # Merge narration
                    text = " ".join(c for c in cells if c and not self._parse_amount(c, strict=True)).strip()
                    if text and len(text) > 2:
                        current_txn.narration += " " + text
                        
        if current_txn: self.transactions.append(current_txn)
        return self.transactions

class GenericBankParser(BaseBankParser):
    """Fallback Parser for unsupported banks using dynamic column detection."""
    
    def parse(self, pdf_path: str, pages: List[int], debug_callback=None) -> List[StatementTransaction]:
        import pdfplumber
        
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx in pages:
                if page_idx >= len(pdf.pages): continue
                page = pdf.pages[page_idx]
                
                # 1. Extract Words
                words = page.extract_words()
                if not words: continue
                
                # 2. Group into Lines
                words.sort(key=lambda x: (x['top'], x['x0']))
                lines = []
                if words:
                    current_line = [words[0]]
                    for w in words[1:]:
                        if abs(w['top'] - current_line[0]['top']) < 3:
                            current_line.append(w)
                        else:
                            lines.append(current_line)
                            current_line = [w]
                    lines.append(current_line)
                
                # 3. Detect Column Centers (Dynamic)
                col_centers = {} # mapping of "date", "narr", "ref", "dr", "cr", "bal" -> center x
                
                for line in lines:
                    line_text = " ".join(w['text'] for w in line).lower()
                    if "date" in line_text and ("balance" in line_text or "amount" in line_text):
                        # Found header row!
                        for w in line:
                            txt = w['text'].lower()
                            mid = (w['x0'] + w['x1']) / 2
                            if "date" in txt:
                                if "date" not in col_centers: col_centers["date"] = mid
                            elif any(x in txt for x in ["particular", "narrat", "descript", "detail"]):
                                col_centers["narr"] = mid
                            elif any(x in txt for x in ["ref", "chq", "instr"]):
                                col_centers["ref"] = mid
                            elif any(x in txt for x in ["withdrawal", "debit", "dr"]):
                                col_centers["dr"] = mid
                            elif any(x in txt for x in ["deposit", "credit", "cr"]):
                                col_centers["cr"] = mid
                            elif "balance" in txt:
                                col_centers["bal"] = mid
                        break
                
                # 4. Process Rows
                current_txn = None
                for line in lines:
                    # Try to find a date in the line
                    line_words = sorted(line, key=lambda x: x['x0'])
                    
                    found_date = None
                    date_word_idx = -1
                    
                    for i, w in enumerate(line_words):
                        parsed = self._parse_date(w['text'], self.profile["date_formats"])
                        if parsed:
                            found_date = parsed
                            date_word_idx = i
                            break
                    
                    if found_date:
                        if current_txn: self.transactions.append(current_txn)
                        current_txn = StatementTransaction(date=found_date)
                        
                        # Map rest of words to columns based on distance to col_centers
                        for i, w in enumerate(line_words):
                            if i == date_word_idx: continue
                            
                            mid = (w['x0'] + w['x1']) / 2
                            txt = w['text']
                            
                            # Assign to closest column center
                            if not col_centers:
                                # Fallback if no headers found: use relative position
                                # (Very rough: Date... Narr... Dr... Cr... Bal)
                                rel = mid / page.width
                                if rel < 0.2: pass # likely date
                                elif rel < 0.6: current_txn.narration += " " + txt
                                elif rel < 0.75: current_txn.debit = self._parse_amount(txt) or current_txn.debit
                                elif rel < 0.88: current_txn.credit = self._parse_amount(txt) or current_txn.credit
                                else: current_txn.balance = self._parse_amount(txt) or current_txn.balance
                            else:
                                best_col = min(col_centers.keys(), key=lambda k: abs(col_centers[k] - mid))
                                dist = abs(col_centers[best_col] - mid)
                                
                                # Only assign if reasonably close (e.g. within 100pts)
                                if dist < 150:
                                    if best_col == "narr": current_txn.narration += " " + txt
                                    elif best_col == "ref": current_txn.ref_no += " " + txt
                                    elif best_col == "dr": current_txn.debit = self._parse_amount(txt) or current_txn.debit
                                    elif best_col == "cr": current_txn.credit = self._parse_amount(txt) or current_txn.credit
                                    elif best_col == "bal": current_txn.balance = self._parse_amount(txt) or current_txn.balance
                        
                        current_txn.narration = current_txn.narration.strip()
                        current_txn.ref_no = current_txn.ref_no.strip()
                        
                        # Validate: Must have at least one numeric field
                        if not any([current_txn.debit, current_txn.credit, current_txn.balance is not None]):
                            current_txn = None
                            continue
                        
                    elif current_txn:
                        # Multi-line narration? Map by centers
                        for w in line:
                            mid = (w['x0'] + w['x1']) / 2
                            if col_centers:
                                best_col = min(col_centers.keys(), key=lambda k: abs(col_centers[k] - mid))
                                if best_col == "narr": current_txn.narration += " " + w['text']
                                elif best_col == "ref" and abs(col_centers["ref"] - mid) < 50: 
                                    current_txn.ref_no += " " + w['text']
                            else:
                                # Loose fallback
                                if 0.2 < (mid / page.width) < 0.6:
                                    current_txn.narration += " " + w['text']
                        
                if current_txn: self.transactions.append(current_txn)
                
        return self.transactions
