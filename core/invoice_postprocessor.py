import re
from datetime import datetime

class InvoicePostprocessor:
    """
    Cleans and normalizes extracted raw data into standard display formats.
    """
    
    @staticmethod
    def normalize_text(text: str) -> str:
        """
        Unifies spaces, fixes broken line breaks from OCR, 
        and prepares strict text for regex processing.
        """
        if not text:
            return ""
        
        # Keep multiple spaces as they are important for column splitting!
        # Only normalize tabs and other whitespace characters to spaces
        text = re.sub(r'[ \t]', ' ', text)
        
        # Remove empty lines
        text = re.sub(r'\n\s*\n', '\n', text)
        
        # Vertical Rejoining for numeric fragments (e.g., "16\n,\n500")
        text = re.sub(r'(\d)\s*\n\s*([,\.])\s*', r'\1\2', text)
        text = re.sub(r'([,\.])\s*\n\s*(\d)', r'\1\2', text)
        
        # Rejoin broken numbers at commas and dots (e.g., "16 , 500" -> "16,500")
        text = re.sub(r'(\d)\s*,\s*(\d)', r'\1,\2', text)
        text = re.sub(r'(\d)\s*,\s*(\d)', r'\1,\2', text) # Double pass for overlapping
        
        # Rejoin broken decimals (e.g., "16,500 . 00" -> "16,500.00")
        # Constrained to 2 spaces to avoid bridging across columns in grids (Sale 95/103)
        text = re.sub(r'(\d)\s{0,2}\.\s{0,2}(\d{1,2})\b', r'\1.\2', text)
        
        return text.strip()

    @staticmethod
    def normalize_date(raw_date: str) -> str:
        """
        Attempts to convert various date strings into strict DD/MM/YYYY format.
        """
        if not raw_date:
            return ""
            
        clean_date = raw_date.strip().upper()
        # Remove common OCR artifacts attached to dates
        clean_date = re.sub(r'[^\w\-\/\.]', '', clean_date)
        
        formats = [
            "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", 
            "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
            "%d-%b-%Y", "%d %b %Y", "%d-%b-%y",
            "%Y-%m-%d", "%Y/%m/%d"
        ]
        
        for fmt in formats:
            try:
                dt = datetime.strptime(clean_date, fmt)
                return dt.strftime("%d/%m/%Y")
            except ValueError:
                continue
                
        # If it couldn't parse, return the cleanest version we got
        return clean_date

    @staticmethod
    def clean_amount(raw_amount: str) -> str:
        """
        Strips currency symbols (Rs, ₹, $, INR) and cleans numeric strings.
        Returns a clean float-compatible string like '1234.50'.
        """
        if not raw_amount:
            return ""
            
        # Keep only digits, decimals, and commas
        clean_amt = re.sub(r'[^\d\.,]', '', str(raw_amount))
        
        # Handle European formats if needed, or strictly remove commas
        # Assuming Indian/US format where comma is thousands separator
        clean_amt = clean_amt.replace(',', '')
        
        try:
            # Ensure it's a valid float
            val = float(clean_amt)
            return f"{val:.2f}"
        except ValueError:
            return "" # Invalid amount

    @staticmethod
    def clean_invoice_number(raw_num: str) -> str:
        """
        Removes weird OCR artifacts from invoice numbers.
        """
        if not raw_num:
            return ""
        
        clean = raw_num.strip()
        # Remove trailing noise common in multi-column layouts
        # Split by multiple spaces or common label starts that bleed in
        clean = re.split(r'(?i)(\s{2,}|DATE|PLACE|SHIP|BILL|GSTIN|PLAC|DAT|VEHIC|ORDER|CHALLAN)', clean)[0].strip()
        # Remove leading/trailing non-alphanumeric (like colons or dashes)
        clean = re.sub(r'^[^A-Za-z0-9]+|[^A-Za-z0-9]+$', '', clean)
        return clean

    @staticmethod
    def clean_party_name(raw_name: str) -> str:
        """
        Cleans vendor names by removing prefixes or weird characters.
        """
        if not raw_name:
            return ""
            
        clean = raw_name.strip()
        # Remove "M/s" or "To:" prefixes
        clean = re.sub(r'^(M/s\.?|To:?|From:?|Supplier:?)\s*', '', clean, flags=re.IGNORECASE)
        
        # Deduplicate (common in multi-column FITZ extractions)
        parts = re.split(r'(?i)\s+(?:M/S\.?|NAME|BUYER|PARTY|RECIPIENT)\s*[:.-]*\s*', clean)
        if len(parts) > 1 and parts[1].strip():
            p1 = parts[0].strip().upper()
            p2 = parts[1].strip().upper()
            if p2 in p1 or p1 in p2:
                clean = parts[0].strip()
        
        # Aggressively remove trailing label fragments ("Madhav Ayurvedic Name" -> "Madhav Ayurvedic")
        clean = re.sub(r'(?i)\s+(NAME|PARTY|RECIPIENT|BUYER|SUPPLIER)$', '', clean).strip()
                
        # Remove weird symbols at ends
        clean = re.sub(r'^[^A-Za-z0-9]+|[^A-Za-z0-9]+$', '', clean)
        return clean.title()
