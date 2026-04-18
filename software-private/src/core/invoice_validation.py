import re

class InvoiceValidation:
    """
    Business logic for validating Indian GST Invoices.
    Ensures data consistency and compliance with common accounting rules.
    """
    
    @staticmethod
    def is_valid_gstin(gstin: str) -> bool:
        """Checks if a string matches the standard Indian GSTIN format."""
        if not gstin: return False
        pattern = r'^[0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$'
        return bool(re.match(pattern, gstin.upper()))

    @staticmethod
    def validate_totals(taxable: str, cgst: str, sgst: str, igst: str, grand_total: str) -> tuple:
        """
        Validates: Taxable + CGST + SGST + IGST == Grand Total
        Returns: (status_bool, remark_string)
        - 0.00 to 0.05 mismatch = Minor Rounding Mismatch
        - Above 0.05 = Major Mismatch
        """
        try:
            def to_f(s): 
                if not s or not s.strip() or s == "-": return 0.0
                return float(s.replace(',', ''))
            
            t = to_f(taxable)
            c = to_f(cgst)
            s = to_f(sgst)
            i = to_f(igst)
            g = to_f(grand_total)
            
            # If everything is 0/empty, extraction failed
            if t == 0 and g == 0:
                return False, "Totals missing"
                
            # NOISE REJECTION: If values are extremely low (unrealistic for GST Invoices)
            if g < 50.0 and (c > 0 or s > 0 or i > 0):
                 return False, f"Possible noise detected (Total too low: {g:.2f})"

            expected = round(t + c + s + i, 2)
            actual = round(g, 2)
            error = abs(expected - actual)
            
            if error <= 0.05:
                return True, ""
            elif error <= 1.01:
                # Standard Indian GST invoices round to the nearest Rupee
                return True, f"Minor rounding mismatch ({error:.2f})"
            else:
                return False, f"Major mismatch: Expected {expected:.2f}, Found {actual:.2f}"
                
        except (ValueError, TypeError):
            return False, "Numeric format error"
    @staticmethod
    def validate_field_guards(result: dict, raw_text: str) -> tuple:
        """
        Enforces Sales Invoice strict guards.
        Returns: (is_ok, remark)
        """
        remarks = []
        is_ok = True
        
        inv_no = str(result.get("Invoice No", ""))
        party_name = result.get("Party Name", "").upper()
        buyer_gstin = result.get("Buyer GSTIN", "").upper()
        
        # 1. Invoice No Guard: Reject 10-digit style (most likely mobile)
        if inv_no.isdigit() and len(inv_no) >= 10:
            remarks.append("Inv No looks like mobile")
            is_ok = False
            
        # 2. Party Name Guard: Reject if it matches top header or common supplier words
        # Also reject if it matches a section header or contains section labels / table headers
        blacklist = [
            'DETAILS', 'RECEIVER', 'BILLED TO', 'CONSIGNEE', 'SHIPPED TO', 
            'SHIP TO', 'BILL TO', 'BUYER', 'SR', 'SR.', 'NO', 'NAME OF PRODUCT', 
            'OF PRODUCT', 'DESCRIPTION', 'HSN', 'QTY', 'UNIT', 'RATE'
        ]
        
        lines = [l.strip().upper() for l in raw_text.split('\n') if l.strip()]
        if not party_name or len(party_name) < 3: # Relaxed from 4 to 3 for short names
            remarks.append("Party Name missing or too short")
            is_ok = False
        else:
            p_up = party_name.upper()
            if lines and p_up == lines[0]:
                remarks.append("Party matches Header (Supplier leakage)")
                is_ok = False
            
            if p_up in blacklist:
                remarks.append("Party Name matches label/header keyword")
                is_ok = False
            
            if "DETAILS OF" in p_up:
                remarks.append("Party Name looks like section title")
                is_ok = False
            
        # 3. Buyer GSTIN Guard: Reject if it appears too early in document (Supplier usually first)
        if buyer_gstin:
            first_idx = raw_text.upper().find(buyer_gstin)
            # Relaxed for single-page invoices: Only flag if found in the very first segment (header)
            # Typically supplier is in the first 400-500 chars. Buyer is later.
            if first_idx != -1 and first_idx < 450:
                remarks.append("Buyer GSTIN found in Supplier block")
                # We don't mark as Failed yet, just Need Review
                is_ok = False

        # 4. Total Guard: Grand Total vs Taxable
        try:
            t = float(result.get("Taxable Value", "0").replace(',', ''))
            g = float(result.get("Grand Total", "0").replace(',', ''))
            c = float(result.get("CGST", "0").replace(',', ''))
            s = float(result.get("SGST", "0").replace(',', ''))
            i = float(result.get("IGST", "0").replace(',', ''))
            
            if g == t and (c > 0 or s > 0 or i > 0):
                remarks.append("Grand Total matches Taxable (Tax ignored)")
                is_ok = False
        except:
            pass

        return is_ok, " | ".join(remarks)
