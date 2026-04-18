"""
Extraction Engine Service for CMA / DPR Builder.
Automated financial data extraction from PDF/Excel statements.
Converts extracted values to lakhs and detects financial years.
"""

import re
import logging
from typing import Dict, Any, List, Optional
import pypdf

logger = logging.getLogger(__name__)

class ExtractionEngineService:
    # Pattern mapping for common financial terms in Indian Banking
    PATTERNS = {
        "revenue": [
            r"revenue\s+from\s+operations", r"gross\s+turnover", r"total\s+revenue", r"sales", 
            r"total\s+income", r"income\s+from\s+services", r"receipts\s+from\s+customers",
            r"sales\s+accounts"
        ],
        "net_profit": [
            r"net\s+profit", r"pat", r"profit\s+after\s+tax", r"profit\s+for\s+the\s+period", 
            r"nett\s+profit", r"surplus\s+for\s+the\s+year"
        ],
        "depreciation": [r"depreciation", r"depr\s+&\s+amortization", r"depr\.", r"depreciation\s+and\s+amortization"],
        "interest_paid": [
            r"finance\s+costs", r"interest\s+on\s+borrowings", r"interest\s+expense", 
            r"interest\s+paid", r"interest\s+on\s+loans"
        ],
        "net_block": [
            r"fixed\s+assets", r"net\s+block", r"property,\s+plant\s+and\s+equipment", 
            r"total\s+fixed\s+assets", r"tangible\s+assets", r"net\s+fixed\s+assets"
        ],
        "investments": [r"investments", r"long\s+term\s+investments", r"total\s+investments"],
        "current_assets": [r"total\s+current\s+assets", r"current\s+assets"],
        "inventory": [r"inventories", r"stock", r"closing\s+stock", r"stock-in-trade", r"work-in-progress"],
        "debtors": [r"trade\s+receivables", r"sundry\s+debtors", r"debtors", r"receivables"],
        "deposits": [r"deposits\s+\(asset\)", r"security\s+deposits"],
        "loans_advances": [r"loans\s+&\s+advances\s+\(asset\)", r"short\s+term\s+loans\s+&\s+advances"],
        "cash_bank": [
            r"cash\s+and\s+bank\s+balances", r"cash\s+and\s+cash\s+equivalents", r"cash\s+at\s+bank",
            r"cash-in-hand", r"bank\s+accounts", r"cash\s+in\s+hand"
        ],
        "share_capital": [
            r"share\s+capital", r"owner'?s?\s+equity", r"proprietor'?s?\s+capital", r"partner'?s?\s+capital",
            r"capital\s+account", r"capital\s+a/c"
        ],
        "reserves_surplus": [r"reserves\s+and\s+surplus", r"retained\s+earnings", r"other\s+equity", r"surplus\s+in\s+statement", r"surplus"],
        "term_loans": [r"non-current\s+borrowings", r"term\s+loans", r"long\s+term\s+borrowings", r"secured\s+loans"],
        "unsecured_loan": [r"unsecured\s+loans", r"unsecured\s+borrowings", r"proprietor'?s?\s+loan", r"partner'?s?\s+loan"],
        "bank_od": [r"bank\s+od", r"overdraft", r"cash\s+credit", r"od\s+a/c"],
        "total_loan_hdr": [r"loan\s+\(liabilities\)", r"total\s+outside\s+borrowings"],
        "current_liabilities": [r"total\s+current\s+liabilities", r"current\s+liabilities"],
        "creditors": [
            r"trade\s+payables", r"sundry\s+creditors", r"creditors", 
            r"dues\s+to\s+micro\s+and\s+small\s+enterprises"
        ],
        "provisions": [r"provisions", r"provision\s+for\s+expenses"],
        "other_current_liabilities": [
            r"gst\s+payable", r"other\s+current\s+liabilities", r"duties\s+and\s+taxes", 
            r"duties\s+&\s+taxes", r"employee", r"other\s+payables", r"outstanding\s+liabilities"
        ],
        "other_current_assets": [
            r"other\s+current\s+assets", r"short\s+term\s+advances", 
            r"prepaid\s+expenses", r"prepaid", r"tcs\s+receivable", r"tds\s+receivable"
        ],
        
        # --- Granular P&L Patterns (Requirement 11) ---
        "opening_stock": [r"opening\s+stock", r"opening\s+inventory"],
        "cogs": [r"cost\s+of\s+goods\s+sold", r"cogs", r"purchase\s+accounts"],
        "gross_profit": [r"gross\s+profit", r"gp"],
        "salary_wages": [r"salary\s+&\s+wages", r"staff\s+costs", r"employee\s+benefit\s+expenses", r"salary\s+exp"],
        "labour_expenses": [r"labour\s+exp", r"labour\s+charges", r"wages", r"direct\s+wages"],
        "power_fuel": [r"power\s+&\s+fuel", r"electricity\s+charges", r"power\s+fuel"],
        "rent_rates": [r"rent\s+&\s+rates", r"rent\s+paid", r"rent\s+expense", r"shop\s+rent"],
        "admin_expenses": [
            r"admin\s+&\s+misc", r"administrative\s+expenses", r"office\s+exp", 
            r"bonus\s+exp", r"other\s+expenses"
        ],
        "other_direct_expenses": [r"other\s+direct\s+expenses", r"direct\s+manufacturing\s+exp"],
        "interest_exp": [r"interest\s+exp", r"interest\s+paid", r"finance\s+charges"],
        "tax_amt": [r"provision\s+for\s+tax", r"income\s+tax", r"tax\s+paid", r"current\s+tax"],
        
        # Header/Total Patterns for Balancing (Requirement 12)
        "total_indirect_exp": [r"total\s+indirect\s+expenses", r"indirect\s+expenses"]
    }

    @classmethod
    def _find_value_for_patterns(cls, lines: List[str], patterns: List[str], sum_mode: bool = False) -> float:
        """Finds numeric values matching patterns with Multi-column and Sum-mode Awareness."""
        total = 0.0
        found = False
        
        for line in lines:
            line_lower = line.lower()
            for pattern in patterns:
                match = re.search(pattern, line_lower)
                if match:
                    # Look for digits on the same line
                    # Updated regex to capture accounting negatives e.g. (-)1,234.00 or (1,234.00)
                    num_strs = re.findall(r"(?:\(-\)|[\(-])?[\d,]+(?:\.\d+)?[\)]?", line)
                    if not num_strs:
                        # Fallback for small numbers or zero (e.g. 0.00)
                        num_strs = re.findall(r"\d+\.\d+|\b\d+\b", line)
                        
                    if num_strs:
                        def parse_accounting_num(s):
                            s = s.strip().replace(",", "")
                            is_neg = False
                            # Detect (-) prefix or ( ) wrapping
                            if s.startswith("(-)") or (s.startswith("(") and s.endswith(")")) or s.startswith("-"):
                                is_neg = True
                                s = s.replace("(-)", "").replace("(", "").replace(")", "").replace("-", "")
                            
                            try:
                                return float(s) * (-1 if is_neg else 1)
                            except:
                                return 0.0

                        # --- Multi-column Logic ---
                        # If label is in the first 40% of the line, it's likely a left-column item
                        val = 0.0
                        if len(num_strs) > 1:
                            term_pos = match.start()
                            if term_pos < len(line) * 0.4:
                                val = parse_accounting_num(num_strs[0])
                            else:
                                val = parse_accounting_num(num_strs[-1])
                        else:
                            val = parse_accounting_num(num_strs[0])
                        
                        if sum_mode:
                            total += val
                            found = True
                            break # Move to next line for this term
                        else:
                            return val
        
        return total if found else 0.0

    @classmethod
    def extract_from_pdf(cls, file_path: str) -> dict:
        """Core extraction engine with layout-aware line parsing."""
        text = ""
        lines = []
        
        try:
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    # Layout=True preserves columns and spacing
                    page_text = page.extract_text(layout=True)
                    if page_text:
                        text += page_text + "\n"
                        lines.extend(page_text.split("\n"))
        except:
            # Fallback to standard pypdf
            try:
                reader = pypdf.PdfReader(file_path)
                for page in reader.pages:
                    t = page.extract_text()
                    if t: text += t + "\n"
                lines = text.split("\n")
            except Exception as e:
                return {"status": "error", "message": f"PDF Read Error: {str(e)}"}

        unit_multiplier = cls._detect_unit_multiplier(text)
        found_years = cls._detect_years(text)
        
        results = {}
        # Keys that should sum all matches (multi-line accounts)
        SUM_KEYS = [
            "cash_bank", "net_block", "inventory", "creditors", "debtors", 
            "other_current_liabilities", "other_current_assets", "investments",
            "provisions", "unsecured_loan", "bank_od"
        ]
        
        for key, patterns in cls.PATTERNS.items():
            sum_mode = key in SUM_KEYS
            val = cls._find_value_for_patterns(lines, patterns, sum_mode=sum_mode)
            # Standard Lakhs conversion
            if abs(val) > 0:
                results[key] = round((val * unit_multiplier) / 100000, 2)
            else:
                results[key] = 0.0

        # Balancing Figure for Indirect Expenses
        if results.get("total_indirect_exp", 0) > 0:
            total_ind = results["total_indirect_exp"]
            specific_exp = (
                results.get("salary_wages", 0) + 
                results.get("power_fuel", 0) + 
                results.get("rent_rates", 0) + 
                results.get("depreciation", 0)
            )
            balancing_admin = max(0, total_ind - specific_exp)
            results["admin_expenses"] = round(balancing_admin, 2)
            
        # Balancing for Other Current Assets (Requirement 13)
        if results.get("current_assets", 0) > 0:
            total_ca = results["current_assets"]
            found_ca = (
                results.get("inventory", 0) + 
                results.get("debtors", 0) + 
                results.get("cash_bank", 0) +
                results.get("loans_advances", 0) +
                results.get("deposits", 0)
            )
            # Residual goes to Other Current Assets
            residual_oca = round(total_ca - found_ca, 2)
            if abs(residual_oca) > abs(results.get("other_current_assets", 0)):
                results["other_current_assets"] = residual_oca
                
        # Balancing for Other Current Liabilities (Requirement 13)
        if results.get("current_liabilities", 0) > 0:
            total_cl = results["current_liabilities"]
            found_cl = (
                results.get("creditors", 0) +
                results.get("provisions", 0)
            )
            # Residual goes to Other Current Liabilities
            residual_ocl = round(total_cl - found_cl, 2)
            if abs(residual_ocl) > abs(results.get("other_current_liabilities", 0)):
                results["other_current_liabilities"] = residual_ocl

        # Balancing for Other Loans (Requirement 13)
        if results.get("total_loan_hdr", 0) > 0:
            total_loans = results["total_loan_hdr"]
            found_loans = (
                results.get("term_loans", 0) +
                results.get("unsecured_loan", 0) +
                results.get("bank_od", 0)
            )
            # Residual goes to Other Loans & Liabilities
            residual_ol = round(total_loans - found_loans, 2)
            if abs(residual_ol) > abs(results.get("other_loans_liabilities", 0)):
                results["other_loans_liabilities"] = residual_ol
        
        return {
            "status": "success",
            "data": results,
            "detected_years": found_years,
            "unit_multiplier": unit_multiplier,
            "units_description": "Lakhs" if unit_multiplier == 100000 else "Crores" if unit_multiplier == 10000000 else "INR"
        }

    @staticmethod
    def _detect_unit_multiplier(text: str) -> float:
        """Detects if figures are in Rupees, Lakhs, or Crores."""
        text_lower = text.lower()
        if "in lakhs" in text_lower or "(lakhs)" in text_lower:
            return 100000.0
        if "in crores" in text_lower or "(crores)" in text_lower:
            return 10000000.0
        if "in thousands" in text_lower or "('000)" in text_lower:
            return 1000.0
        return 1.0

    @staticmethod
    def _detect_years(text: str) -> List[str]:
        """Finds dates like 31-03-2024 or 2023-24."""
        years = re.findall(r"20\d{2}", text)
        return sorted(list(set(years)), reverse=True)

