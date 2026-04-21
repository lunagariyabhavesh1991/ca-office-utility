"""
Base class and common data structures for modular bank statement parsers.
"""
import re
import os
from dataclasses import dataclass
from typing import List, Optional

@dataclass
class StatementTransaction:
    """Normalized bank transaction data structure."""
    sr_no: str = ""
    date: str = ""
    time: str = ""
    value_date: str = ""
    narration: str = ""
    ref_no: str = ""
    signed_amount: Optional[float] = None
    debit: float = 0.0
    credit: float = 0.0
    balance: Optional[float] = None
    confidence: float = 1.0

    def to_list(self, headers: List[str]) -> List[str]:
        """Map fields to a list based on provided headers."""
        row = []
        for h in headers:
            hl = h.lower()
            if hl in ["#", "sr no"]: row.append(self.sr_no)
            elif any(x in hl for x in ["txn date", "transaction date", "date"]): row.append(self.date)
            elif "time" in hl: row.append(self.time)
            elif "value date" in hl: row.append(self.value_date)
            elif any(x in hl for x in ["detail", "narration", "particular", "description"]): row.append(self.narration)
            elif any(x in hl for x in ["ref", "chq", "cheque"]): row.append(self.ref_no)
            elif "signed" in hl: row.append(f"{self.signed_amount:.2f}" if self.signed_amount is not None else "")
            elif "debit" in hl or "withdrawal" in hl: row.append(f"{self.debit:.2f}" if self.debit else "")
            elif "credit" in hl or "deposit" in hl: row.append(f"{self.credit:.2f}" if self.credit else "")
            elif "balance" in hl: row.append(f"{self.balance:.2f}" if self.balance is not None else "")
            else: row.append("")
        return row

class BaseBankParser:
    """Abstract base class for all bank-specific parsers."""
    
    def __init__(self, profile):
        self.profile = profile
        self.transactions = []
        self.custom_headers = []

    def parse(self, pdf_path: str, pages: List[int], debug_callback=None) -> List[StatementTransaction]:
        """Core parsing method to be implemented by subclasses."""
        raise NotImplementedError("Subclasses must implement parse()")

    @staticmethod
    def _parse_amount(text: str, strict: bool = False) -> Optional[float]:
        """Static helper for numeric parsing."""
        if not text: return None
        try:
            # Clean: remove symbols, commas, spaces
            clean = re.sub(r'[^\d.\-+]', '', text)
            if not clean or clean == "-" or clean == "+": return None
            
            # Handle Indian numbering system or standard
            val = float(clean)
            
            # Year rejection hardening (AU/Kotak specific)
            if strict and (2020 <= val <= 2030): return None
            
            return val
        except:
            return None

    @staticmethod
    def _parse_date(text: str, formats: List[str]) -> Optional[str]:
        """Static helper for date parsing."""
        from datetime import datetime
        if not text: return None
        text = text.strip()
        for fmt in formats:
            try:
                return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
            except:
                continue
        return None
