import re
from typing import Dict, List, Tuple

class InvoiceFieldExtractors:
    @staticmethod
    def _clean_num(s: str, reject_val: float = None) -> float:
        if not s: return 0.0
        # Rejection of HSN-like or PIN-like codes (4, 6, 8 digits without decimal)
        s_raw = s.strip().replace(' ', '').replace(',', '')
        if s_raw.isdigit() and len(s_raw) in [4, 6, 8]:
             # Rejection of纯数字代码 (Like Invoice No 0238 or HSN 3506)
             if not (s_raw.endswith('00') or s_raw.endswith('50')): return 0.0
             
        s_clean = re.sub(r'[^\d\.\,]', '', s).replace(',', '')
        try:
            if s_clean.isdigit() and len(s_clean) >= 10: return 0.0
            val = float(s_clean)
            if reject_val is not None and abs(val - reject_val) < 0.01: return 0.0
            # Rejection of years (Standalone integers 2000-2100 with no decimal part)
            if 2000 <= val <= 2100 and '.' not in s and ',' not in s: 
                return 0.0
            return val
        except:
            return 0.0

    @staticmethod
    def _get_sections(text: str) -> Dict[str, str]:
        sections = {"header": "", "bill_to": "", "details": "", "totals": ""}
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if not lines: return sections
        sections["header"] = "\n".join(lines[:20])
        bill_to_start = -1
        for i, line in enumerate(lines):
            if re.search(r'(?i)(DETAILS\s*OF\s*RECEIVER|BILL\s*TO|BUYER\s*TO\s*PARTY|BUYER|BILLED\s*TO|RECEIVER|PARTY\s*DETAILS|CUSTOMER|M/S\.?)', line.upper()):
                bill_to_start = i; break
        table_start = -1
        for i, line in enumerate(lines):
            if any(x in line.upper() for x in ['SR. NO', 'NAME OF PRODUCT', 'DESCRIPTION', 'HSN', 'PARTICULARS', 'ITEM NAME']):
                table_start = i; break
        if bill_to_start != -1:
            end = table_start if table_start != -1 else min(bill_to_start + 20, len(lines))
            sections["bill_to"] = "\n".join(lines[bill_to_start:end])
            sections["details"] = "\n".join(lines[max(0, bill_to_start-5):min(len(lines), end+10)])
        # Totals: scan bottom-up
        totals_start = -1
        for i in range(len(lines)-1, -1, -1):
            if 'RATE WISE SUMMARY' in lines[i].upper(): totals_start = i; break
        if totals_start == -1:
            for i in range(len(lines)-1, -1, -1):
                l_up = lines[i].upper()
                if any(x in l_up for x in ['GRAND TOTAL', 'TOTAL AMOUNT', 'NET AMOUNT', 'AMOUNT PAYABLE', 'TOTAL', 'SUB TOTAL', 'TAXABLE AMOUNT', 'AMOUNT WITH TAX']):
                    # Ensure it is not an item row with HSN or Sr No
                    if any(h in l_up for h in ['HSN', 'SAC', 'QTY', 'UNIT', 'RATE', 'SR.', 'NO.', 'S.R.', 'ITEM & DESCRIPTION']): continue
                    # Rejection of rows that look like invoice headers (e.g. Invoice No)
                    if any(h in l_up for h in ['INVOICE NO', 'DATE', 'GSTIN']): continue
                    totals_start = i
                    for j in range(i-1, max(0, i-60), -1):
                        if any(x in lines[j].upper() for x in ['SR. NO', 'DESCRIPTION OF GOODS', 'NAME OF PRODUCT', 'HSN/ SAC', 'ITEM NAME', 'ITEM & DESCRIPTION']): break
                        totals_start = j
                    break


        if totals_start != -1:
            raw = lines[totals_start:]
            clean_t = []
            for line in raw:
                if any(x in line.upper() for x in ['ACK NO', 'IRN :', 'ACKNOWLEDGEMENT', 'E-INVOICING DETAIL']): break
                clean_t.append(line)
            sections["totals"] = "\n".join(clean_t)
        return sections

    @staticmethod
    def _extract_label_value_pairs(text: str, reject_id: float = None) -> Dict[str, str]:
        """Extract amounts from explicit 'Label\\n:\\n₹Value' patterns (Anjani/Madhukar style)."""
        res = {"taxable": "", "cgst": "", "sgst": "", "igst": "", "grand_total": "", "round_off": "", "discount": ""}
        lines = [l.strip() for l in text.split('\n')]

        
        label_map = {
            'taxable': ['TOTAL TAXABLE AMOUNT', 'TOTAL AMOUNT BEFORE TAX', 'TAXABLE VALUE', 'SUB TOTAL', 'NET TAXABLE', 'BASIC AMOUNT'],
            'cgst': ['ADD : CGST', 'ADD :CGST', 'CGST AMOUNT', 'CGST@', 'CGST @', 'CENTRAL TAX', 'CGST', 'CENTRAL'],
            'sgst': ['ADD : SGST', 'ADD :SGST', 'SGST AMOUNT', 'SGST@', 'SGST @', 'STATE TAX', 'SGST', 'STATE/UT', 'STATE / UT'],

            'igst': ['ADD : IGST', 'ADD :IGST', 'IGST AMOUNT', 'IGST@', 'IGST @', 'INTEGRATED TAX', 'IGST'],
            'grand_total': ['AMOUNT WITH TAX', 'TOTAL AMOUNT', 'GRAND TOTAL', 'NET AMOUNT', 'AMOUNT PAYABLE', 'G.TOTAL AMOUNT', 'G.TOTAL', 'NET PAYABLE', 'TOTAL PAYABLE', 'TOTAL'],
            'round_off': ['ROUND OFF', 'ROUNDOFF', 'ROUNDING OFF'],
            'discount': ['DISC.AMT.', 'DISC.AMT', 'DISCOUNT', 'LESS DISC']
        }

        
        for i, line in enumerate(lines):
            l_up = line.upper().strip()
            # If line contains Invoice/Date, do NOT extract amounts from it
            if any(h in l_up for h in ['INVOICE NO', 'BILL NO', 'DATE']): continue
            
            for field, labels in label_map.items():
                for lbl in labels:
                    if lbl in l_up:
                        # REJECT if line contains Shipping/Freight/HSN noise
                        if any(x in l_up for x in ['SHIPPING', 'FREIGHT', 'POSTAGE', 'CHARGE', 'HSN']): 
                            if field in ['cgst', 'sgst', 'igst']: continue
                        
                        # Check for value on same line after : or ₹ or - or large space gap
                        m = re.search(r'(?:[:\₹\-]|(?:\s{2,}))\s*([\d,]+(?:\.\d{1,2})?)', line.split(lbl)[-1])

                        # For 'TOTAL', ensure it's not actually 'Total Taxable Amount'
                        if lbl == 'TOTAL' and ('TAXABLE' in l_up or 'BASIC' in l_up): continue

                        if m:

                            val = InvoiceFieldExtractors._clean_num(m.group(1), reject_id)

                            if val >= 0 or (field in ['cgst', 'sgst', 'igst', 'round_off']):
                                if not res[field] or val > float(res[field] or 0):
                                    res[field] = f"{val:.2f}"
                                break

                        # Next 1-2 lines
                        for k in range(1, 3):
                            if i + k < len(lines):
                                next_line = lines[i + k].strip()
                                m2 = re.search(r'[₹]?\s*([\d,]+(?:\.\d{1,2})?)', next_line)
                                if m2:
                                    val = InvoiceFieldExtractors._clean_num(m2.group(1), reject_id)
                                    if val >= 0:
                                        if not res[field] or val > float(res[field] or 0):
                                            res[field] = f"{val:.2f}"
                                        break
        return res

    @staticmethod
    def _extract_tax_summary_grid(text: str, reject_id: float = None) -> Dict[str, str]:
        """Extract from 'Tax Summary' grids (Vyapar/Shreeji style)."""
        res = {"taxable": "", "cgst": "", "sgst": "", "igst": "", "grand_total": ""}
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        
        summary_start = -1
        for i, line in enumerate(lines):
            if re.search(r'(?i)TAX\s*SUMMARY', line): summary_start = i; break
        if summary_start == -1: return res
        
        header_text = " ".join(lines[summary_start:min(summary_start+8, len(lines))]).upper()
        has_igst = 'IGST' in header_text and 'CGST' not in header_text
        has_cgst = 'CGST' in header_text
        has_sgst = 'SGST' in header_text
        
        # Grand Total from summary area
        for i in range(summary_start, min(summary_start + 15, len(lines))):
            l_up = lines[i].upper()
            if 'SUB TOTAL' in l_up or 'GRAND TOTAL' in l_up:
                m = re.search(r'(?:SUB\s*TOTAL|GRAND\s*TOTAL)\s*[:\s]*[₹]?\s*([\d,]+(?:\.\d{1,2})?)', l_up)
                if m:
                    val = InvoiceFieldExtractors._clean_num(m.group(1), reject_id)
                    if val > 0: res["grand_total"] = f"{val:.2f}"; break
                if i + 1 < len(lines):
                    m3 = re.search(r'[₹]?\s*([\d,]+(?:\.\d{1,2})?)', lines[i+1])
                    if m3:
                        val = InvoiceFieldExtractors._clean_num(m3.group(1), reject_id)
                        if val > 0: res["grand_total"] = f"{val:.2f}"; break
        
        # Merged Block: Join fragmented grid lines locally for vertical layout support (Sale 95/103)
        # We use a large 4-space gap to ensure re.findall never bridges across columns.
        grid_block = "    ".join(lines[summary_start:min(summary_start + 20, len(lines))])
        
        # Extract from the whole grid block at once to handle vertical splitting
        # Using strict continuous decimal pattern to prevent bridging
        amounts = re.findall(r'(?<![\d\.])[₹]?\s*(\d[\d,]*(?:\.\d{1,2})?)\b(?!%)', grid_block)
        
        valid = []
        for a in amounts:
            v = InvoiceFieldExtractors._clean_num(a, reject_id)
            if v > 2.0: valid.append(v)
            
        if len(valid) >= 2:
            # Use the same Rate-Match and Consensus heuristics on the whole grid block
            gt = float(res.get("grand_total") or 0)
            candidates = [v for v in sorted(set(valid), reverse=True) if v > 0 and abs(v - gt) > 10.0]
            if not candidates: candidates = sorted(set(valid), reverse=True)
            taxable = candidates[0]
            
            standard_rates = [0.025, 0.06, 0.09, 0.14, 0.005, 0.00125, 0.015]
            rate_match_val = 0
            for r in standard_rates:
                expected = round(taxable * r, 2)
                for v in valid:
                    if abs(v - expected) < 1.0:
                        count = valid.count(v)
                        has_total = any(abs(total - v*2) < 2.0 for total in valid)
                        if count >= 2 or has_total:
                            rate_match_val = v
                            break
                if rate_match_val > 0: break
            
            if rate_match_val > 0:
                res["taxable"] = f"{taxable:.2f}"
                if has_cgst: res["cgst"] = f"{rate_match_val:.2f}"
                if has_sgst: res["sgst"] = f"{rate_match_val:.2f}"
                res["_tax_set"] = True
            else:
                # CONSENSUS HEURISTIC on merged block
                pair_val = 0
                for i in range(len(valid)):
                    for j in range(i + 1, len(valid)):
                        if abs(valid[i] - valid[j]) < 0.05:
                            sum_val = valid[i] + valid[j]
                            if any(abs(v - sum_val) < 1.1 for v in valid):
                                pair_val = valid[i]
                                break
                    if pair_val > 0: break
                
                if pair_val > 0:
                    res["taxable"] = f"{taxable:.2f}"
                    if has_cgst: res["cgst"] = f"{pair_val:.2f}"
                    if has_sgst: res["sgst"] = f"{pair_val:.2f}"
                    res["_tax_set"] = True

        if not res.get("_tax_set"):
            # Fallback for data rows if not set as a block (processed line by line)
            for i in range(summary_start, min(summary_start + 15, len(lines))):
                line = lines[i]
                l_up = line.upper().strip()
                if re.search(r'(?i)TAX\s*SUMMARY', l_up) and not re.search(r'\d{3,}', line.replace(',','').replace(' ','')): continue
                amounts = re.findall(r'(?<![\d\.])[₹]?\s*(\d[\d,]*(?:\.\d{1,2})?)\b(?!%)', line)
                valid = [InvoiceFieldExtractors._clean_num(a, reject_id) for a in amounts]
                valid = [v for v in valid if v > 2.0]
                if len(valid) >= 2:
                    # (Rest of fallback logic omitted for brevity as it is redundant to block logic)
                    unique_v = sorted(set(valid), reverse=True)
                    taxable = unique_v[0]
                    gt_val = float(res.get("grand_total") or 0)
                    for v in unique_v:
                        if gt_val > 0 and abs(v - gt_val) < 1.0: continue
                        if not res.get("taxable"):
                            res["taxable"] = f"{v:.2f}"
                            continue
                        elif not res.get("_tax_set"):
                            tax_v = min(valid)
                            if has_cgst: res["cgst"] = f"{tax_v:.2f}"
                            if has_sgst: res["sgst"] = f"{tax_v:.2f}"
                            res["_tax_set"] = True
                            break
        res.pop("_tax_set", None)
        return res

    @staticmethod
    def extract_amounts(text: str) -> Dict[str, str]:
        res = {
            "taxable": "", "cgst": "", "cgst_rate": "",
            "sgst": "", "sgst_rate": "",
            "igst": "", "igst_rate": "",
            "grand_total": "", "round_off": ""
        }
        
        # Identify Invoice Number FIRST for blacklist
        inv_no_str = InvoiceFieldExtractors.extract_invoice_number(text)
        reject_id = None
        try:
            reject_id = float(re.sub(r'[^\d]', '', inv_no_str)) if inv_no_str else None
        except: pass
        
        is_angel = any(x in text.upper() for x in ['ANGEL ENTERPRISE', 'ANGEL ENTERPRISES', 'ANGEL HARDWARE', 'Slab Taxable Value'])
        if is_angel and 'RATE WISE SUMMARY' in text.upper():
            angel_res = InvoiceFieldExtractors._extract_angel_summary(text)
            if angel_res["taxable"] and (angel_res["cgst"] or angel_res["grand_total"]):
                for k in res:
                    if angel_res.get(k): res[k] = angel_res[k]
                return res


        lv = InvoiceFieldExtractors._extract_label_value_pairs(text, reject_id)
        has_lv = bool(lv["grand_total"] or lv["taxable"])
        
        grid = InvoiceFieldExtractors._extract_tax_summary_grid(text, reject_id)
        has_grid = bool(grid["taxable"] and (grid["igst"] or grid["cgst"]))

        sections = InvoiceFieldExtractors._get_sections(text)
        totals_text = sections["totals"] if sections["totals"] else text
        ctx = InvoiceFieldExtractors._extract_context_based(totals_text, reject_id)

        # Merge Scoring
        def _balance_score(d):
            t = float(d.get('taxable') or 0)
            g = float(d.get('grand_total') or 0)
            c = float(d.get('cgst') or 0)
            s = float(d.get('sgst') or 0)
            ig = float(d.get('igst') or 0)
            r = float(d.get('round_off') or 0)
            if g == 0 and t == 0: return 999999
            if g < 50 and t < 50: return 999998
            if g == 0: g = t
            return abs((t + c + s + ig + r) - g)

        
        strategies = []
        if has_grid:
            merged_grid = dict(grid)
            if not merged_grid.get('grand_total'):
                merged_grid['grand_total'] = lv.get('grand_total') or ctx.get('grand_total') or ''
            strategies.append(merged_grid)
        if has_lv:
            merged_lv = dict(lv)
            for k in ['cgst', 'sgst', 'igst', 'grand_total']:
                if not merged_lv.get(k) and ctx.get(k): merged_lv[k] = ctx[k]
            strategies.append(merged_lv)
        strategies.append(ctx)
        
        best = min(strategies, key=_balance_score)
        for k in ['taxable', 'cgst', 'sgst', 'igst', 'grand_total', 'round_off']:
            if best.get(k): res[k] = best[k]

        # Final Math balancing (Refined to avoid noise subtraction)
        try:
            t, g = float(res["taxable"] or 0), float(res["grand_total"] or 0)
            c, s, ig = float(res["cgst"] or 0), float(res["sgst"] or 0), float(res["igst"] or 0)
            
            if g > 50:
                r_off = float(res.get("round_off") or 0)
                # Handle signed round off (e.g. -0.18)
                if "- " in text or ": -" in text:
                    signed_m = re.search(r'(?i)ROUND[\sOFF]*[:\-]*\s*(-[\d,]+\.?\s*\d{2})', text)
                    if signed_m: r_off = float(signed_m.group(1).replace(',', ''))
                
                # BigBond special: Basic - Discount = Taxable
                disc = float(res.get("discount") or 0)
                if disc > 0 and abs((t - disc) + c + s + ig + r_off - g) < 1.0:
                    t = round(t - disc, 2); res["taxable"] = f"{t:.2f}"

                # If math is unbalanced and we have a grand total, trust the total and recalculate taxable
                curr_sum = t + c + s + ig + r_off
                if g > 50 and abs(curr_sum - g) > 2.0:
                    # Recalculate taxable as the user suggested: T = G - Taxes - RoundOff
                    new_t = round(g - (c + s + ig + r_off), 2)
                    if new_t > 0:
                        t = new_t; res["taxable"] = f"{t:.2f}"
                
                if t == 0 and (c+s+ig+r_off) > 0 and g > (c+s+ig+r_off):

                    t = round(g - (c+s+ig+r_off), 2); res["taxable"] = f"{t:.2f}"
                if t == 0 and (c+s+ig+r_off) == 0: t = g; res["taxable"] = f"{t:.2f}"
                if t > 0 and (c+s+ig+r_off) == 0 and g > (t+r_off):
                    ig = round(g - t - r_off, 2); res["igst"] = f"{ig:.2f}"
                if t > 0 and c > 0 and s == 0 and abs(g - (t+c+c+r_off)) < 2.5:
                    s = c; res["sgst"] = f"{s:.2f}"
                if t > 0 and s > 0 and c == 0 and abs(g - (t+s+s+r_off)) < 2.5:
                    c = s; res["cgst"] = f"{c:.2f}"
                
                # REFINEMENT: If CGST and SGST are extracted but vastly different (fragmentation noise)
                # and their sum doesn't match expected tax, try to equalize them if one looks like a fragment
                if t > 0 and c > 0 and s > 0 and abs(c - s) > 1.0:
                    expected_total_tax = round(g - t - r_off, 2)
                    if abs((c + s) - expected_total_tax) > 2.0:
                        # If one is double the other, or one is very small
                        if abs(c*2 - expected_total_tax) < 2.0: 
                             s = c; res["sgst"] = f"{s:.2f}"
                        elif abs(s*2 - expected_total_tax) < 2.0:
                             c = s; res["cgst"] = f"{c:.2f}"

                
                curr_g = round(t + c + s + ig, 2)
                if (g == 0 or abs(curr_g - g) > 2.0) and curr_g > 50:
                    if curr_g >= t: g = curr_g; res["grand_total"] = f"{g:.2f}"
        except: pass
        
        for k, lbl in [("cgst_rate", "CGST"), ("sgst_rate", "SGST"), ("igst_rate", "IGST")]:
            m_rate = re.search(r'(?i)' + lbl + r'[\s@]*(\d{1,2}(?:\.\d+)?)[\s]*%', text)
            if m_rate: res[k] = m_rate.group(1)
        return res

    @staticmethod
    def _extract_context_based(totals_text: str, reject_id: float = None) -> Dict[str, str]:
        """Legacy context-based extraction for BigBond and similar layouts."""
        res = {"taxable": "", "cgst": "", "sgst": "", "igst": "", "grand_total": "", "round_off": ""}
        lines = [l.strip() for l in totals_text.split('\n') if l.strip()]
        
        map_keys = {
            "taxable": ['TOTAL TAXABLE AMOUNT', 'TAXABLE VALUE', 'SUB TOTAL', 'BEFORE TAX', 'NET TAXABLE', 'TOTAL AMOUNT BEFORE TAX', 'TAXABLE AMOUNT', 'BASIC AMOUNT'],
            "total": ['GRAND TOTAL', 'AMOUNT PAYABLE', 'TOTAL AMOUNT', 'NET AMOUNT', 'AMOUNT WITH TAX', 'G.TOTAL AMOUNT', 'G.TOTAL', 'TOTAL PAYABLE', 'TOTAL'],

            "igst": ['ADD : IGST', 'INTEGRATED TAX', 'ADD: IGST', 'IGST'],
            "cgst": ['ADD : CGST', 'CENTRAL TAX', 'ADD: CGST', 'CGST', 'CENTRAL'],
            "sgst": ['ADD : SGST', 'STATE TAX', 'ADD: SGST', 'SGST', 'STATE/UT', 'STATE / UT'],

            "round_off": ['ROUND OFF', 'ROUNDOFF', 'ROUNDING'],
            "discount": ['DISC.AMT', 'DISCOUNT', 'LESS DISC']
        }




        
        candidates = {k: [] for k in map_keys}
        for i, line in enumerate(lines):
            l_up = line.upper()
            # Explicitly ignore Qty/Unit/HSN/Bank rows
            if any(x in l_up for x in ['HSN', 'SAC', 'QTY', 'UNIT', 'GSTIN', 'BANK', 'A/C', 'CODE', 'BRANCH', 'IFSC', 'STATE CODE', 'SR.', 'SR NO', 'SRNO', 'NOS']): continue
            # Avoid header labels
            if any(h in l_up for h in ['INVOICE NO', 'DATE', 'BILL NO']): continue
            
            line_clean = re.sub(r'(\d{1,2}(?:\.\d+)?)\s*%', '', line)
            nums = [InvoiceFieldExtractors._clean_num(n, reject_id) for n in re.findall(r'(\d+[\d,]*(?:\.\d{1,2})?)', line_clean)]
            if not nums: continue
            
            # Explicitly reject standalone '1', '2', '3' if they appear at the start of a line (often Sr No)
            nums = [n for n in nums if not (n < 4 and re.match(r'^\s*[\d]+\s*$', line))]
            if not nums: continue
            
            # Use a slightly adaptive lookback: for total/roundoff, look only at 1 lines back or current line
            lb_full = " ".join([l.upper() for l in lines[max(0, i-3):i+1]])
            lb_strict = " ".join([l.upper() for l in lines[max(0, i-1):i+1]])

            
            for j, val in enumerate(nums):
                if val <= 0.01: continue
                
                # REJECT if line contains Shipping/Freight/HSN noise
                if any(x in l_up for x in ['SHIPPING', 'FREIGHT', 'POSTAGE', 'CHARGE', 'HSN']): 
                    continue

                assigned = False

                # Narrow label matching: Look for 'ADD :' specifically for small values to avoid grid noise
                if val < 50:
                    if 'ADD : CGST' in lb_full or 'CENTRAL TAX' in lb_full:
                        candidates["cgst"].append(val); assigned = True
                    elif 'ADD : SGST' in lb_full or 'STATE TAX' in lb_full:
                        candidates["sgst"].append(val); assigned = True
                    elif 'ADD : IGST' in lb_full or 'INTEGRATED TAX' in lb_full:
                        candidates["igst"].append(val); assigned = True

                if not assigned:
                    for k in ["taxable", "total", "round_off"]:
                        # Strict check: only look at current or previous line
                        if any(lbl in lb_strict for lbl in map_keys[k]):
                            if 'TOTAL' in lb_strict:
                                if any(x in lb_strict for x in ['TAXABLE', 'BASIC']): continue
                                if any(q in l_up for q in ['QTY', 'UNIT', 'NOS', 'NOS.', 'PCS', 'PCS.']):
                                    if k == "total": continue
                            candidates[k].append(val); assigned = True; break

                if not assigned:
                    for k in ["cgst", "sgst", "igst", "discount"]:
                        if any(lbl in lb_full for lbl in map_keys[k]):
                            candidates[k].append(val); assigned = True; break

                if not assigned:
                    for k in ["taxable", "total"]:
                        if any(lbl in lb_full for lbl in map_keys[k]):
                            candidates[k].append(val); assigned = True; break



        if candidates["taxable"]: res["taxable"] = f"{max(candidates['taxable']):.2f}"
        if candidates["cgst"]: res["cgst"] = f"{max(candidates['cgst']):.2f}"
        if candidates["sgst"]: res["sgst"] = f"{max(candidates['sgst']):.2f}"
        if candidates["igst"]: res["igst"] = f"{max(candidates['igst']):.2f}"
        if candidates["total"]: res["grand_total"] = f"{max(candidates['total']):.2f}"
        if candidates["round_off"]: res["round_off"] = f"{candidates['round_off'][0]:.2f}"
        return res

    @staticmethod
    def _extract_angel_summary(text: str) -> Dict[str, str]:
        res = {"taxable": "", "cgst": "", "sgst": "", "igst": "", "grand_total": ""}
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        start = -1
        for i in range(len(lines)-1, -1, -1):
            if 'RATE WISE SUMMARY' in lines[i].upper(): start = i; break
        if start != -1:
            for j in range(start + 1, min(start + 15, len(lines))):
                l_up = lines[j].upper()
                if any(x in l_up for x in ['TERMS', 'CONDITION', 'FOR,']): break
                
                l_c = re.sub(r'\b\d{1,2}(?:\.\d+)?\s*%', '', l_up)
                nums = [InvoiceFieldExtractors._clean_num(a) for a in re.findall(r'(\d+[\d,]*\.\d{1,2})', l_c)]
                if len(nums) >= 3:
                    sorted_m = sorted([n for n in nums if n > 0.01], reverse=True)
                    res["taxable"] = f"{(float(res['taxable'] or 0) + sorted_m[0]):.2f}"
                    res["cgst"] = f"{(float(res['cgst'] or 0) + sorted_m[1]):.2f}"
                    res["sgst"] = f"{(float(res['sgst'] or 0) + sorted_m[2]):.2f}"
        if not res["grand_total"]:
            for l_up in [lx.upper() for lx in lines]:
                if any(x in l_up for x in ['GRAND TOTAL', 'BILL AMOUNT', 'NET AMOUNT', 'TOTAL']):
                    gt_m = re.findall(r'(\d+[\d,]*\.\d{2})', l_up)
                    if gt_m:
                        val = InvoiceFieldExtractors._clean_num(gt_m[-1])
                        if not res["grand_total"] or val > float(res["grand_total"]): 
                            res["grand_total"] = f"{val:.2f}"
        if not res["grand_total"] and res["taxable"]:
            res["grand_total"] = f"{(float(res['taxable']) + float(res['cgst'] or 0) + float(res['sgst'] or 0)):.2f}"
        return res

    @staticmethod
    def extract_invoice_number(text: str) -> str:
        patterns = [
            r'(?i)(?:Invoice No\.?|Inv\.? No\.?|Bill No\.?|Voucher\s*No\.?|INVOICE\s*NUMBER)[\s\.\:\#\-]*\s*\n?[\s\.\:\#\-]*([A-Z0-9\/\\ \-]{1,15})',
            r'(?i)(?:No\.?|\#)[\s\.\:\#\-]*\s*\n?[\s\.\:\#\-]*([A-Z0-9\/\\ \-]{1,15})'
        ]
        reject_p = r'(?i)\b(PHONE|MOBILE|CONTACT|ACCOUNT|BANK|VEHICLE|CHALLAN|ORDER|GSTIN|HSN|SAC|ITEM|QTY|RATE|PRICE|VALUE|AMOUNT|TOTAL|YES|NO|TRUE|FALSE)\b'
        for p in patterns:
            for m in re.finditer(p, text):
                cand = re.sub(r'^[\:\.\-\#\s]+', '', re.split(r'\s{2,}', m.group(1).strip())[0].strip()).strip()
                # Basic cleaning of invoice number
                cand = re.sub(r'^0+', '', cand) if cand.isdigit() else cand
                if cand and not (cand.isdigit() and len(cand) >= 10) and not re.search(reject_p, cand.upper()):
                    if not any(x in cand.upper() for x in ['GIDC', 'ROAD', 'HIGHWAY', 'STREET', 'NEAR', 'PAISA', 'RUPEES']): return cand
        return ""

    @staticmethod
    def extract_date(text: str) -> str:
        patterns = [
            r'(?i)(?:Invoice\s*Date|Date|Dt\.?)[\s\.\:\#\-]*(\d{1,2}[\/\-\.\s]+(?:[A-Za-z]{3,10}|\d{1,2})[\/\-\.\s,]+\d{2,4})',
            r'\b(\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2,4})\b',
            r'(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})'
        ]
        for p in patterns:
            m = re.search(p, text)
            if m: return m.group(1).strip()
        return ""

    @staticmethod
    def extract_party_name(text: str) -> str:
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        receiver_labels = ['DETAILS OF RECEIVER', 'BILL TO', 'BUYER', 'BILLED TO', 'RECEIVER', 'PARTY DETAILS', 'BUYER TO PARTY', 'CUSTOMER']
        consignee_labels = ['DETAILS OF CONSIGNEE', 'SHIPPED TO', 'SHIP TO', 'CONSIGNEE', 'CUSTOMER']
        table_pattern = r'(?i)\b(SR[ \.\n]*NO|DESCRIPTION|NAME\s*OF\s*PRODUCT|HSN|PARTICULARS|ITEM\s*NAME|QTY|UNIT|RATE|TAXABLE)\b'
        start_idx = -1
        for i, line in enumerate(lines):
            l_up = line.upper()
            if any(l in l_up for l in receiver_labels) or 'M/S' in l_up: start_idx = i; break
        if start_idx == -1:
            for i, line in enumerate(lines):
                if any(l in line.upper() for l in consignee_labels): start_idx = i; break
        if start_idx == -1: return ""
        end_idx = len(lines)
        for j in range(start_idx + 1, min(start_idx + 30, len(lines))):
            if re.search(table_pattern, lines[j].upper()): end_idx = j; break
        block_lines = lines[start_idx:end_idx]
        blacklist = ['SR', 'SR.', 'NO', 'PRODUCT', 'HSN', 'QTY', 'UNIT', 'RATE', 'TAXABLE', 'CGST', 'SGST', 'IGST', 'TOTAL', 'DETAILS', 'BUYER', 'NAME', 'ADDRESS', 'BILL', 'BILLED']
        rejection_keywords = ['BANK', 'A/C', 'ACCOUNT', 'IFSC', 'BRANCH', 'ICICI', 'HDFC', 'SBI', 'AXIS', 'PNB', 'BOB', 'KOTAK', 'YES BANK', 'IDFC']
        skip_labels = ['NAME', 'ADDRESS', 'MOBILE', 'MOB', 'STATE', 'GSTIN', 'PHONE', 'PAN', 'PLACE', 'BILL', 'BILLED', 'DETAILS OF', 'BILL TO', 'BILLED TO', 'CUSTOMER']
        for line in block_lines:
            m = re.search(r'(?i)(?:NAME|M/S\.?)\s*[\:\.\-\s]*\s*([A-Z0-9][A-Z0-9\s\.\&\-\(\)\/\\,\,]{3,100})', line)
            if m:
                cand = re.split(r'(?i)(\s{3,}|[|]|\t|ADDRESS|STATE|MOBILE|MOB|GSTIN|PHONE|PAN|PLACE|NO[\:\.]|DATE[\:\.])', m.group(1).strip())[0].strip()
                if len(cand) >= 3 and cand.upper() not in blacklist and not any(k in cand.upper() for k in rejection_keywords): return cand
        for line in block_lines:
            l_up = line.upper(); anchor_found = False
            for al in receiver_labels + consignee_labels:
                if al in l_up:
                    rem = l_up
                    for lbl in receiver_labels + consignee_labels + skip_labels: rem = rem.replace(lbl, '')
                    if len(re.sub(r'[\|\:\.\-\s\d]+', '', rem).strip()) < 4: anchor_found = True; break
            if anchor_found or l_up.strip() in ['M/S.', 'M/S']: continue
            if any(l_up.startswith(lbl) for lbl in skip_labels): continue
            if any(k in l_up for k in rejection_keywords): continue
            cand = re.split(r'\s{4,}|[|]', re.sub(r'^[\|\.\:\s\-]+', '', line).strip())[0].strip()
            if len(cand) >= 4 and cand.upper() not in blacklist:
                if not re.match(r'^\d+$', cand) and not re.match(r'^[0-3][0-9][A-Z]{5}', cand):
                    return re.sub(r'(?i)^\s*M/S\.?\s*', '', cand).strip()
        return ""

    @staticmethod
    def extract_buyer_gstin(text: str) -> str:
        sections = InvoiceFieldExtractors._get_sections(text)
        bill_to, header = sections["bill_to"].upper(), sections["header"].upper()
        inline_p = r'\b([0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])\b'
        def find_gst(src):
            m = re.search(inline_p, src)
            if m: return m.group(1)
            
            # Support 14 chars + 1 standalone char on same or next line
            # First, find 14-char base
            base_p = r'\b([0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z)\b'
            for m_base in re.finditer(base_p, src):
                base = m_base.group(1)
                tail_src = src[m_base.end():m_base.end()+50]
                # Look for a standalone character (0-9A-Z), skipping 'PAN No.' noise
                tail_m = re.search(r'(?i)(?:\s+|PAN(?:\s*No\.?)?|\:)*\b([0-9A-Z])\b', tail_src)
                if tail_m:
                    return base + tail_m.group(1).upper()
            return None

        supp_g = find_gst(header) or ""
        buyer_g = find_gst(bill_to)
        if buyer_g: return buyer_g
        all_text = text.upper()
        for m in re.finditer(inline_p, all_text):
            if m.start() > 400 and m.group(1) != supp_g: return m.group(1)
        for m in re.finditer(r'\b([0-3][0-9][A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z)\s*\n?\s*([0-9A-Z])', all_text):
            g = m.group(1) + m.group(2)
            if m.start() > 400 and g != supp_g: return g
        return ""

    @staticmethod
    def extract_supplier_name(text: str) -> str:
        lines = [l.strip() for l in InvoiceFieldExtractors._get_sections(text)["header"].split('\n') if l.strip()]
        return re.sub(r'(?i)TAX INVOICE|INVOICE', '', lines[0]).strip() if lines else ""
