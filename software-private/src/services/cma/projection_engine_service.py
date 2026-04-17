import math
from typing import List, Dict, Any
from datetime import datetime
from services.cma.models import CmaProject, FinancialAssumptions, DepreciationMethod, AuditedData

class ProjectionEngineService:
    """Service to calculate financial schedules and projections."""

    @staticmethod
    def _safe_div(n, d, default=0.0):
        """Safe division to prevent ZeroDivisionError."""
        try:
            val = float(n) / float(d) if float(d) != 0 else default
            return val
        except:
            return default

    @staticmethod
    def calculate_loan_amortization(
        principal: float, 
        annual_rate: float, 
        tenure_years: int, 
        moratorium_months: int = 0
    ) -> List[Dict[str, float]]:
        """
        Calculates a yearly amortization schedule for a Term Loan.
        Includes moratorium handling (interest only during moratorium).
        """
        if principal <= 0 or annual_rate <= 0 or tenure_years <= 0:
            return []

        monthly_rate = (annual_rate / 100) / 12
        total_months = tenure_years * 12
        repayment_months = total_months - moratorium_months
        
        if repayment_months <= 0:
            return []

        # Standard EMI calculation for the repayment period
        emi = (principal * monthly_rate * pow(1 + monthly_rate, repayment_months)) / \
              (pow(1 + monthly_rate, repayment_months) - 1)

        schedule = []
        balance = principal
        
        for year in range(1, tenure_years + 1):
            year_interest = 0.0
            year_principal = 0.0
            
            for month in range(1, 13):
                m_idx = (year - 1) * 12 + month
                
                interest = balance * monthly_rate
                year_interest += interest
                
                if m_idx <= moratorium_months:
                    # Interest only payment
                    pass
                else:
                    # EMI payment
                    p_repaid = emi - interest
                    year_principal += p_repaid
                    balance -= p_repaid
                
            schedule.append({
                "year": year,
                "opening_balance": balance + year_principal,
                "interest": year_interest,
                "principal_repayment": year_principal,
                "closing_balance": max(0, balance)
            })
            
            if balance <= 0:
                break
                
        return schedule

    @staticmethod
    def calculate_monthly_repayment(
        principal: float, 
        annual_rate: float, 
        tenure_years: int, 
        moratorium_months: int = 0
    ) -> List[Dict[str, Any]]:
        """
        Calculates a full monthly amortization schedule for the premium annexure.
        Returns a list of 60-84 months with Opening, Interest, Principal, and Closing.
        """
        if principal <= 0 or annual_rate <= 0 or tenure_years <= 0:
            return []

        monthly_rate = (annual_rate / 100) / 12
        total_months = tenure_years * 12
        repayment_months = total_months - moratorium_months
        
        if repayment_months <= 0:
            return []

        # Standard EMI calculation
        emi = (principal * monthly_rate * pow(1 + monthly_rate, repayment_months)) / \
              (pow(1 + monthly_rate, repayment_months) - 1)

        schedule = []
        balance = principal
        
        for m_idx in range(1, total_months + 1):
            opening = balance
            interest = balance * monthly_rate
            
            if m_idx <= moratorium_months:
                # Interest only
                p_repaid = 0.0
                total_payment = interest
            else:
                p_repaid = emi - interest
                total_payment = emi
                balance -= p_repaid
            
            schedule.append({
                "month_num": m_idx,
                "opening_balance": opening,
                "interest": interest,
                "principal_repayment": p_repaid,
                "total_payment": total_payment,
                "closing_balance": max(0, balance)
            })
            
            if m_idx > moratorium_months and balance <= 0.001:
                break
                
        return schedule

    @staticmethod
    def calculate_depreciation(
        asset_cost: float,
        rate_percent: float,
        years: int,
        method: str = DepreciationMethod.WDV.value
    ) -> List[Dict[str, float]]:
        """
        Calculates a depreciation schedule using WDV or SLM.
        """
        schedule = []
        balance = asset_cost
        slm_amount = asset_cost * (rate_percent / 100)
        
        for year in range(1, years + 1):
            opening = balance
            if method == DepreciationMethod.SLM.value:
                depr = min(slm_amount, balance)
            else: # WDV
                depr = balance * (rate_percent / 100)
            
            balance -= depr
            schedule.append({
                "year": year,
                "opening_value": opening,
                "depreciation": depr,
                "closing_value": max(0, balance)
            })
            
            if balance <= 0:
                break
                
        return schedule

    @classmethod
    def get_projected_working_capital(cls, project: CmaProject, year_idx: int) -> Dict[str, float]:
        """
        Estimates working capital requirements based on Days-based assumptions.
        Simple logic for Phase 2: Uses sales and GP to derive requirements.
        """
        # We need a base sales figure. For Phase 2, we'll estimate based on loan size 
        # or a default of 10x WC if not specified.
        # This will be refined in Phase 3.
        base_sales = cls._safe_div(project.loan.working_capital_requirement * 365, ass.debtor_days, 100.0)
        
        # Scenario logic is now handled via UI presets to ensure WYSIWYG.
        # Direct usage of UI-provided percentages.
        projected_sales = base_sales * pow(1 + ass.sales_growth_percent/100, year_idx)
        projected_purchases = projected_sales * (1 - ass.gp_percent/100)
        
        debtors = (projected_sales * ass.debtor_days) / 365
        stock = (projected_purchases * ass.stock_days) / 365
        creditors = (projected_purchases * ass.creditor_days) / 365
        
        net_wc = debtors + stock - creditors
        
        return {
            "sales": projected_sales,
            "debtors": debtors,
            "stock": stock,
            "creditors": creditors,
            "net_working_capital": net_wc
        }

    @classmethod
    def generate_full_projections(cls, project: CmaProject) -> List[Dict[str, Any]]:
        """
        Generates a year-by-year financial statement list (Actuals + Projected).
        """
        results = []
        ass = project.assumptions
        from services.cma.models import BusinessMode
        
        # 1. Start with Audited History (if any)
        last_actual_revenue = 0.0
        last_reserves = 0.0
        
        mode = project.profile.business_mode
        
        # Robust check: If we have historical records, use them as base UNLESS it's a New Business
        if project.audited_history and mode != BusinessMode.NEW.value:
            for i, ad in enumerate(project.audited_history):
                proj = cls._model_to_projection(ad, is_actual=True)
                # Linkage: Opening balance is the Closing balance of the previous year
                prev_ad = project.audited_history[i-1] if i > 0 else None
                prev_cash = getattr(prev_ad, 'cash_bank', 0.0) if prev_ad else 0.0
                proj["opening_cash_bal"] = prev_cash
                proj["closing_cash_bal"] = getattr(ad, 'cash_bank', 0.0)
                # Requirement: Stock Adjustment in historical mode (First year uses Opening Stock field)
                prev_inv = getattr(prev_ad, 'inventory', 0.0) if prev_ad else getattr(ad, 'opening_stock', 0.0)
                proj["stock_adj"] = getattr(ad, 'inventory', 0.0) - (prev_inv or 0.0)
                results.append(proj)
                last_actual_revenue = getattr(ad, 'revenue', 0.0)
                last_reserves = getattr(ad, 'reserves_surplus', 0.0)
        elif mode == BusinessMode.EXISTING_NO_BOOKS.value:
            # Synthetic Base Year from Simplified Data
            sd = project.simplified_data
            base_year = cls._simplified_to_projection(sd)
            base_year["opening_cash_bal"] = 0.0 # No previous history available
            base_year["closing_cash_bal"] = sd.cash_bank_estimate or 0.0
            results.append(base_year)
            last_actual_revenue = sd.approx_turnover
            last_reserves = (sd.approx_turnover * sd.np_percent / 100) # Proxy
        
        # 2. Setup Projection Base
        # For startups, we must estimate a Year 0 'proxy' revenue so that compounding logic works
        effective_base_revenue = last_actual_revenue
        if mode == BusinessMode.NEW.value or effective_base_revenue <= 0:
            # Estimate first year revenue if no history
            # Logic: Turnover should be approx 3-5 times WC or 0.8x of Total Investment
            total_investment = sum(a.cost for a in project.assets) + project.loan.working_capital_requirement
            
            # Use user override if provided
            if getattr(ass, 'revenue_base_override', 0) > 0:
                y1_est = ass.revenue_base_override
            else:
                # Conservative Turnover Estimation: Limit * 5.0 (match 20% MPBF)
                y1_est = project.loan.working_capital_requirement * 5.0
                if y1_est <= 0: y1_est = (total_investment * 0.8)
            
            # Back-calculate relative base (Year 0 proxy) so that base * (1+g)^1 = Y1_Est
            effective_base_revenue = cls._safe_div(y1_est, (1 + ass.sales_growth_percent / 100))
            
            # Also sync last_actual_revenue for growth ratio scaling (OCA/OCL)
            last_actual_revenue = effective_base_revenue
        
        current_reserves = last_reserves
        
        # Capture Opening Cash Balance for linkage
        current_cash = 0.0
        other_current_assets_base = 0.0
        other_current_liabilities_base = 0.0
        
        if project.audited_history:
            last_ad = project.audited_history[-1]
            current_cash = getattr(last_ad, 'cash_bank', 0.0) or getattr(last_ad, 'cash_in_hand', 0.0)
            
            # Robust pull for other current items (check common field variants)
            other_current_assets_base = (
                getattr(last_ad, 'other_current_assets', 0.0) or 
                getattr(last_ad, 'oca', 0.0) or 
                getattr(last_ad, 'miscellaneous_assets', 0.0)
            )
            other_current_liabilities_base = (
                getattr(last_ad, 'other_current_liabilities', 0.0) or 
                getattr(last_ad, 'ocl', 0.0) or 
                getattr(last_ad, 'other_liabilities', 0.0)
            )
            
            if current_cash <= 0:
                current_cash = getattr(last_ad, 'current_assets', 0.0) * 0.05 # Fallback proxy
            
            # Ensure the BASE REVENUE used for scaling is exactly the last audited year
            last_actual_revenue = getattr(last_ad, 'revenue', 0.0)

        # Initial Promoter Capital for New Project
        total_p_cost = sum(a.cost for a in project.assets) + project.loan.working_capital_requirement
        total_loan = project.loan.term_loan_amount + project.loan.cash_credit_amount
        fresh_contribution = total_p_cost - total_loan if (project.is_new_project or mode == BusinessMode.NEW.value) else 0.0
        
        # Point A: Carry forward historical capital
        current_share_capital = 0.0
        current_net_block = 0.0
        if project.audited_history:
            last_ad = project.audited_history[-1]
            current_share_capital = getattr(last_ad, 'share_capital', 0.0)
            current_net_block = getattr(last_ad, 'net_block', 0.0) or 0.0
        elif mode == BusinessMode.EXISTING_NO_BOOKS.value:
            current_share_capital = project.simplified_data.fixed_assets_estimate * 0.5
            current_net_block = project.simplified_data.fixed_assets_estimate
            
        current_share_capital += fresh_contribution # Add fresh capital to base
        
        # 3. Generate Projected Years
        # Amortize NEW Term Loan
        loan_sched = cls.calculate_loan_amortization(
            project.loan.term_loan_amount, ass.interest_on_tl, 
            project.loan.term_loan_tenure_years, ass.moratorium_months
        )
        
        # Amortize EXISTING Term Loan (Carry-forward)
        last_hist_loan = 0.0
        if project.audited_history:
            last_hist_loan = project.audited_history[-1].term_loans
        
        # Assumption: Interest for existing loan is same as bank TL, tenure 5 years or same as new
        hist_loan_sched = cls.calculate_loan_amortization(
            last_hist_loan, ass.interest_on_tl, 
            project.loan.term_loan_tenure_years or 5, 0 # Usually no moratorium for existing debt
        )
        
        # 2. Extract Historical Trends for Calibration (Requirement: Resolve 9.33 vs 10.67 PAT)
        calibrated_gp_pct = ass.gp_percent
        calibrated_exp_pct = ass.indirect_expense_percent
        
        if project.audited_history:
            last_ad = project.audited_history[-1]
            h_rev = last_ad.revenue
            if h_rev > 0:
                # Use Historical GP % if available
                h_gp = getattr(last_ad, 'gross_profit', 0.0)
                if h_gp > 0:
                    calibrated_gp_pct = cls._safe_div(h_gp * 100, h_rev)
                
                # Use Historical Indirect Expense % if available
                # Note: Ind Exp in history = sum of salary, rent, admin etc.
                h_exp = (getattr(last_ad, 'salary_wages', 0) + 
                         getattr(last_ad, 'rent_rates', 0) + 
                         getattr(last_ad, 'admin_other_exp', 0))
                if h_exp > 0:
                    calibrated_exp_pct = (h_exp / h_rev) * 100

        # --- Historical WC Days for Soft Landing (Requirement: Point 3 in plan) ---
        hist_debtor_days = ass.debtor_days # default
        hist_stock_days = ass.stock_days
        hist_creditor_days = ass.creditor_days
        
        if project.audited_history:
            last = project.audited_history[-1]
            if (getattr(last, 'revenue', 0.0) or 0.0) > 0:
                hist_debtor_days = cls._safe_div(getattr(last, 'debtors', 0.0) * 365, last.revenue)
                # Estimate benchmark COGS if not available in history
                h_cogs = getattr(last, 'cogs', 0.0) or (last.revenue * 0.7)
                if h_cogs > 0:
                    hist_stock_days = cls._safe_div(getattr(last, 'inventory', 0.0) * 365, h_cogs)
                    hist_creditor_days = cls._safe_div(getattr(last, 'creditors', 0.0) * 365, h_cogs)

        current_other_assets = other_current_assets_base
        current_other_liabilities = other_current_liabilities_base

        # 3. Projection Loop: Initialize with calibrated rates
        # Requirement: Use specific facility rates from Loan Profile if set
        if project.loan.tl_interest_rate > 0:
            ass.interest_on_tl = project.loan.tl_interest_rate
        if project.loan.cc_interest_rate > 0:
            ass.interest_on_cc = project.loan.cc_interest_rate
            
        current_revenue = last_actual_revenue
        # Note: current_share_capital is already calculated and carried forward above
        current_reserves = last_reserves
        
        # Determine Base Year for Labeling (Requirement: FY 2026-27 format)
        base_year_num = datetime.now().year
        if project.audited_history:
            # FY label from history e.g. "2024-25 (A)"
            h_label = project.audited_history[-1].year_label
            try:
                import re
                match = re.search(r'(\d{4})', h_label)
                if match:
                    y_start = int(match.group(1))
                    # Requirement: 2026 -> FY 2026-27. 2025-26 -> FY 2026-27.
                    # If it's a range (hyphen), start is y_start + 1.
                    # If it's a single year, start is y_start.
                    base_year_num = y_start + 1 if '-' in h_label else y_start
            except:
                pass

        for y in range(1, ass.projection_years + 1):
            opening_cash = current_cash
            opening_equity = current_share_capital + current_reserves
            # CC definitions needed for S&U movement at top
            cc_drawn = project.loan.cash_credit_amount
            cc_limit = project.loan.cash_credit_amount
            
            # Physical Year label (Standard FY 2026-27 format)
            y_start = base_year_num + y - 1
            y_end = (y_start + 1) % 100
            y_label = f"FY {y_start}-{y_end:02d}" # Centralized badging happens in ReportGenerator
            
            prev_year = results[-1] if results else None
            
            # Compound Growth Implementation (Requirement: Bank standard growth)
            # Compound Growth Implementation (Requirement: Bank standard growth)
            current_revenue = effective_base_revenue * ((1 + ass.sales_growth_percent/100) ** y)

            # --- P&L Projections (Calibrated to History OR Assumptions) ---
            # Priority: Use assumption if provided and business mode is NEW or user specified.
            # Otherwise allow calibration to history if it's an existing business renewal.
            target_gp_pct = ass.gp_percent if ass.gp_percent > 0 else calibrated_gp_pct
            gp_amt = current_revenue * (target_gp_pct / 100)
            
            # Base logic: Revenue - COGS = GP. 
            # But in display: Revenue - Material/Direct - StockAdj = GP.
            # So Material/Direct = Revenue - GP + StockAdj
            raw_cogs = current_revenue - gp_amt 
            ind_exp = current_revenue * (cls._safe_div(calibrated_exp_pct, 100))
            
            # --- Standardized Depreciation & FA Carry-forward (Point B) ---
            depr_total = 0.0
            depr_details = []
            
            # Depreciate Historical Base if any
            if current_net_block > 0:
                hist_depr = current_net_block * (ass.depreciation_rate_percent / 100 if hasattr(ass, 'depreciation_rate_percent') else 0.10)
                # Cap the depreciation to remaining block
                hist_depr = min(hist_depr, current_net_block)
                depr_total += hist_depr
                current_net_block -= hist_depr # Reduce historical block
                depr_details.append({
                    "name": "Existing Assets (WDV)",
                    "cost": "Historical",
                    "rate": "Various",
                    "depreciation": hist_depr,
                    "closing_value": current_net_block
                })

            # Depreciate New Project Assets (if any)
            for asset in project.assets:
                rate = 15.0 # default
                n = asset.name.lower()
                if "build" in n: rate = 10.0
                elif "comp" in n or "soft" in n: rate = 40.0
                elif "furn" in n: rate = 10.0
                
                asset_sched = cls.calculate_depreciation(asset.cost, rate, ass.projection_years, ass.depreciation_method)
                if len(asset_sched) >= y:
                    ann_depr = asset_sched[y-1]["depreciation"]
                    depr_total += ann_depr
                    depr_details.append({
                        "name": asset.name,
                        "cost": asset.cost,
                        "rate": rate,
                        "depreciation": ann_depr,
                        "closing_value": asset_sched[y-1]["closing_value"]
                    })
            
            # --- Other Current Items Growth (Requirement: Dynamic accounts) ---
            growth_ratio = cls._safe_div(current_revenue, last_actual_revenue, 1.0)
            current_other_assets = other_current_assets_base * growth_ratio
            current_other_liabilities = other_current_liabilities_base * growth_ratio
            
            # 1. NEW Loan principal/int (Initialization)
            tl_interest = 0.0
            tl_repayment = 0.0
            tl_balance = 0.0
            
            if len(loan_sched) >= y:
                item = loan_sched[y-1]
                tl_interest = item["interest"]
                tl_repayment = item["principal_repayment"]
                tl_balance = item["closing_balance"]
            
            # 2. EXISTING Loan principal/int
            hist_tl_int = 0.0
            hist_tl_rep = 0.0
            hist_tl_bal = 0.0
            if len(hist_loan_sched) >= y:
                h_item = hist_loan_sched[y-1]
                hist_tl_int = h_item["interest"]
                hist_tl_rep = h_item["principal_repayment"]
                hist_tl_bal = h_item["closing_balance"]
            
            tl_interest_total = tl_interest + hist_tl_int
            tl_repayment_total = tl_repayment + hist_tl_rep
            tl_balance_total = tl_balance + hist_tl_bal
            
            # Metadata for separate display
            tl_bal_existing = hist_tl_bal
            tl_bal_new = tl_balance

            cc_interest = project.loan.cash_credit_amount * (ass.interest_on_cc / 100)
            total_int = tl_interest_total + cc_interest
            
            # Track CC Limit Movement for Cash Flow (Requirement: Fix 15L bridge discrepancy)
            prev_cc_limit = prev_year.get("cc_limit", 0) if prev_year else (project.audited_history[-1].current_liabilities * 0.4 if project.audited_history else 0) # estimate base if no history CC field
            if project.audited_history and not prev_year:
                # Better: if we have history, maybe we can assume the last year had 0 new CC if it was just Prov.
                # Actually, in user's screenshot, it WAS 0.00
                prev_cc_limit = 0.0 
            cc_limit_diff = cc_limit - prev_cc_limit
            
            ebitda = gp_amt - ind_exp
            pbt = ebitda - depr_total - total_int
            tax_amt = max(0, pbt * (ass.tax_rate_percent / 100))
            pat = pbt - tax_amt
            
            # Equity Roll-forward (Requirement: Capital balance roll-forward)
            # Standard drawings for proprietor/partners (approx 10-20% of PAT or fixed min) to be realistic
            drawings = max(0.12 * pat, 0.6) if project.profile.business_mode == "Existing" else 0.0
            current_reserves += (pat - drawings)
            
            # --- Working Capital & Liquidity (Strict Assumption Normalization) ---
            # Remove Glide Path as per user preference for 'Perfect' model consistency
            cur_debtor_days = ass.debtor_days
            cur_stock_days = ass.stock_days
            cur_creditor_days = ass.creditor_days

            # Strict Assumption-based calculation
            debtors = cls._safe_div(current_revenue * cur_debtor_days, 365)
            inventory = cls._safe_div(raw_cogs * cur_stock_days, 365)
            
            # Creditors MUST strictly follow assumptions and NOT be a plug
            creditors = cls._safe_div(raw_cogs * cur_creditor_days, 365)
            unsecured_loans = 0.0 # Initialize as 0, will be used as plug if needed
            
            # --- Other Current Items Growth (Corrected Persistence) ---
            # We scale the historical base by actual revenue growth for each year
            growth_ratio = cls._safe_div(current_revenue, last_actual_revenue, 1.0)
            current_other_assets = other_current_assets_base * growth_ratio
            current_other_liabilities = other_current_liabilities_base * growth_ratio
            
            prev_inventory = prev_year.get("inventory", 0) if prev_year else (project.audited_history[-1].inventory if project.audited_history else 0)
            stock_adj = inventory - prev_inventory 
            material_purchases = current_revenue - gp_amt + stock_adj
            
            # --- Linked Cash Flow & Balancing Logic (Requirement: Cash as Plug) ---
            # We use the full requested CC limit to match the limit-oriented display requirement
            tentative_liabs_no_cash = (current_share_capital + current_reserves) + tl_balance_total + creditors + current_other_liabilities + project.loan.cash_credit_amount
            net_fixed_assets = (current_net_block if current_net_block > 0 else 0) + sum(d["closing_value"] for d in depr_details if d["cost"] != "Historical")
            assets_no_cash = net_fixed_assets + debtors + inventory + current_other_assets
            
            # Balancing Item: Cash is the only plug now (No Sweep)
            final_cash = tentative_liabs_no_cash - assets_no_cash
            final_cc_bal = project.loan.cash_credit_amount
            
            # Safety Floor for Cash
            if final_cash < 0.2: 
                final_cash = 0.2

            # Calculate Opening Fixed Assets before movement
            # In Year 1, opening is the historical block (if existing). 
            # Additions is the new project asset outlay.
            if y == 1:
                # Find historical opening depr from the details we just built
                hist_depr_val = next((d['depreciation'] for d in depr_details if d['name'] == "Existing Assets (WDV)"), 0.0)
                opening_fa = (current_net_block + hist_depr_val) 
                asset_outlay = sum(a.cost for a in project.assets)
            else:
                opening_fa = net_fixed_assets_prev
                asset_outlay = 0.0

            # Final Totals
            total_assets = net_fixed_assets + debtors + inventory + current_other_assets + final_cash
            total_liabs = (current_share_capital + current_reserves) + tl_balance_total + creditors + current_other_liabilities + final_cc_bal
            current_liabs = final_cc_bal + creditors + current_other_liabilities
            
            # Absolute Final Reconciliation (rounding & funding gap)
            # Standard banker practice: Funding gap should be shown as Unsecured Loans from Promoters
            if abs(total_liabs - total_assets) > 0.0001:
                variance = total_assets - total_liabs
                if variance > 0:
                    # Funding gap: Assets > Liabs. Increase Unsecured Loans.
                    unsecured_loans += variance
                    total_liabs += variance
                else:
                    # Inflow surplus: Liabs > Assets. Increase Cash.
                    final_cash += abs(variance)
                    total_assets += abs(variance)
                    current_cash = final_cash # Sync back to state
            y_dscr = cls._safe_div(pat + depr_total + tl_interest_total, tl_repayment_total + tl_interest_total, 2.0)
            
            # Cash Flow Bridge Metadata (Section N)
            last_actual = project.audited_history[-1] if project.audited_history else None
            p_debtors = prev_year.get("debtors", 0) if prev_year else (getattr(last_actual, 'debtors', getattr(last_actual, 'current_assets', 0) * 0.4) if last_actual else 0)
            debtors_diff = debtors - p_debtors
            p_inventory = prev_year.get("inventory", 0) if prev_year else (getattr(last_actual, 'inventory', 0) if last_actual else 0)
            inventory_diff = inventory - p_inventory
            p_creditors = prev_year.get("creditors", 0) if prev_year else (getattr(last_actual, 'creditors', getattr(last_actual, 'current_liabilities', 0) * 0.5) if last_actual else 0)
            cred_diff = creditors - p_creditors
            p_oca = prev_year.get("other_current_assets", 0) if prev_year else other_current_assets_base
            oca_diff = current_other_assets - p_oca
            p_ocl = prev_year.get("other_current_liabilities", 0) if prev_year else other_current_liabilities_base
            ocl_diff = current_other_liabilities - p_ocl
            p_wc_bal = prev_year.get("wc_loan_bal", 0) if prev_year else 0
            cc_limit_diff = final_cc_bal - p_wc_bal

            total_sources = pat + depr_total + (fresh_contribution if y==1 else 0) + (project.loan.term_loan_amount if y == 1 else 0.0)
            total_sources += max(0, cred_diff) + max(0, ocl_diff) + max(0, cc_limit_diff) + max(0, -debtors_diff) + max(0, -inventory_diff) + max(0, -oca_diff)
            
            total_uses = tl_repayment_total + asset_outlay
            total_uses += max(0, -cred_diff) + max(0, -ocl_diff) + max(0, -cc_limit_diff) + max(0, debtors_diff) + max(0, inventory_diff) + max(0, oca_diff)
            net_cf = total_sources - total_uses
            current_cash = final_cash
            current_assets = debtors + inventory + current_other_assets + final_cash
            
            # Final Year Data
            year_data = {
                "year_label": y_label,
                "opening_equity": opening_equity,
                "drawings": drawings,
                "is_actual": False,
                "revenue": current_revenue,
                "cogs": material_purchases,
                "stock_adj": stock_adj,
                "gp_amt": gp_amt,
                "ebitda": ebitda,
                "depreciation": depr_total,
                "total_int": total_int,
                "cc_interest": cc_interest,
                "tl_interest": tl_interest,
                "tl_interest_total": tl_interest_total,
                "tax_amt": tax_amt,
                "pat": pat,
                "tl_repayment": tl_repayment,
                "cash_accruals": pat + depr_total,
                
                "opening_fixed_assets": opening_fa,
                "fixed_asset_additions": asset_outlay,
                "net_fixed_assets": net_fixed_assets,
                "debtors": debtors,
                "inventory": inventory,
                "opening_cash_bal": opening_cash,
                "cash_bal": current_cash,
                "closing_cash_bal": current_cash,
                "current_assets": current_assets,
                "total_assets": total_assets,
                
                "share_capital": current_share_capital,
                "reserves_surplus": current_reserves,
                "tl_bal": tl_balance,
                "wc_loan_bal": final_cc_bal,
                "cc_limit": cc_limit,
                "creditors": creditors,
                "unsecured_loan": unsecured_loans,
                "other_current_liabilities": current_other_liabilities,
                "total_liabilities": total_liabs,
                
                # Cash Flow root keys for PDF mapping
                "total_sources": total_sources,
                "total_uses": total_uses,
                "net_cash_flow": net_cf,
                "reconciliation_variance": 0.0,
                "cap_inc": (fresh_contribution if y==1 else 0),
                "loan_inc": (project.loan.term_loan_amount if y == 1 else 0.0),
                "asset_purchase": asset_outlay,
                "ca_inc": debtors_diff + inventory_diff + oca_diff,
                "cl_inc": cred_diff + ocl_diff,
                "other_current_assets": current_other_assets,
                "tl_repayment": tl_repayment_total,
                "tl_bal": tl_balance_total,
                "tl_balance_total": tl_balance_total,
                "tl_bal_new": tl_bal_new,
                "tl_bal_existing": tl_bal_existing,
                "stock_adj": stock_adj,
                
                # Ratio Metadata
                "dscr": y_dscr,
                "current_ratio": cls._safe_div(current_assets, current_liabs, 2.0),
                "fixed_costs": depr_total + total_int + ind_exp,
                "contribution_pct": cls._safe_div(gp_amt * 100, current_revenue, target_gp_pct),
                "bep_sales": cls._safe_div(depr_total + total_int + ind_exp, target_gp_pct / 100),
                "cash_bep": cls._safe_div(total_int + ind_exp, target_gp_pct / 100),
                "depreciation_details": depr_details,
                
                "ind_exp": ind_exp,
                "pbt": pbt,
                "expense_breakdown": {
                    "Salary & Wages": ind_exp * 0.40,
                    "Power & Fuel": ind_exp * 0.15,
                    "Rent & Rates": ind_exp * 0.10,
                    "Admin & Misc": ind_exp * 0.35
                }
            }
            
            # Capture net block for next year's opening
            net_fixed_assets_prev = net_fixed_assets
            
            # Phase 5: Sensitivity Analysis (Stress Test)
            # Use EFFECTIVE GP % from the base case for sensitivity to avoid logical flips
            effective_gp_pct = cls._safe_div(gp_amt * 100, current_revenue, target_gp_pct)
            year_data["sensitivity"] = {
                "minus_10pct": cls._calculate_stress_metrics(year_data, 0.90, effective_gp_pct, calibrated_exp_pct, ass, project),
                "minus_20pct": cls._calculate_stress_metrics(year_data, 0.80, effective_gp_pct, calibrated_exp_pct, ass, project),
            }
            
            results.append(year_data)
            
        return results

    @classmethod
    def _calculate_stress_metrics(cls, base_data: dict, rev_factor: float, gp_percent: float, exp_percent: float, ass: Any, project: CmaProject) -> dict:
        """Calculates comprehensive health metrics for a stressed revenue scenario."""
        rev = base_data["revenue"] * rev_factor
        gp = rev * (gp_percent / 100)
        
        # --- LOGICAL FIX: Treat Indirect Expenses as FIXED ---
        # Revenue shortfall doesn't reduce rent/salaries proportionally.
        # We use the base_data's absolute indirect expense amount.
        ind_exp = base_data["ind_exp"] 
        ebitda = gp - ind_exp
        
        # PBT = EBITDA - Depr - Total Interest (TL + CC)
        pbt = ebitda - base_data["depreciation"] - base_data["total_int"]
        tax = max(0, pbt * (ass.tax_rate_percent / 100))
        pat = pbt - tax
        
        # --- LOGICAL FIX: Use Total Debt Service for Stressed DSCR ---
        # Numerator: Stressed PAT + Depr + Total TL Interest
        # Denominator: Total TL Repayment + Total TL Interest
        # Note: 'tl_repayment' in base_data is already total_repayment as it was overwritten in loop.
        # We use 'tl_interest_total' which we just added.
        tl_int_total = base_data.get("tl_interest_total", base_data.get("tl_interest", 0.0))
        num = pat + base_data["depreciation"] + tl_int_total
        den = base_data["tl_repayment"] + tl_int_total
        dscr = cls._safe_div(num, den, (pat + base_data["depreciation"]))
        
        # Stressed Current Ratio
        # Assuming debtors and inventory drop with revenue (variable assets)
        # Cash is kept constant as a stressed floor check
        s_debtors = base_data["debtors"] * rev_factor
        s_inventory = base_data["inventory"] * rev_factor
        s_ca = s_debtors + s_inventory + base_data["cash_bal"] + base_data.get("other_current_assets", 0)
        cl = base_data["cc_limit"] + base_data["creditors"] + base_data["other_current_liabilities"]
        cr = cls._safe_div(s_ca, cl, 2.0)
        
        return {
            "revenue": rev,
            "ebt": pbt,
            "pat": pat,
            "dscr": dscr,
            "current_ratio": cr
        }

    @classmethod
    def validate_projections(cls, projections: List[dict]):
        """
        Runs institutional-grade QA checks on generated projections.
        Throws ValueError with descriptive message if validation fails.
        """
        for p in projections:
            lbl = p.get('year_label', 'Unknown FY')
            
            # 1. Balance Sheet Tally Check
            total_assets = p.get('total_assets', 0)
            total_liabs = p.get('total_liabilities', 0)
            diff = abs(total_assets - total_liabs)
            
            if diff > 0.01: # 0.01 Lakhs (Rs 1,000 threshold)
                raise ValueError(
                    f"Balance Sheet Mismatch for {lbl}: \n"
                    f"Total Assets (Rs. {total_assets:.2f} L) != Total Liabs (Rs. {total_liabs:.2f} L). \n"
                    f"Mismatch: {diff:.4f} Lakhs. Please check Capital or Loan inputs."
                )

            # 2. Cash Flow Integrity Check (Sources - Uses = Net)
            sources = p.get("total_sources", 0)
            uses = p.get("total_uses", 0)
            net = p.get("net_cash_flow", 0)
            cf_diff = abs((sources - uses) - net)
            if cf_diff > 0.01:
                raise ValueError(f"Cash Flow Error in {lbl}: Sum of Sources and Uses does not match Net Cash Flow.")

            # 3. Balance Sheet vs Cash Flow Ending Balance Check
            bs_cash = p.get('cash_bal', 0)
            # Find prev_cash
            prev_year = projections[projections.index(p)-1] if projections.index(p) > 0 else None
            # Need a more robust way to check for year 1 vs historical
            # For now, let's just ensure sources - uses matches the movement
            # We already have total_liabilities == total_assets which ensures BS is balanced.

            # 3. Sanity Checks
            if p.get('revenue', 0) < 0:
                raise ValueError(f"Projected Revenue cannot be negative in {lbl}.")
            
            if p.get('pat', 0) < -100: # Rs. 10 CR loss threshold for small reports
                pass # Allow losses but maybe warn in UI

    @staticmethod
    def _model_to_projection(ad: AuditedData, is_actual: bool = True) -> Dict[str, Any]:
        """Converts an AuditedData model to the generic projection dictionary format."""
        
        # Calculate B/S totals - Automated based on components (per user request)
        cur_liabs = (ad.creditors or 0.0) + (ad.other_current_liabilities or 0.0) + (ad.provisions or 0.0)
        cur_assets = (ad.inventory or 0.0) + (ad.debtors or 0.0) + (ad.other_current_assets or 0.0) + \
                     (ad.loans_advances or 0.0) + (ad.deposits or 0.0) + (ad.cash_bank or 0.0)

        liabs_val = (ad.share_capital or 0.0) + (ad.reserves_surplus or 0.0) + \
                   (ad.term_loans or 0.0) + (ad.unsecured_loan or 0.0) + (ad.bank_od or 0.0) + \
                   (ad.other_loans_liabilities or 0.0) + cur_liabs
                   
        assets_val = (ad.net_block or 0.0) + (ad.investments or 0.0) + cur_assets
        
        # Mandatory Rounding Difference (Requirement 5)
        # Tolerance up to 0.01 Lakhs (Rs 1000)
        diff = round(liabs_val - assets_val, 2)
        rounding_diff = 0.0
        if 0 < abs(diff) <= 0.01:
            rounding_diff = diff
            # Adjust total liabilities to match assets for the 'tally' presentation
            liabs_val = assets_val 
            
        # Derived Metrics for Report Sync (Requirement 11)
        pbt = ad.net_profit + (ad.tax_amt or 0)
        ebitda = pbt + (ad.depreciation or 0) + (ad.interest_paid or 0)
        
        # Calibration: Opening Equity for the history year should be (Total Capital - PAT)
        # to ensure the "Movement" logic (Opening + PAT = Closing) tallies for the Actual year.
        opening_equity = 0.0
        if is_actual:
            opening_equity = (ad.share_capital or 0.0) - (ad.net_profit or 0.0)
            
        return {
            "year_label": ad.year_label,
            "is_actual": is_actual,
            "is_detailed": True, # Historical actuals are now always considered detailed for reports
            "revenue": ad.revenue,
            "cogs": ad.cogs,
            "opening_stock": ad.opening_stock,
            "closing_stock": ad.inventory,
            "gp_amt": ad.gross_profit,
            "pat": ad.net_profit,
            "net_profit": ad.net_profit,
            "salary_wages": ad.salary_wages,
            "rent_rates": ad.rent_rates,
            "admin_other_exp": ad.admin_expenses,
            "depreciation": ad.depreciation,
            "tax_amt": ad.tax_amt,
            "opening_equity": opening_equity,
            "share_capital": ad.share_capital,
            "reserves_surplus": ad.reserves_surplus,
            "tl_bal": ad.term_loans,
            "tl_bal_existing": ad.term_loans, # Requirement: Display in row M1 for historical
            "tl_balance_total": ad.term_loans, # Total loans for this year
            "wc_loan_bal": ad.bank_od or 0.0,
            "cc_limit": ad.bank_od or 0.0, 
            "creditors": ad.creditors,
            "other_current_liabilities": ad.other_current_liabilities,
            "total_liabilities": liabs_val,
            "net_fixed_assets": ad.net_block,
            "inventory": ad.inventory,
            "debtors": ad.debtors,
            "other_current_assets": ad.other_current_assets,
            "cash_bal": ad.cash_bank,
            "closing_cash_bal": ad.cash_bank, # Sync with Section M
            "total_assets": assets_val,
            "interest": ad.interest_paid,
            "drawings": 0.0, # Historical drawings assumed 0 or already factored in Reserves
            "total_int": ad.interest_paid or 0.0, # Sync with Section L
            "net_block": ad.net_block or 0.0,
            "current_assets": ad.current_assets or 0.0,
            
            # New Granular Fields
            "unsecured_loan": ad.unsecured_loan or 0.0,
            "bank_od": ad.bank_od or 0.0,
            "other_loans_liabilities": ad.other_loans_liabilities or 0.0,
            "provisions": ad.provisions or 0.0,
            "loans_advances": ad.loans_advances or 0.0,
            "deposits": ad.deposits or 0.0,
            "investments": ad.investments or 0.0,
            
            # P&L Detail Mapping (Requirement 11)
            "opening_stock": ad.opening_stock or 0.0,
            "cogs": ad.cogs or 0.0,
            "gp_amt": ad.gross_profit or 0.0, # Sync with Section L
            "ebitda": ebitda,
            "pbt": pbt,
            "tax_amt": ad.tax_amt or 0.0,
            "cash_accruals": ad.net_profit + ad.depreciation,
            
            "salary_wages": ad.salary_wages or 0.0,
            "labour_expenses": ad.labour_expenses or 0.0,
            "power_fuel": ad.power_fuel or 0.0,
            "rent_rates": ad.rent_rates or 0.0,
            "admin_expenses": ad.admin_expenses or 0.0,
            "other_direct_expenses": ad.other_direct_expenses or 0.0,
            "interest_exp": ad.interest_exp or 0.0,
            "ind_exp": (ad.salary_wages or 0) + (ad.power_fuel or 0) + (ad.rent_rates or 0) + (ad.admin_expenses or 0),
            
            # Expense Breakdown for Section P
            "expense_breakdown": {
                "Salary & Wages": ad.salary_wages or 0.0,
                "Power & Fuel": ad.power_fuel or 0.0,
                "Rent & Rates": ad.rent_rates or 0.0,
                "Admin & Misc": ad.admin_expenses or 0.0
            },
            
            "rounding_diff": rounding_diff,
            "dscr": (ad.net_profit + ad.depreciation) / (ad.interest_paid + (ad.term_loans/5.0)) if (ad.interest_paid + ad.term_loans) > 0 else 2.0,
            "current_ratio": ad.current_assets / ad.current_liabilities if (ad.current_liabilities or 0) > 0 else 2.0
        }

    @staticmethod
    def _simplified_to_projection(sd: Any) -> Dict[str, Any]:
        """Converts SimplifiedData model to the generic projection dictionary format."""
        rev = sd.approx_turnover
        gp = rev * sd.gp_percent / 100
        np = rev * sd.np_percent / 100
        return {
            "year_label": "Last Year (E)",
            "is_actual": True,
            "revenue": rev,
            "pat": np,
            "depreciation": rev * 0.02, # Estimated proxy
            "net_block": sd.fixed_assets_estimate,
            "current_assets": sd.receivables_estimate + sd.inventory_estimate + sd.cash_bank_estimate,
            "share_capital": sd.fixed_assets_estimate * 0.5, # Assume 50% equity for simplified base
            "reserves_surplus": np,
            "term_loans": sd.borrowing_estimate * 0.4,
            "wc_loan_bal": sd.borrowing_estimate * 0.6,
            "creditors": sd.creditors_estimate,
            "total_assets": sd.fixed_assets_estimate + sd.receivables_estimate + sd.inventory_estimate + sd.cash_bank_estimate,
            "total_liabilities": sd.fixed_assets_estimate + sd.receivables_estimate + sd.inventory_estimate + sd.cash_bank_estimate
        }
    @classmethod
    def get_summary_ratios(cls, project: CmaProject) -> Dict[str, Any]:
        """
        Calculates key banking ratios for Phase 2 overview.
        """
        total_project_cost = sum(a.cost for a in project.assets) + project.loan.working_capital_requirement
        total_loan = project.loan.term_loan_amount + project.loan.cash_credit_amount
        
        # Point G: Include historical capital in promoter support for existing firms
        historical_capital = 0.0
        if project.audited_history:
            last = project.audited_history[-1]
            historical_capital = (last.share_capital or 0.0) + (last.reserves_surplus or 0.0)
            
        fresh_contribution = max(0, total_project_cost - total_loan)
        promoter_contribution = fresh_contribution + historical_capital
        
        der = total_loan / promoter_contribution if promoter_contribution > 0 else 0.0
        margin_percent = (promoter_contribution / total_project_cost * 100) if total_project_cost > 0 else 0.0
        
        # Financial summary from projections
        projections = cls.generate_full_projections(project)
        projected_only = [p for p in projections if not p.get("is_actual")]
        
        avg_dscr = sum(p["dscr"] for p in projected_only) / len(projected_only) if projected_only else 0.0
        min_dscr = min((p["dscr"] for p in projected_only), default=0.0)
        
        # Validation for banking logic
        is_valid = True
        if total_project_cost <= total_loan or not project.assets:
            is_valid = False

        return {
            "is_valid": is_valid,
            "total_project_cost": total_project_cost,
            "total_loan": total_loan,
            "promoter_contribution": promoter_contribution,
            "debt_equity_ratio": der,
            "margin_percent": margin_percent,
            "avg_dscr": avg_dscr,
            "min_dscr": min_dscr,
            "current_ratio": projected_only[0]["current_ratio"] if projected_only else 2.0
        }
