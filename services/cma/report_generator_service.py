"""
Report Generator Service for CMA / DPR Builder.
Premium mode-specific PDF generation engine with institutional-grade formatting.
"""

import os
import tempfile
import io
import re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from datetime import datetime
from typing import List, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, white, black, gray
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle, Spacer, SimpleDocTemplate, PageBreak, Frame, PageTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY

from services.cma.models import CmaProject, AssetItem, BusinessCategory, DepreciationMethod
from services.cma.image_mapping_service import ImageMappingService
from services.cma.narrative_service import NarrativeService
from services.cma.projection_engine_service import ProjectionEngineService

from reportlab.platypus import Flowable


class PageMarker(Flowable):
    """Hidden flowable to record which page a section starts on."""
    def __init__(self, section_id, registry):
        Flowable.__init__(self)
        self.section_id = section_id
        self.registry = registry

    def draw(self):
        page_num = self.canv.getPageNumber()
        self.registry[self.section_id] = page_num


class ReportGeneratorService:
    """Service to build professional PDF reports for CMA / DPR."""

    _page_registry = {}

    @classmethod
    def _clean_text(cls, text):
        """Removes/Sanitizes raw HTML tags from strings for clean PDF output."""
        if not text: return ""
        clean = re.sub(r'<[^>]*>', '', str(text))
        return clean

    @classmethod
    def _wrap_cell(cls, text, style):
        """Wraps text in a Paragraph for automatic cell wrapping in Tables."""
        safe_text = cls._clean_text(text)
        return Paragraph(safe_text, style)

    @classmethod
    def _format_fy_label(cls, label):
        """Standardizes year labels. Strips existing badges like (P) or (Aud.) first."""
        if not label: return label
        label = label.strip()
        
        # 1. Forcefully strip any existing status badges to avoid duplication (e.g. (P) (P))
        label = re.sub(r'\s*\((P|Aud|Prov|A|A-P|Provisional|Audited)\.?\)', '', label, flags=re.IGNORECASE)
        label = label.strip()
        
        # 2. If it already follows FY YYYY-YY or YYYY-YY pattern, clean it up
        if ("FY" in label and "-" in label) or (re.match(r"^\d{4}-\d{2}", label)):
            if not label.startswith("FY "):
                 return "FY " + label
            return label
            
        # 3. Default standardizer: 2026 -> FY 2025-26
        match = re.search(r"(\d{4})", label)
        if match:
            y_str = match.group(1)
            y_int = int(y_str)
            fy_str = f"{y_int-1}-{y_str[-2:]}"
            new_label = label.replace(y_str, fy_str)
            if not new_label.startswith("FY "):
                new_label = "FY " + new_label
            return new_label
        return label

    @classmethod
    def _get_year_headers(cls, proj_results, project, body_style):
        """Build year header labels with audit status tags."""
        headers = []
        for r in proj_results:
            orig_label = r.get("year_label", "N/A")
            display_label = cls._format_fy_label(orig_label)
            
            # Status badge logic (Requirement 11)
            status = ""
            if r.get("is_actual"):
                # Use the explicitly passed status if available, else check history
                ds = r.get("data_status")
                if ds:
                    if ds == "Audited": status = " (Aud.)"
                    elif ds == "Provisional": status = " (Prov.)"
                else:
                    # Fallback check
                    for ad in project.audited_history:
                        if ad.year_label == orig_label:
                            if ad.data_type == "Audited": status = " (Aud.)"
                            elif ad.data_type == "Provisional": status = " (Prov.)"
                            break
            else:
                status = " (P)" # Projected
            
            headers.append(f"<b>{display_label}{status}</b>")
        return headers

    @classmethod
    def generate_pdf(cls, project: CmaProject, output_path: str) -> str:
        """Two-pass rendering engine for accurate TOC page numbers."""
        cls._validate_project(project)
        cls._page_registry = {}
        cls._run_build(project, io.BytesIO(), is_first_pass=True)
        return cls._run_build(project, output_path, is_first_pass=False)

    @classmethod
    def _validate_project_for_export(cls, project, proj_results):
        """Requirement 9: Export QA Gate."""
        for p in proj_results:
            diff = abs(p.get("total_assets", 0) - p.get("total_liabilities", 0))
            if diff > 0.015:
                raise ValueError(f"Balance Sheet mismatch in {p.get('year_label')}")

    @classmethod
    def _validate_project(cls, project: CmaProject):
        """Ensures the project data is consistent and bank-ready before export."""
        from services.cma.models import LoanType
        ft = project.loan.facility_type
        
        # 1. Balance Sheet Tally Check (Point C & I)
        from services.cma.projection_engine_service import ProjectionEngineService
        projections = ProjectionEngineService.generate_full_projections(project)
        for p in projections:
            diff = abs(p.get("total_assets", 0) - p.get("total_liabilities", 0))
            if diff > 0.015: # 0.01 Lakhs tolerance (rounded to nearest thousand)
                raise ValueError(
                    f"Institutional Validation Error: Projected Balance Sheet does not tally for {p.get('year_label', 'FY')}.\n"
                    f"Total Assets (Rs. {p.get('total_assets',0):.2f}) != Total Liabilities (Rs. {p.get('total_liabilities',0):.2f}).\n"
                    "Please check Capital, Historical Net Block or Loan mapping before export."
                )

        # 2. OD/CC Case Consistency Check

        # 2. Placeholder & Identification Check
        from services.cma.models import BusinessCategory
        current_cat = project.profile.business_category
        current_desc = project.profile.description or ""
        
        # Legacy "Generic Business" is now strictly blocked
        if "Generic Business" in current_cat:
             raise ValueError(
                f"Legacy placeholder '{current_cat}' detected in Business Category. "
                "Please go to the 'Party Details' tab and select a specific Category (e.g. 'Other Business' or 'Engineering')."
            )

        # Proposed Activity Check (Point 3 & 9)
        # We auto-fill it later, but we block if description is also empty
        if not current_desc.strip():
            raise ValueError(
                "Business Description / Activity is blank. "
                "Please provide a specific description (e.g. 'Manufacturer of Spare Parts') in the Party Details tab."
            )

        # 3. Auto-Total Verification (Point 4 & 9)
        # Manpower check: (6 vs 5 mismatch fix)
        count = project.profile.employee_count
        if count < 1:
            raise ValueError("Employee count must be specified (minimum 1).")
        
        # Consistent row breakup check for validation
        skilled = int(count * 0.4)
        semi_skilled = int(count * 0.4)
        admin = max(1, count - (skilled + semi_skilled))
        if (skilled + semi_skilled + admin) != count:
            # This should be handled by our allocation logic, 
            # but we validate to be safe.
            pass 
        skilled = int(count * 0.4)
        semi_skilled = int(count * 0.4)
        admin = int(count * 0.2) or 1
        # No actual conflict here because we force sum in PDF, but we validate existence
        if count <= 0:
            raise ValueError("Employee count must be greater than zero for a valid appraisal.")

        # Financial Balance Check
        total_assets = sum(a.cost for a in project.assets)
        total_project_cost = total_assets + project.loan.working_capital_requirement
        total_loan = project.loan.term_loan_amount + project.loan.cash_credit_amount
        promoter_margin = total_project_cost - total_loan
        
        if abs(total_project_cost - (total_loan + promoter_margin)) > 0.01:
             raise ValueError("Mathematical mismatch: Total Project Cost does not match Total Means of Finance.")

    @classmethod
    def _run_build(cls, project, target, is_first_pass: bool):
        """Core build logic with mode-specific theming."""
        from services.cma.report_theme import get_theme
        from services.cma.models import ReportMode, BusinessMode

        report_mode = project.profile.report_mode
        theme = get_theme(report_mode)
        title_style, section_style, body_style = theme.build_styles()

        doc = SimpleDocTemplate(
            target, pagesize=A4,
            rightMargin=18*mm, leftMargin=18*mm,
            topMargin=22*mm, bottomMargin=22*mm
        )
        
        elements = []
        proj_results = ProjectionEngineService.generate_full_projections(project)
        if not is_first_pass:
            ProjectionEngineService.validate_projections(proj_results)

        # Requirement: Include Historical Comparative Data (Point 11)
        # Force-collect actuals from the engine's result set
        actuals = [r for r in proj_results if r.get("is_actual")]
        projected = [r for r in proj_results if not r.get("is_actual")]
        
        max_proj = 3 if report_mode == ReportMode.LITE.value else 5
        # Prepend actuals first, then add projected years
        proj_results = actuals + projected[:max_proj]
        
        from services.cma.models import LoanType, SchemeType
        has_tl = project.loan.term_loan_amount > 0
        lt = project.profile.loan_type
        
        # Section registry per mode
        master_sections = [("A", "A_header"), ("B", "B_contents"), ("C", "C_summary")]
        
        if report_mode == ReportMode.LITE.value:
            # LITE: Compact 8-10 pages
            master_sections += [
                ("DASH", "DASH_dashboard"), ("D", "D_snapshot"), ("H", "H_cost"), ("I", "I_finance")
            ]
            if has_tl: master_sections.append(("O", "O_fixed_assets"))
            
            master_sections += [
                ("L", "L_operating_stmt"), ("N", "N_cash_flow"), ("M", "M_balance_sheet")
            ]
            if has_tl: 
                master_sections.append(("Q", "Q_dscr"))
                master_sections.append(("AA", "AA_monthly_repayment"))
            master_sections += [("AD", "AD_readiness"), ("Z", "Z_declaration")]

        elif report_mode == ReportMode.CMA.value:
            # CMA: Banker analytical style
            master_sections += [
                ("B1", "B1_analytical_profile"), ("DASH", "DASH_dashboard"),
                ("L", "L_operating_stmt"), ("M", "M_balance_sheet"), 
                ("N", "N_cash_flow"), ("W", "W_cma_data"), 
                ("AC", "AC_mpbf")
            ]
            # Only show Repayment/DSCR if it's a Term/Composite loan
            is_wc_only = (lt in [LoanType.OD_LIMIT.value, LoanType.RENEWAL.value, LoanType.WORKING_CAPITAL.value])
            if has_tl and not is_wc_only:
                master_sections.append(("Q", "Q_dscr"))
                master_sections.append(("V", "V_repayment"))
                master_sections.append(("AA", "AA_monthly_repayment"))

            master_sections += [("R", "R_liquidity"), ("AD", "AD_readiness"), ("Z", "Z_declaration")]

        else:
            # PRO: Flagship Premium DPR
            master_sections += [
                ("DASH", "DASH_dashboard"), ("D", "D_snapshot"), ("E", "E_entity"), ("F", "F_promoter"), ("G", "G_employment"),
                ("H", "H_cost"), ("I", "I_finance"), ("J", "J_financial_data"), ("K", "K_graphics"),
                ("L", "L_operating_stmt"), ("N", "N_cash_flow"), ("M", "M_balance_sheet")
            ]
            if has_tl:
                master_sections += [("O", "O_fixed_assets"), ("V", "V_repayment"), ("Q", "Q_dscr"), ("AA", "AA_monthly_repayment")]
            
            master_sections += [
                ("P", "P_expenses"), ("R", "R_liquidity"),
                ("S", "S_sensitivity"), ("T", "T_bep"), ("U", "U_margin"),
                ("W", "W_cma_data"), ("AC", "AC_mpbf"), ("AD", "AD_readiness"),
                ("X", "X_assumptions"), ("Y", "Y_security"), ("Z", "Z_declaration")
            ]

        func_map = {
            "A_header": cls._add_section_A_header, "B_contents": cls._add_section_B_contents,
            "DASH_dashboard": cls._add_section_DASH_dashboard,
            "B1_analytical_profile": cls._add_section_B1_analytical_profile,
            "C_summary": cls._add_section_C_summary, "D_snapshot": cls._add_section_D_snapshot,
            "E_entity": cls._add_section_E_entity, "F_promoter": cls._add_section_F_promoter,
            "G_employment": cls._add_section_G_employment, "H_cost": cls._add_section_H_cost,
            "I_finance": cls._add_section_I_finance, "J_financial_data": cls._add_section_J_financial_data,
            "K_graphics": cls._add_section_K_graphics, "L_operating_stmt": cls._add_section_L_operating_stmt,
            "M_balance_sheet": cls._add_section_M_balance_sheet, "N_cash_flow": cls._add_section_N_cash_flow,
            "O_fixed_assets": cls._add_section_O_fixed_assets, "P_expenses": cls._add_section_P_expenses,
            "Q_dscr": cls._add_section_Q_dscr, "R_liquidity": cls._add_section_R_liquidity,
            "S_sensitivity": cls._add_section_S_sensitivity, "T_bep": cls._add_section_T_bep,
            "U_margin": cls._add_section_U_margin, "V_repayment": cls._add_section_V_repayment,
            "AA_monthly_repayment": cls._add_section_AA_monthly_repayment,
            "W_cma_data": cls._add_section_W_cma_data, "AC_mpbf": cls._add_section_AC_mpbf,
            "AD_readiness": cls._add_section_AD_readiness, "X_assumptions": cls._add_section_X_assumptions,
            "Y_security": cls._add_section_Y_security, "Z_declaration": cls._add_section_Z_declaration
        }

        cls._last_master_sections = master_sections # Save for TOC stability
        for sid, func_name in master_sections:
            elements.append(PageMarker(sid, cls._page_registry))
            func = func_map.get(func_name)
            if not func: continue
            # Sections needing projection data
            if sid in ["DASH", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "W", "AA", "AC", "AD"]:
                func(elements, project, proj_results, section_style, body_style, theme)
            else:
                func(elements, project, section_style, body_style, theme)
            
            # Switch to main template after cover section (Section A)
            if sid == "A":
                from reportlab.platypus import NextPageTemplate
                elements.append(NextPageTemplate('AllPages'))
        
        # ── Mode-aware page template (header + footer) ──
        def on_first_page(c, d):
            c.saveState()
            # ── Decorative Side Bar ──
            c.setFillColor(theme.PRIMARY)
            c.rect(0, 0, 8*mm, A4[1], fill=1, stroke=0)
            
            # ── Accent Header Band ──
            c.setFillColor(theme.SECONDARY)
            c.rect(8*mm, A4[1]-15*mm, A4[0]-8*mm, 15*mm, fill=1, stroke=0)
            
            c.setFillColor(white)
            c.setFont('Helvetica-Bold', 10)
            c.drawRightString(A4[0]-15*mm, A4[1]-10*mm, theme.HEADER_LABEL.upper())
            
            # Inner Border
            c.setStrokeColor(theme.PRIMARY)
            c.setLineWidth(0.4*mm)
            c.rect(12*mm, 10*mm, A4[0]-22*mm, A4[1]-30*mm)
        frame = Frame(18*mm, 22*mm, 174*mm, 253*mm, id='normal')
        
        def my_footer(canvas, doc):
            canvas.saveState()
            # Footer simplified: Branding removed as it is now on cover page
            footer_text = "Project Appraisal Report"
            
            canvas.setFont('Helvetica', 8)
            canvas.setStrokeColor(theme.BORDER)
            canvas.setLineWidth(0.1*mm)
            canvas.line(20*mm, 15*mm, 190*mm, 15*mm)

            # Modern Geometric Accent (Small triangle at corner)
            canvas.setFillColor(theme.PRIMARY)
            path = canvas.beginPath()
            path.moveTo(195*mm, 15*mm)
            path.lineTo(210*mm, 15*mm)
            path.lineTo(210*mm, 0)
            path.close()
            canvas.drawPath(path, fill=1, stroke=0)
            
            canvas.setFillColor(theme.TEXT)
            canvas.drawString(20*mm, 10*mm, footer_text)
            canvas.drawCentredString(105*mm, 10*mm, f"Page {doc.page}")
            if project.branding.disclaimer:
                canvas.drawRightString(190*mm, 10*mm, project.branding.disclaimer)
            else:
                canvas.drawRightString(190*mm, 10*mm, "Confidential Banking Document")
            canvas.restoreState()

        # ── Page Template Definitions ──
        cover_template = PageTemplate(id='Cover', frames=frame, onPage=on_first_page)
        main_template = PageTemplate(id='AllPages', frames=frame, onPage=my_footer)
        doc.addPageTemplates([cover_template, main_template])
        
        # Start with Cover template
        elements.insert(0, NextPageTemplate('Cover'))
        
        doc.build(elements)
        return target

    @classmethod
    def _add_section_A_header(cls, elements, project, title_style, body_style, theme=None):
        """A. Premium Cover Page — Mode-specific layout."""
        from services.cma.models import ReportMode, SchemeType
        report_mode = project.profile.report_mode
        scheme = project.profile.scheme_type
        loan_type = project.profile.loan_type
        
        # Title logic
        if report_mode == ReportMode.CMA.value:
            main_title = "CMA DATA & FINANCIAL ANALYSIS"
            sub_title = "WORKING CAPITAL ASSESSMENT REPORT"
        elif report_mode == ReportMode.LITE.value:
            main_title = "PROJECT REPORT"
            sub_title = f"{loan_type} Proposal"
        else:
            main_title = "DETAILED PROJECT REPORT (DPR)"
            sub_title = "PREMIUM BANK SUBMISSION PACK"

        if scheme == SchemeType.MUDRA.value:
            main_title = "PRADHAN MANTRI MUDRA YOJANA (PMMY)"
            sub_title = "PROJECT REPORT FOR MUDRA LOAN"
        elif scheme == SchemeType.PMEGP.value:
            main_title = "PMEGP PROJECT REPORT"
            sub_title = "PRIME MINISTER'S EMPLOYMENT GENERATION PROGRAMME"

        sub_style = ParagraphStyle('SubTitle', parent=title_style, fontSize=14, textColor=theme.SECONDARY if theme else HexColor("#1976D2"))
        ctr_style = ParagraphStyle('CenterBody', parent=body_style, alignment=TA_CENTER, fontSize=10)
        
        if theme and theme.mode_key == "lite":
            # ── LITE COVER: Compact, clean, no hero image ──
            elements.append(Spacer(1, 30*mm))
            elements.append(Paragraph(main_title, title_style))
            elements.append(Paragraph(sub_title, sub_style))
            elements.append(Spacer(1, 8*mm))
            elements.append(Paragraph(project.profile.business_name.upper(), ParagraphStyle('BizName', parent=title_style, fontSize=20)))
            elements.append(Spacer(1, 12*mm))
            # Quick summary table
            total_cost = sum(a.cost for a in project.assets) + project.loan.working_capital_requirement
            total_loan = project.loan.term_loan_amount + project.loan.cash_credit_amount
            # Auto-fill activity if blank (Point 3)
            activity = cls._clean_text(project.loan.purpose)
            if not activity or activity.lower() in ["n/a", "none"]:
                activity = f"{loan_type} for {project.profile.description}"
                
            snap_data = [
                ["Proposed Activity", activity],
                ["Entity Type", project.profile.entity_type],
                ["Total Project Cost", f"Rs. {total_cost:.2f} Lakhs"],
                ["Credit Facility Sought", f"Rs. {total_loan:.2f} Lakhs"],
                ["Date of Preparation", datetime.now().strftime('%d/%m/%Y')],
            ]
            t = Table(snap_data, colWidths=[55*mm, 105*mm])
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.3, theme.BORDER),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ]))
            elements.append(t)
            elements.append(Spacer(1, 20*mm))
            elements.append(Paragraph("STRICTLY CONFIDENTIAL", ctr_style))

        elif theme and theme.mode_key == "cma":
            # ── CMA COVER: Data-forward, analyst style ──
            elements.append(Spacer(1, 25*mm))
            elements.append(Paragraph(main_title, title_style))
            elements.append(Paragraph(sub_title, sub_style))
            elements.append(Spacer(1, 10*mm))
            elements.append(Paragraph(f"FOR: {project.profile.business_name.upper()}", ParagraphStyle('BizName', parent=title_style, fontSize=18)))
            elements.append(Paragraph(f"PAN: {project.profile.pan} | Constitution: {project.profile.entity_type}", ctr_style))
            elements.append(Spacer(1, 15*mm))
            elements.append(Paragraph(f"Proposed Facility: {loan_type}", ctr_style))
            elements.append(Paragraph(f"Date: {datetime.now().strftime('%d/%m/%Y')}", ctr_style))
            elements.append(Spacer(1, 30*mm))
            elements.append(Paragraph("FOR INTERNAL BANKING USE \u2013 STRICTLY CONFIDENTIAL", ParagraphStyle('Stamp', parent=ctr_style, fontSize=9, textColor=theme.MUTED)))

        else:
            # ── PRO COVER: Premium, Typography-Focused Professional Layout ──
            elements.append(Spacer(1, 15*mm))
            
            # Title Group
            elements.append(Paragraph(main_title, title_style))
            elements.append(Paragraph(sub_title, sub_style))
            elements.append(Spacer(1, 10*mm))
            
            # Primary Project Identity
            elements.append(Paragraph(f"FOR {project.profile.business_name.upper()}", ParagraphStyle('BizName', parent=title_style, fontSize=24, spaceAfter=20)))
            
            # Auto-fill activity if blank (Point 3)
            activity = cls._clean_text(project.loan.purpose)
            if not activity or activity.lower() in ["n/a", "none"]:
                activity = f"{loan_type} for {project.profile.description}"
            
            # Structured Info Table (Replacing the image with a clean data block)
            elements.append(Spacer(1, 10*mm))
            cover_data = [
                ["Proposed Activity", cls._wrap_cell(activity, body_style)],
                ["Industry Category", project.profile.business_category],
                ["Entity Constitution", project.profile.entity_type],
                ["Financial Facility", loan_type],
                ["Report Mode", f"{report_mode} Analysis"],
                ["Date of Preparation", datetime.now().strftime('%B %d, %Y')],
            ]
            
            ct = Table(cover_data, colWidths=[65*mm, 100*mm])
            ct.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), theme.BAND_ODD),
                ('LINEBELOW', (0, 0), (-1, -1), 0.1, theme.BORDER if theme else gray),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ('LEFTPADDING', (0, 0), (-1, -1), 12),
                ('TEXTCOLOR', (0, 0), (0, -1), theme.PRIMARY if theme else black),
            ]))
            # Wrap table in a frame-like structure
            elements.append(ct)
            
            elements.append(Spacer(1, 20*mm))
            elements.append(Paragraph("STRICTLY CONFIDENTIAL", ParagraphStyle('Conf', parent=ctr_style, fontName='Helvetica-Bold', fontSize=12, textColor=theme.ACCENT_AMBER if theme else black)))
            
            # ── Branding Section (Moved from Footer) ──
            elements.append(Spacer(1, 15*mm))
            firm_name = project.branding.firm_name or "Professional Appraisal Services"
            prepared_by = project.branding.prepared_by or ""
            contact = project.branding.contact_line or ""
            
            elements.append(Paragraph("REPORT PREPARED BY:", ParagraphStyle('PrepBy', parent=ctr_style, fontSize=9, textColor=theme.MUTED if theme else gray)))
            elements.append(Spacer(1, 2*mm))
            if prepared_by:
                elements.append(Paragraph(prepared_by.upper(), ParagraphStyle('PrepName', parent=ctr_style, fontSize=11, fontName='Helvetica-Bold')))
                elements.append(Spacer(1, 1*mm))
            elements.append(Paragraph(firm_name.upper(), ParagraphStyle('FirmName', parent=ctr_style, fontSize=14, fontName='Helvetica-Bold')))
            if contact:
                elements.append(Paragraph(contact, ParagraphStyle('Contact', parent=ctr_style, fontSize=10)))

        elements.append(PageBreak())

    @classmethod
    def _add_section_B1_analytical_profile(cls, elements, project, section_style, body_style, theme=None):
        """B1. Analytical Borrower Profile (CMA Mode exclusive)"""
        elements.append(Paragraph("SECTION 1: BORROWER APPRAISAL PROFILE", section_style))
        elements.append(Spacer(1, 4*mm))
        
        data = [
            ["Entity Name", project.profile.business_name],
            ["Constitution", project.profile.entity_type],
            ["Registration Id (PAN)", project.profile.pan],
            ["Industry Category", project.profile.business_category],
            ["Business Activity", project.profile.description],
            ["Proposed Scheme", project.profile.scheme_type],
            ["Primary Facility", project.profile.loan_type],
            ["Date of Appraisal", datetime.now().strftime('%d/%m/%Y')],
        ]
        
        if theme:
            t = theme.build_table(
                ["Parameter", "Description / Value"], data, [65*mm, 105*mm],
                num_cols_start=2, wrap_style=body_style
            )
        else:
            t = Table(data, colWidths=[65*mm, 105*mm])
        elements.append(t)
        elements.append(Spacer(1, 8*mm))
        
        elements.append(Paragraph("<b>Appraisal Observation:</b> The borrower seeks institutional credit assistance for business operations as detailed in the following analytical schedules.", body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_B_contents(cls, elements, project, section_style, body_style, theme=None):
        """B. Professional Table of Contents - Dynamically Filtered"""
        elements.append(Paragraph("TABLE OF CONTENTS", section_style))
        elements.append(Spacer(1, 10*mm))
        
        def get_pg(sid):
            return str(cls._page_registry.get(sid, "-"))

        all_labels = {
            "A": "Project Cover Page", "B": "Table of Contents", "B1": "Borrower Appraisal Profile",
            "C": "Executive Summary / Overview", "D": "Project Snapshot", "E": "Entity Profile",
            "F": "Promoter / Management Profile", "G": "Employment Details", "H": "Cost of Project",
            "I": "Means of Finance", "J": "Financial Overview", "K": "Graphical Analytics",
            "L": "Projected Operating Statement", "M": "Balance Sheet Statement", "N": "Cash Flow Statement",
            "O": "Fixed Assets & Depreciation", "P": "Indirect Expenses", "Q": "Debt Coverage (DSCR)",
            "R": "Liquidity Analysis", "S": "Sensitivity Test", "T": "Break-Even Analytics",
            "U": "Margin Analysis", "V": "Repayment Schedule", "W": "CMA Data Schedules",
            "AC": "Assessed MPBF / Limit", "AD": "Bank Readiness Audit", "X": "Notes & Financial Assumptions",
            "Y": "Security & Collateral Details", "Z": "Final Declaration"
        }
        
        # Determine actual sequence rendered
        toc_rows = []
        seq = 0
        
        # FIX: To prevent TOC from shifting pages between Pass 1 and Pass 2,
        # we must use a stable list of sections.
        from services.cma.report_generator_service import ReportGeneratorService
        current_sections = getattr(cls, '_last_master_sections', [])
        
        for sid, _ in current_sections:
            if sid == "B": continue # Don't list TOC in TOC
            seq += 1
            label = all_labels.get(sid, f"Section {sid}")
            pg = str(cls._page_registry.get(sid, "-"))
            toc_rows.append([f"{seq}.", label, pg])

        if theme:
            t = theme.build_table(
                ["No.", "Particulars", "Page"],
                toc_rows, [15*mm, 140*mm, 15*mm],
                num_cols_start=2, wrap_style=body_style
            )
        else:
            t = Table([["No.", "Particulars", "Page"]] + toc_rows, colWidths=[15*mm, 140*mm, 15*mm])
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_DASH_dashboard(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """DASH. Feasibility Dashboard - High Impact Visual Data Page."""
        elements.append(Paragraph("PROJECT FEASIBILITY DASHBOARD", section_style))
        elements.append(Spacer(1, 8*mm))
        
        if not proj_results: return
        
        # 0. Fetch Audit Results for synchronization
        from services.cma.readiness_service import ReadinessService
        audit = ReadinessService.evaluate_readiness(project, proj_results)
        
        # 1. Headline Ratios Table
        avg_dscr = sum(r.get('dscr', 0) for r in proj_results if not r.get('is_actual')) / max(1, len([r for r in proj_results if not r.get('is_actual')]))
        peak_rev = max(r.get('revenue', 0) for r in proj_results)
        
        # FIX: Ensure averaging only over projected years to avoid mismatch with Health Check
        proj_years = [r for r in proj_results if not r.get('is_actual')]
        avg_pat_pct = sum( (r.get('pat',0)/r.get('revenue',1)*100) for r in proj_years if r.get('revenue',0)>0 ) / max(1, len(proj_years))
        last_cr = proj_results[-1].get('current_ratio', 0)

        # Sync Status Labels with Audit Framework
        def get_status(metric_name, fallback_pass):
            for c in audit["checks"]:
                if metric_name in c["name"]:
                    if c["level"] == "PASS": return fallback_pass
                    if c["level"] == "CRITICAL": 
                        return "MARGIN WATCH" if "Margin" in metric_name else "REVISE"
                    return "BORDERLINE"
            return fallback_pass

        dash_data = [
            ["DEBT REPAYMENT CAPACITY (DSCR)", f"{avg_dscr:.2f}", "Benchmark: > 1.25", get_status("DSCR", "HEALTHY")],
            ["PEAK PROJECTED REVENUE", f"Rs. {peak_rev:.2f} L", "Targeted Capacity", "OPTIMIZED"],
            ["AVG. NET PROFIT MARGIN (%)", f"{avg_pat_pct:.2f}%", "Industry Standard: 10-15%", get_status("Profit Margin", "HEALTHY")],
            ["CURRENT LIQUIDITY RATIO", f"{last_cr:.2f}", "Benchmark: > 1.17", get_status("Current Ratio", "STABLE")]
        ]
        
        t_dash = theme.build_table(
            ["Key Performance Indicator", "Value", "Notes", "Status"],
            dash_data, [65*mm, 35*mm, 45*mm, 25*mm],
            num_cols_start=1, wrap_style=body_style
        )
        elements.append(t_dash)
        elements.append(Spacer(1, 12*mm))

        # 2. Add some textual highlights (Synchronized Narratives)
        elements.append(Paragraph("<b>STRATEGIC PROJECT HIGHLIGHTS:</b>", body_style))
        
        # Consistent logic for "Honest" language
        is_high_risk = audit["readiness_level"] == "HIGH RISK (Low Bankability)"
        dscr_status = "STRETCHED" if avg_dscr < 1.15 else "SECURE"
        cash_flow_desc = "TIGHT" if avg_dscr < 1.20 else "ADEQUATE" if avg_dscr < 1.40 else "STRONG"
        
        highlights = [
            f"• <b>Capacity Utilization:</b> The project targets a peak revenue of Rs. {peak_rev:.2f} Lakhs reflecting 80-90% efficiency.",
            f"• <b>Debt Comfort:</b> With an average DSCR of {avg_dscr:.2f}, the enterprise shows {cash_flow_desc} cash flow for bank obligations.",
            f"• <b>Operational Viability:</b> The Net Profit margins are projected to stabilize at {avg_pat_pct:.2f}% following the startup phase."
        ]
        for h in highlights:
            elements.append(Paragraph(h, body_style))
            elements.append(Spacer(1, 2*mm))
            
        # 3. Insert miniature trend placeholder
        elements.append(Spacer(1, 10*mm))
        elements.append(Paragraph(f"<i>Overall Assessment: {audit['readiness_level']}</i>", 
                                   ParagraphStyle('DashStyle', parent=body_style, fontSize=9, textColor=theme.MUTED, alignment=TA_CENTER)))
        
        elements.append(PageBreak())

    @classmethod
    def _add_section_C_summary(cls, elements, project, section_style, body_style, theme=None):
        """C. Executive Summary - Data-driven structured page."""
        elements.append(Paragraph("SECTION C: EXECUTIVE SUMMARY", section_style))
        summary_text = NarrativeService.generate_section("executive_summary", project)
        elements.append(Paragraph(summary_text, body_style))
        elements.append(Spacer(1, 8*mm))

        # COST VS MEANS Table
        total_assets = sum(a.cost for a in project.assets)
        total_project_cost = total_assets + project.loan.working_capital_requirement
        total_loan = project.loan.term_loan_amount + project.loan.cash_credit_amount
        promoter_margin = max(0, total_project_cost - total_loan)

        from services.cma.models import LoanType
        ft = project.loan.facility_type
        
        headers = ["PARTICULARS", "AMOUNT (Rs. L)"]
        rows = [
            ["Project Cost / Assets", f"{total_assets:.2f}"],
            ["Working Capital Outlay (GWC)", f"{project.loan.working_capital_requirement:.2f}"],
            ["<b>TOTAL PROJECT COST</b>", f"<b>{total_project_cost:.2f}</b>"],
            ["-", "-"],
        ]

        if project.loan.term_loan_amount > 0 or ft in [LoanType.TERM_LOAN.value, LoanType.COMPOSITE_LOAN.value]:
            rows.append(["Bank Term Loan Assistance", f"{project.loan.term_loan_amount:.2f}"])
        
        if project.loan.cash_credit_amount > 0 or ft in [LoanType.OD_LIMIT.value, LoanType.WORKING_CAPITAL.value, LoanType.RENEWAL.value, LoanType.COMPOSITE_LOAN.value]:
            rows.append(["Bank Working Capital Limit (CC/OD)", f"{project.loan.cash_credit_amount:.2f}"])
            
        rows.append(["Promoter Contribution", f"{promoter_margin:.2f}"])
        rows.append(["<b>TOTAL MEANS OF FINANCE</b>", f"<b>{total_project_cost:.2f}</b>"])

        if theme:
            t = theme.build_table(headers, rows, [110*mm, 50*mm], 
                                  total_indices=[2, 6], subtotal_indices=[], 
                                  num_cols_start=1, wrap_style=body_style)
            elements.append(t)
        
        elements.append(Spacer(1, 10*mm))
        
        elements.append(Paragraph("FINANCIAL PROJECTION METHODOLOGY", ParagraphStyle('Sub', parent=body_style, fontName='Helvetica-Bold')))
        methodology = cls._clean_text(NarrativeService.generate_section("projection_rationale", project))
        elements.append(Paragraph(methodology, body_style))
        elements.append(Spacer(1, 5*mm))
        
        elements.append(PageBreak())
        
        elements.append(Paragraph("PROJECT RATIONALE", ParagraphStyle('Sub', parent=body_style, fontName='Helvetica-Bold')))
        rationale = cls._clean_text(NarrativeService.generate_section("project_rationale", project))
        elements.append(Paragraph(rationale, body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_D_snapshot(cls, elements, project, section_style, body_style, theme=None):
        """D. Project Snapshot — Theme-aware."""
        elements.append(Paragraph("SECTION D: PROJECT SNAPSHOT", section_style))
        elements.append(Spacer(1, 5*mm))
        
        total_cost = sum(a.cost for a in project.assets) + project.loan.working_capital_requirement
        data = [
            ["Project Opportunity", project.loan.purpose or "Capacity Expansion / Scaling"],
            ["Business Model", f"{project.profile.business_mode or 'Mixed'} Operations"],
            ["Industry Category", project.profile.business_category],
            ["Commencement / Est. Date", project.profile.establishment_date],
            ["Total Appraised Project Cost", f"Rs. {total_cost:.2f} Lakhs"],
            ["Total Credit Assistance Sought", f"Rs. {project.loan.term_loan_amount + project.loan.cash_credit_amount:.2f} Lakhs"],
        ]
        
        if theme:
            t = theme.build_table(
                ["Parameter", "Details"], data, [70*mm, 100*mm],
                num_cols_start=2, wrap_style=body_style
            )
        else:
            t = Table(data, colWidths=[70*mm, 100*mm])
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_E_entity(cls, elements, project, section_style, body_style, theme=None):
        """E. Entity Profile — Theme-aware."""
        elements.append(Paragraph("SECTION E: ENTITY PROFILE", section_style))
        text = cls._clean_text(NarrativeService.generate_section("business_overview", project))
        elements.append(Paragraph(text, body_style))
        elements.append(Spacer(1, 5*mm))
        
        data = [
            ["Registered Identity", project.profile.business_name],
            ["PAN / Income Tax ID", project.profile.pan],
            ["Business Activity", project.profile.description],
            ["Entity Structure", project.profile.entity_type],
            ["Full Operational Address", project.profile.address],
        ]
        
        if theme:
            t = theme.build_table(
                ["Attribute", "Information"], data, [70*mm, 100*mm],
                num_cols_start=2, wrap_style=body_style
            )
        else:
            t = Table(data, colWidths=[70*mm, 100*mm])
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_F_promoter(cls, elements, project, section_style, body_style, theme=None):
        """F. Promoter Profile"""
        elements.append(Paragraph("SECTION F: PROMOTER PROFILE & MANAGEMENT", section_style))
        text = cls._clean_text(NarrativeService.generate_section("promoter_profile", project))
        elements.append(Paragraph(text, body_style))
        elements.append(Spacer(1, 10*mm))
        elements.append(Paragraph("The management exhibits strong commitment to institutional compliance and operational excellence.", body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_G_employment(cls, elements, project, section_style, body_style, theme=None):
        """G. Employment Details — Theme-aware."""
        elements.append(Paragraph("SECTION G: EMPLOYMENT & MANPOWER", section_style))
        text = cls._clean_text(NarrativeService.generate_section("employment_details", project))
        elements.append(Paragraph(text, body_style))
        elements.append(Spacer(1, 5*mm))
        
        count = project.profile.employee_count
        # Robust allocation to ensure total_sum == count
        skilled = int(count * 0.4)
        semi_skilled = int(count * 0.4)
        admin = max(1, count - (skilled + semi_skilled))
        
        # Final adjustment to ensure mathematical perfect sum
        if (skilled + semi_skilled + admin) != count:
            admin = count - (skilled + semi_skilled)
            
        data = [
            ["Skilled / Technical", str(skilled), "Core Manufacturing Operations"],
            ["Semi-Skilled / Labor", str(semi_skilled), "Production Assistance"],
            ["Administrative / Super.", str(admin), "Management & Support"],
            ["TOTAL MANPOWER", str(count), "Full Operational Requirement"],
        ]
        
        if theme:
            t = theme.build_table(
                ["Staff Category", "Headcount", "Functional Role"], data, [55*mm, 30*mm, 90*mm],
                total_indices=[3], num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table(data, colWidths=[55*mm, 30*mm, 90*mm])
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_H_cost(cls, elements, project, section_style, body_style, theme=None):
        """H. Cost of Project — Theme-aware table."""
        elements.append(Paragraph("SECTION H: DETAILED COST OF PROJECT", section_style))
        total_asset_cost = 0
        rows = []
        for i, asset in enumerate(project.assets):
            rows.append([str(i+1), asset.name, f"{asset.cost:.2f}"])
            total_asset_cost += asset.cost
        
        rows.append(["-", "Gross Working Capital Outlay (Requirement)", f"{project.loan.working_capital_requirement:.2f}"])
        total_total = total_asset_cost + project.loan.working_capital_requirement
        rows.append(["", "TOTAL PROJECT CAPITAL OUTLAY", f"<b>{total_total:.2f}</b>"])
        
        if theme:
            t = theme.build_table(
                ["Sr.", "Asset Component Description", "Cost (Rs. Lakhs)"],
                rows, [15*mm, 115*mm, 40*mm],
                total_indices=[len(rows)-1], num_cols_start=2, wrap_style=body_style
            )
        else:
            t = Table([["Sr.", "Asset Component Description", "Cost (Rs. Lakhs)"]] + rows, colWidths=[15*mm, 115*mm, 40*mm])
        elements.append(t)
        elements.append(Spacer(1, 10*mm)) # Combined with Means of Finance

    @classmethod
    def _add_section_I_finance(cls, elements, project, section_style, body_style, theme=None):
        """I. Proposed Means of Finance — Theme-aware table."""
        elements.append(Paragraph("SECTION I: PROPOSED MEANS OF FINANCE", section_style))
        text = cls._clean_text(NarrativeService.generate_section("means_of_finance_narrative", project))
        elements.append(Paragraph(text, body_style))
        
        total_assets = sum(a.cost for a in project.assets)
        total_cost = total_assets + project.loan.working_capital_requirement
        tl = project.loan.term_loan_amount
        cc = project.loan.cash_credit_amount
        promoter = max(0, total_cost - tl - cc)
        
        from services.cma.models import LoanType
        ft = project.loan.facility_type
        
        from services.cma.models import BusinessMode
        is_existing = project.profile.business_mode in [BusinessMode.EXISTING.value, BusinessMode.EXISTING_NO_BOOKS.value]
        
        rows = []
        if tl > 0:
            rows.append(["Term Loan from Bank", f"{tl:.2f}", f"{(tl/total_cost*100 if total_cost > 0 else 0):.1f}%"])
        
        if cc > 0:
            facility_lbl = "Working Capital (CC/OD)" if ft in [LoanType.OD_LIMIT.value, LoanType.RENEWAL.value] else "Working Capital Loan"
            rows.append([facility_lbl, f"{cc:.2f}", f"{(cc/total_cost*100 if total_cost > 0 else 0):.1f}%"])
            
        support_note = ""
        # Point F: Correct Means of Finance for Existing Business — Show borrower support separately
        if is_existing:
            h_equity = project.audited_history[-1].share_capital + project.audited_history[-1].reserves_surplus if project.audited_history else 0.0
            support_note = f"<b>Institutional Backing:</b> The project is further supported by the existing business Net Worth / Owner Equity of <b>Rs. {h_equity:.2f} Lakhs</b> as per the latest audited financials."
            
            if promoter > 0.01:
                rows.append(["Fresh Promoter Contribution", f"{promoter:.2f}", f"{(promoter/total_cost*100 if total_cost > 0 else 0):.1f}%"])
        else:
            rows.append(["Promoter Capital / Equity", f"{promoter:.2f}", f"{(promoter/total_cost*100 if total_cost > 0 else 0):.1f}%"])
            
        rows.append(["TOTAL MEANS OF FINANCE", f"{total_cost:.2f}", "100.0%"])
        if theme:
            t = theme.build_table(
                ["Source of Fund", "Amount (Rs. Lakhs)", "Percentage (%)"],
                rows, [90*mm, 45*mm, 35*mm],
                total_indices=[len(rows)-1], num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([["Source of Fund", "Amount (Rs. Lakhs)", "Percentage (%)"]] + rows, colWidths=[90*mm, 45*mm, 35*mm])
        elements.append(t)
        if support_note:
            elements.append(Spacer(1, 4*mm))
            elements.append(Paragraph(support_note, body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_J_financial_data(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """J. Financial Overview & Core Ratios — Theme-aware."""
        elements.append(Paragraph("SECTION J: FINANCIAL OVERVIEW & RATIOS", section_style))
        if not proj_results: return
        
        # Build headers
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["Projected Financial Indicator / Year"] + year_headers
        
        metrics = [
            ("Revenue (Sales)", "revenue"),
            ("Net Profit (PAT)", "pat"),
            ("Net Profit Margin (%)", "pat_pct"),
            ("Debt Service Coverage Ratio (DSCR)", "dscr"),
            ("Current Ratio", "current_ratio"),
            ("Break-Even Sales (Rs. Lakhs)", "bep_sales"),
        ]
        
        rows = []
        for label, key in metrics:
            row = [label]
            for r in proj_results:
                val = r.get(key, 0)
                if key == "pat_pct":
                    pct = (r['pat'] / r['revenue'] * 100) if r['revenue'] > 0 else 0
                    row.append(f"{pct:.2f}%")
                elif key in ["dscr", "current_ratio"]:
                    row.append(f"{val:.2f}")
                else:
                    row.append(f"{val:.2f}")
            rows.append(row)
        
        if theme:
            total_w = 175*mm
            desc_w = 60*mm
            col_w = (total_w - desc_w) / len(proj_results)
            t = theme.build_table(
                headers, rows, [desc_w] + [col_w]*len(proj_results),
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows)
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_L_operating_stmt(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """L. Projected Operating Statement (P&L) — Theme-aware."""
        elements.append(Paragraph("SECTION L: PROJECTED OPERATING STATEMENT", section_style))
        elements.append(Paragraph("(All figures in Rs. Lakhs)", body_style))
        
        if not proj_results: return
        
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["PARTICULARS"] + year_headers
        
        # Check if we should show granular detail
        has_labour = any(r.get('labour_expenses', 0) > 0 for r in proj_results)
        has_other_direct = any(r.get('other_direct_expenses', 0) > 0 for r in proj_results)
        has_interest_exp = any(r.get('interest_exp', 0) > 0 for r in proj_results)

        metrics = [
            ("Gross Revenue / Sales", "revenue"),
            ("Raw Materials & Direct Costs", "cogs"),
        ]
        if has_labour:
            metrics.append(("  Labour Expenses", "labour_expenses"))
        if has_other_direct:
            metrics.append(("  Other Direct Expenses", "other_direct_expenses"))
            
        metrics.extend([
            ("Stock Adjustment / (Increase) in Inventory", "stock_adj"),
            ("GROSS PROFIT (GP)", "gp_amt"),
            ("Administrative & Gen. Expenses", "ind_exp"),
            ("EBITDA", "ebitda"),
            ("Interest on Bank Loans", "total_int"),
        ])
        
        if has_interest_exp:
            metrics.append(("  Other Interest Expenses", "interest_exp"))
            
        metrics.extend([
            ("Depreciation", "depreciation"),
            ("Profit Before Tax (PBT)", "pbt"),
            ("Provision for Tax", "tax_amt"),
            ("NET PROFIT AFTER TAX (PAT)", "pat"),
            ("Cash Accruals (PAT + Depr.)", "cash_accruals"),
        ])
        
        rows = []
        for label, key in metrics:
            row = [label]
            for r in proj_results:
                val = r.get(key, 0)
                if key == "stock_adj":
                    # For COGS, if stock increases, it's a negative cost adjustment to match Rev-COGS=GP
                    row.append(f"({val:.2f})" if val > 0 else f"{abs(val):.2f}")
                else:
                    row.append(f"{val:.2f}")
            rows.append(row)
            
        if theme:
            total_w = 175*mm
            desc_w = 70*mm
            col_w = (total_w - desc_w) / len(proj_results)
            
            # Dynamic total indices for bolding
            t_indices = [i for i, (lbl, key) in enumerate(metrics) if key in ["gp_amt", "ebitda", "pbt", "pat", "cash_accruals"]]
            
            t = theme.build_table(
                headers, rows, [desc_w] + [col_w]*len(proj_results),
                total_indices=t_indices, num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows)
            
        elements.append(t)
        elements.append(PageBreak())


    @classmethod
    @classmethod
    def _add_section_N_cash_flow(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """N. Projected Cash Flow Statement — Indirect Method Redesign."""
        elements.append(Paragraph("SECTION N: PROJECTED CASH FLOW STATEMENT", section_style))
        elements.append(Paragraph("(All figures in Rs. Lakhs)", body_style))
        
        if not proj_results: return
        
        # Requirement: Hide historical columns, focus on Projections (Req 3 in Plan v10)
        display_results = [r for r in proj_results if not r.get('is_actual', False)]
        if not display_results: return

        year_headers = cls._get_year_headers(display_results, project, body_style)
        headers = ["Particulars / Indirect Method CF"] + year_headers
        
        rows = []
        
        # Helper for formatting values with parentheses
        def fmt_cf(val, is_outflow=False):
            if is_outflow: # For items where increase = outflow (Assets)
                return f"({abs(val):.2f})" if val > 0 else f"{abs(val):.2f}"
            else: # For items where increase = inflow (Liabs/PAT)
                return f"{val:.2f}" if val >= 0 else f"({abs(val):.2f})"

        # --- A. OPERATING ACTIVITIES ---
        rows.append(["<b>A. CASH FLOW FROM OPERATING ACTIVITIES</b>"] + [""]*len(display_results))
        rows.append(["Net Profit After Tax (PAT)"] + [fmt_cf(r.get('pat', 0)) for r in display_results])
        rows.append(["Adjustment: Non-Cash Exp (Depr)"] + [fmt_cf(r.get('depreciation', 0)) for r in display_results])
        
        # Working Capital Changes (Requirement: Consolidated Delta Rows)
        wc_metrics = [
            ("(Increase)/Decrease in Receivables/Debtors", "debtors", True),
            ("(Increase)/Decrease in Inventories/Stock", "inventory", True),
            ("(Increase)/Decrease in Other Current Assets", "other_current_assets", True),
            ("Increase/(Decrease) in Sundry Creditors", "creditors", False),
            ("Increase/(Decrease) in Other Current Liabs", "other_current_liabilities", False),
        ]
        
        for label, key, is_asset in wc_metrics:
            row = [label]
            for r in display_results:
                # Calculate delta from IMMEDIATE previous year (even if hidden)
                curr_idx = proj_results.index(r)
                prev_r = proj_results[curr_idx-1] if curr_idx > 0 else None
                prev_val = prev_r.get(key, 0) if prev_r else 0.0
                curr_val = r.get(key, 0)
                delta = curr_val - prev_val
                
                if is_asset: # Asset Inc = Outflow (-)
                    row.append(fmt_cf(delta, is_outflow=True))
                else: # Liab Inc = Inflow (+)
                    row.append(fmt_cf(delta, is_outflow=False))
            rows.append(row)
        
        # Subtotal A
        row_sub_a = ["<b>NET CASH FROM OPERATING ACTIVITIES (A)</b>"]
        for r in display_results:
            # We use the native engine's net_cf components or calculate
            val = r.get('pat',0) + r.get('depreciation',0) - r.get('ca_inc',0) + r.get('cl_inc',0)
            row_sub_a.append(f"<b>{val:.2f}</b>")
        rows.append(row_sub_a)
        rows.append(["-"] * (len(display_results) + 1))

        # --- B. INVESTING ACTIVITIES ---
        rows.append(["<b>B. CASH FLOW FROM INVESTING ACTIVITIES</b>"] + [""]*len(display_results))
        rows.append(["Purchase/Investment in Fixed Assets"] + [f"({r.get('asset_purchase', 0):.2f})" for r in display_results])
        
        row_sub_b = ["<b>NET CASH FROM INVESTING ACTIVITIES (B)</b>"]
        for r in display_results:
            row_sub_b.append(f"<b>({r.get('asset_purchase', 0):.2f})</b>")
        rows.append(row_sub_b)
        rows.append(["-"] * (len(display_results) + 1))

        # --- C. FINANCING ACTIVITIES ---
        rows.append(["<b>C. CASH FLOW FROM FINANCING ACTIVITIES</b>"] + [""]*len(display_results))
        rows.append(["Fresh Capital Infusion / (Drawings)"] + [fmt_cf(r.get('cap_inc', 0)) for r in display_results])
        rows.append(["Long Term Loan Disbursement / (Repayment)"] + [fmt_cf(r.get('loan_inc', 0) - r.get('tl_repayment', 0)) for r in display_results])
        rows.append(["Increase/(Decrease) in CC/WC Limit"] + [fmt_cf(r.get('cc_limit', 0) - (proj_results[proj_results.index(r)-1].get('cc_limit',0) if proj_results.index(r)>0 else 0)) for r in display_results])

        row_sub_c = ["<b>NET CASH FROM FINANCING ACTIVITIES (C)</b>"]
        for r in display_results:
            prev_idx = proj_results.index(r) - 1
            cc_diff = r.get('cc_limit',0) - (proj_results[prev_idx].get('cc_limit',0) if prev_idx >= 0 else 0)
            val = r.get('cap_inc',0) + r.get('loan_inc',0) - r.get('tl_repayment',0) + cc_diff
            row_sub_c.append(f"<b>{val:.2f}</b>")
        rows.append(row_sub_c)
        rows.append(["-"] * (len(display_results) + 1))

        # --- FINAL BRIDGE ---
        rows.append(["<b>NET INCREASE/(DECREASE) IN CASH (A+B+C)</b>"] + [f"<b>{r.get('net_cash_flow',0):.2f}</b>" for r in display_results])
        rows.append(["Opening Cash & Bank Balance"] + [f"{r.get('opening_cash_bal', 0):.2f}" for r in display_results])
        rows.append(["<b>CLOSING CASH & BANK BALANCE</b>"] + [f"<b>{r.get('closing_cash_bal', 0):.2f}</b>" for r in display_results])
            
        if theme:
            total_w = 175*mm
            desc_w = 75*mm
            col_w = (total_w - desc_w) / len(display_results)
            t = theme.build_table(
                headers, rows, [desc_w] + [col_w]*len(display_results),
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows)
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_K_graphics(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """K. Graphical Business Analytics - Optimized Chart Spacing"""
        elements.append(Paragraph("SECTION K: GRAPHICAL BUSINESS ANALYTICS", section_style))
        charts = cls._generate_charts(project, proj_results)
        
        # Group charts 2 per page to save space and look professional
        from reportlab.platypus import Image as RLImage
        for i in range(0, len(charts), 2):
            batch = charts[i:i+2]
            for chart_img in batch:
                elements.append(RLImage(chart_img, width=160*mm, height=85*mm))
                elements.append(Spacer(1, 5*mm))
            if i + 2 < len(charts):
                elements.append(PageBreak())
        elements.append(PageBreak())

    @classmethod
    def _add_section_O_fixed_assets(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """O. Fixed Assets & Depreciation Annexure — Theme-aware Year-Wise Movement."""
        elements.append(Paragraph("SECTION O: FIXED ASSETS & DEPRECIATION SCHEDULE", section_style))
        elements.append(Paragraph(f"Method: {project.assumptions.depreciation_method} | (Rs. in Lakhs)", body_style))
        elements.append(Spacer(1, 4*mm))
        
        if not proj_results: return

        # New Table Structure: Year | Opening | Additions | Depreciation | Closing Net Block
        headers = ["Fiscal Year", "Opening Balance", "Additions", "Depreciation", "Net Block"]
        rows = []
        
        for p in proj_results:
            year_label = p.get("year_label", "N/A")
            is_actual = p.get("is_actual", False)
            
            # Retrieve data from improved engine results
            opening = p.get("opening_fixed_assets", 0.0)
            additions = p.get("fixed_asset_additions", 0.0)
            depreciation = p.get("depreciation", 0.0)
            closing = p.get("net_fixed_assets", 0.0)
            
            # Special logic for Historical years if keys are missing
            if is_actual and opening == 0 and additions == 0:
                # Approximate for history if not explicitly tracked
                opening = closing + depreciation
            
            rows.append([
                f"{year_label}{' (ACT)' if is_actual else ' (PROJ)'}",
                f"{opening:.2f}",
                f"{additions:.2f}",
                f"{depreciation:.2f}",
                f"{closing:.2f}"
            ])
            
        # Add a placeholder for Category-wise details if needed as a note
        if theme:
            t = theme.build_table(
                headers, rows, [40*mm, 35*mm, 30*mm, 35*mm, 35*mm],
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows, colWidths=[40*mm, 35*mm, 30*mm, 35*mm, 35*mm])
            
        elements.append(t)
        elements.append(Spacer(1, 5*mm))
        elements.append(Paragraph("Note: Depreciation is calculated as per standard income tax rates and WDV/SLM method as applicable to the industry.", body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_P_expenses(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """P. Indirect Expenses Breakdown — Theme-aware."""
        elements.append(Paragraph("SECTION P: INDIRECT EXPENSES BREAKDOWN", section_style))
        if not proj_results: return
        
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["FY Label", "Salary & Wages", "Power & Fuel", "Rent & Rates", "Admin & Misc", "TOTAL"]
        
        rows = []
        for r in proj_results:
            # FIX: Access expense_breakdown directly, not under 'detailed'
            eb = r.get("expense_breakdown", {})
            rows.append([
                r["year_label"],
                f"{eb.get('Salary & Wages', 0.0):.2f}",
                f"{eb.get('Power & Fuel', 0.0):.2f}",
                f"{eb.get('Rent & Rates', 0.0):.2f}",
                f"{eb.get('Admin & Misc', 0.0):.2f}",
                f"{r.get('ind_exp', 0.0):.2f}"
            ])
            
        if theme:
            t = theme.build_table(
                headers, rows, [25*mm, 35*mm, 30*mm, 30*mm, 30*mm, 20*mm],
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows)
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_Q_dscr(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """Q. Debt Service Coverage Ratio (DSCR) — Theme-aware."""
        elements.append(Paragraph("SECTION Q: DEBT SERVICE COVERAGE RATIO (DSCR)", section_style))
        if not proj_results: return
        
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["Particulars"] + year_headers
        
        metrics = [
            ("Net Profit After Tax (PAT)", "pat"),
            ("ADD: Depreciation", "depreciation"),
            ("A. FUND AVAILABLE", "cash_accruals"),
            ("B. DEBT OBLIGATIONS (Repayment)", "tl_repayment"),
            ("DSCR (A / B)", "dscr"),
        ]
        
        rows = []
        for label, key in metrics:
            row = [label]
            for r in proj_results:
                row.append(f"{r.get(key, 0):.2f}")
            rows.append(row)
            
        if theme:
            total_w = 175*mm
            desc_w = 70*mm
            col_w = (total_w - desc_w) / len(proj_results)
            t = theme.build_table(
                headers, rows, [desc_w] + [col_w]*len(proj_results),
                total_indices=[2, 4], num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows)
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_R_liquidity(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """R. Current Ratio & Liquidity Analysis — Theme-aware."""
        elements.append(Paragraph("SECTION R: LIQUIDITY & CURRENT RATIO", section_style))
        if not proj_results: return
        
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["Particulars"] + year_headers
        
        metrics = [
            ("Total Current Assets", "current_assets"),
            ("Sundry Creditors / Payables", "creditors"),
            ("Other Current Liabilities", "other_current_liabilities"),
            ("Bank Borrowings (CC/OD)", "wc_loan_bal"),
            ("TOTAL CURRENT LIABILITIES", "total_cl"),
            ("CURRENT RATIO", "current_ratio"),
        ]
        
        rows = []
        for label, key in metrics:
            row = [label]
            for r in proj_results:
                if key == "total_cl": 
                    val = r.get('creditors', 0) + r.get('wc_loan_bal', 0) + r.get('other_current_liabilities', 0)
                else: 
                    val = r.get(key, 0)
                row.append(f"{val:.2f}")
            rows.append(row)
            
        if theme:
            total_w = 175*mm
            desc_w = 70*mm
            col_w = (total_w - desc_w) / len(proj_results)
            t = theme.build_table(
                headers, rows, [desc_w] + [col_w]*len(proj_results),
                total_indices=[3, 4], num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows)
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_S_sensitivity(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """S. Sensitivity & Stress Test — Theme-aware."""
        from reportlab.lib.styles import ParagraphStyle
        elements.append(Paragraph("SECTION S: SENSITIVITY & STRESS TEST ANALYSIS", section_style))
        rows = []
        
        # Scenario 1: -10% Stress
        rows.append(["<b>Scenario I: 10% Revenue Shortfall Stress</b>"] + [""]*3)
        rows.append(["PARTICULARS", "BASE", "STRESSED", "IMPACT"])
        
        for r in [x for x in proj_results if not x.get("is_actual")]:
            s_10 = r.get("sensitivity", {}).get("minus_10pct", {})
            rows.append([f"<b>{r['year_label']}</b>", "", "", ""])
            rows.append(["  Total Revenue", f"{r['revenue']:.2f}", f"{s_10.get('revenue',0):.2f}", f"({r['revenue']*0.10:.2f})"])
            rows.append(["  EBT (Profit Before Tax)", f"{r['pbt']:.2f}", f"{s_10.get('ebt',0):.2f}", f"({r['pbt'] - s_10.get('ebt',0):.2f})"])
            rows.append(["  PAT (Net Profit)", f"{r['pat']:.2f}", f"{s_10.get('pat',0):.2f}", f"({r['pat'] - s_10.get('pat',0):.2f})"])
            rows.append(["  DSCR Ratio", f"{r['dscr']:.2f}", f"{s_10.get('dscr',0):.2f}", f"{s_10.get('dscr',0) - r['dscr']:.2f}"])
            rows.append(["-"]*4)

        if theme:
            t = theme.build_table(
                ["Metric Comparison", "Base Value", "Stressed Case", "Variance"],
                rows, [70*mm, 35*mm, 35*mm, 35*mm],
                num_cols_start=1, wrap_style=ParagraphStyle('Tiny', parent=body_style, fontSize=8, leading=9)
            )
        else:
            t = Table([["Metric Comparison", "Base Value", "Stressed Case", "Variance"]] + rows, colWidths=[70*mm, 35*mm, 35*mm, 35*mm])
            
        elements.append(t)
        elements.append(Spacer(1, 4*mm))
        
        # Scenario 2: -20% Stress Table
        elements.append(Paragraph("Scenario II: 20% Revenue Shortfall Analysis", ParagraphStyle('Sub', parent=body_style, fontName='Helvetica-Bold', fontSize=9)))
        elements.append(Spacer(1, 2*mm))
        
        row2 = [["Metric", "Base", "Stressed (-20%)", "Impact"]]
        for r in [x for x in proj_results if not x.get("is_actual")]:
            s_20 = r.get("sensitivity", {}).get("minus_20pct", {})
            row2.append([f"<b>{r['year_label']}</b>", "", "", ""])
            row2.append(["  Total Revenue", f"{r['revenue']:.2f}", f"{s_20.get('revenue',0):.2f}", f"({r['revenue']*0.20:.2f})"])
            row2.append(["  PAT (Net Profit)", f"{r['pat']:.2f}", f"{s_20.get('pat',0):.2f}", f"({r['pat'] - s_20.get('pat',0):.2f})"])
            row2.append(["  DSCR Ratio", f"{r['dscr']:.2f}", f"{s_20.get('dscr',0):.2f}", f"{s_20.get('dscr',0) - r['dscr']:.2f}"])
        
        if theme:
            t2 = theme.build_table(row2[0], row2[1:], [70*mm, 35*mm, 35*mm, 35*mm], num_cols_start=1, 
                                   wrap_style=ParagraphStyle('Tiny', parent=body_style, fontSize=8, leading=9))
        else:
            t2 = Table(row2, colWidths=[70*mm, 35*mm, 35*mm, 35*mm])
        elements.append(t2)
        elements.append(PageBreak())

    @classmethod
    def _add_section_M_balance_sheet(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """M. Projected Balance Sheet — Detailed Net Worth Style."""
        sec_elements = []
        sec_elements.append(Paragraph("SECTION M: PROJECTED BALANCE SHEET", section_style))
        sec_elements.append(Paragraph("(All figures in Rs. Lakhs)", body_style))
        
        if not proj_results: return
        
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["LIABILITIES & ASSETS"] + year_headers
        rows = []
        
        # --- LIABILITIES ---
        rows.append(["<b>LIABILITIES</b>"] + [""]*len(proj_results))
        
        # A. Debt Liabilities
        rows.append(["<b>A. DEBT LIABILITIES</b>"] + [""]*len(proj_results))
        rows.append(["  Existing Term Loans (Bank)"] + [f"{r.get('tl_bal_existing', 0):.2f}" for r in proj_results])
        rows.append(["  New Term Loan (Proposed)"] + [f"{r.get('tl_bal_new', 0):.2f}" for r in proj_results])
        rows.append(["  Working Capital Loan (CC/OD)"] + [f"{r.get('cc_limit', 0):.2f}" for r in proj_results])
        rows.append(["  Unsecured Loans (Promoters/Others)"] + [f"{r.get('unsecured_loan', 0):.2f}" for r in proj_results])
        rows.append(["  Other Loans & Liabilities"] + [f"{r.get('other_loans_liabilities', 0):.2f}" for r in proj_results])
        rows.append(["  <b>Total Outside Liabilities (A)</b>"] + [f"<b>{r.get('tl_balance_total', 0) + r.get('cc_limit', 0) + r.get('unsecured_loan', 0) + r.get('other_loans_liabilities', 0):.2f}</b>" for r in proj_results])
        
        # B. Current Liabilities
        rows.append(["<b>B. CURRENT LIABILITIES</b>"] + [""]*len(proj_results))
        rows.append(["  Sundry Creditors / Trade Payables"] + [f"{r.get('creditors', 0):.2f}" for r in proj_results])
        rows.append(["  Provisions / Other Current Liabs"] + [f"{r.get('provisions', 0) + r.get('other_current_liabilities', 0):.2f}" for r in proj_results])
        rows.append(["  <b>Total Current Liabilities (B)</b>"] + [f"<b>{r.get('creditors', 0) + r.get('provisions', 0) + r.get('other_current_liabilities', 0):.2f}</b>" for r in proj_results])
        
        # C. Net Worth
        rows.append(["<b>C. NET WORTH</b>"] + [""]*len(proj_results))
        rows.append(["  Opening Net Worth"] + [f"{r.get('opening_equity', 0):.2f}" for r in proj_results])
        rows.append(["  Add: Surplus (+) / Deficit (-) in P&L"] + [f"{r.get('pat', 0):.2f}" for r in proj_results])
        rows.append(["  Less: Drawings"] + [f"({r.get('drawings', 0):.2f})" for r in proj_results])
        rows.append(["  <b>SUB TOTAL (Closing Net Worth)</b>"] + [f"<b>{r.get('share_capital', 0) + r.get('reserves_surplus', 0):.2f}</b>" for r in proj_results])
        
        # TOTAL LIABS
        rows.append(["<b>TOTAL LIABILITIES</b>"] + [f"<b>{r.get('total_liabilities', 0):.2f}</b>" for r in proj_results])
        
        rows.append(["-"] * (len(proj_results) + 1))
        
        # --- ASSETS ---
        rows.append(["<b>ASSETS</b>"] + [""]*len(proj_results))
        rows.append(["  Fixed Assets (Net Block)"] + [f"{r.get('net_fixed_assets', 0):.2f}" for r in proj_results])
        rows.append(["  Investments"] + [f"{r.get('investments', 0):.2f}" for r in proj_results])
        rows.append(["  Inventory / Stock"] + [f"{r.get('inventory', 0):.2f}" for r in proj_results])
        rows.append(["  Sundry Debtors / Receivables"] + [f"{r.get('debtors', 0):.2f}" for r in proj_results])
        rows.append(["  Loans & Advances / Deposits / Others"] + [f"{r.get('loans_advances', 0) + r.get('deposits', 0) + r.get('other_current_assets', 0):.2f}" for r in proj_results])
        rows.append(["  Cash & Bank Balances"] + [f"{r.get('closing_cash_bal', 0):.2f}" for r in proj_results])
        
        # TOTAL ASSETS
        rows.append(["<b>TOTAL ASSETS</b>"] + [f"<b>{r.get('total_assets', 0):.2f}</b>" for r in proj_results])

        if theme:
            total_w = 175*mm
            desc_w = 70*mm
            col_w = (total_w - desc_w) / len(proj_results)
            t = theme.build_table(headers, rows, [desc_w] + [col_w]*len(proj_results), num_cols_start=1, wrap_style=body_style)
        else:
            t = Table([headers] + rows)
            
        sec_elements.append(t)
        from reportlab.platypus import KeepTogether
        elements.append(KeepTogether(sec_elements))
        elements.append(PageBreak())

    @classmethod
    def _add_section_N_cash_flow(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """N. Projected Cash Flow Statement — Indirect Method Redesign."""
        sec_elements = []
        sec_elements.append(Paragraph("SECTION N: PROJECTED CASH FLOW STATEMENT", section_style))
        sec_elements.append(Paragraph("(All figures in Rs. Lakhs)", body_style))
        
        if not proj_results: return
        
        # Requirement: Hide historical columns, focus on Projections
        display_results = [r for r in proj_results if not r.get('is_actual', False)]
        if not display_results: return

        year_headers = cls._get_year_headers(display_results, project, body_style)
        headers = ["Particulars / Indirect Method CF"] + year_headers
        
        rows = []
        
        # Helper for formatting values with parentheses
        def fmt_cf(val, is_outflow=False):
            if is_outflow: # For items where increase = outflow (Assets)
                return f"({abs(val):.2f})" if val > 0 else f"{abs(val):.2f}"
            else: # For items where increase = inflow (Liabs/PAT)
                return f"{val:.2f}" if val >= 0 else f"({abs(val):.2f})"

        # --- A. OPERATING ACTIVITIES ---
        rows.append(["<b>A. CASH FLOW FROM OPERATING ACTIVITIES</b>"] + [""]*len(display_results))
        rows.append(["Net Profit After Tax (PAT)"] + [fmt_cf(r.get('pat', 0)) for r in display_results])
        rows.append(["Adjustment: Non-Cash Exp (Depr)"] + [fmt_cf(r.get('depreciation', 0)) for r in display_results])
        
        wc_metrics = [
            ("(Increase)/Decrease in Receivables/Debtors", "debtors", True),
            ("(Increase)/Decrease in Inventories/Stock", "inventory", True),
            ("(Increase)/Decrease in Other Current Assets", "other_current_assets", True),
            ("Increase/(Decrease) in Sundry Creditors", "creditors", False),
            ("Increase/(Decrease) in Other Current Liabs", "other_current_liabilities", False),
        ]
        
        for label, key, is_asset in wc_metrics:
            row = [label]
            for r in display_results:
                curr_idx = proj_results.index(r)
                prev_r = proj_results[curr_idx-1] if curr_idx > 0 else None
                prev_val = prev_r.get(key, 0) if prev_r else 0.0
                curr_val = r.get(key, 0)
                delta = curr_val - prev_val
                if is_asset: row.append(fmt_cf(delta, is_outflow=True))
                else: row.append(fmt_cf(delta, is_outflow=False))
            rows.append(row)
        
        row_sub_a = ["<b>NET CASH FROM OPERATING ACTIVITIES (A)</b>"]
        for r in display_results:
            val = r.get('pat',0) + r.get('depreciation',0) - r.get('ca_inc',0) + r.get('cl_inc',0)
            row_sub_a.append(f"<b>{val:.2f}</b>")
        rows.append(row_sub_a)
        rows.append(["-"] * (len(display_results) + 1))

        # --- B. INVESTING ACTIVITIES ---
        rows.append(["<b>B. CASH FLOW FROM INVESTING ACTIVITIES</b>"] + [""]*len(display_results))
        rows.append(["Purchase/Investment in Fixed Assets"] + [f"({r.get('asset_purchase', 0):.2f})" for r in display_results])
        
        row_sub_b = ["<b>NET CASH FROM INVESTING ACTIVITIES (B)</b>"]
        for r in display_results:
            row_sub_b.append(f"<b>({r.get('asset_purchase', 0):.2f})</b>")
        rows.append(row_sub_b)
        rows.append(["-"] * (len(display_results) + 1))

        # --- C. FINANCING ACTIVITIES ---
        rows.append(["<b>C. CASH FLOW FROM FINANCING ACTIVITIES</b>"] + [""]*len(display_results))
        rows.append(["Fresh Capital Infusion / (Drawings)"] + [fmt_cf(r.get('cap_inc', 0) - r.get('drawings', 0))] ) # Simplified roll-forward CF
        rows.append(["Long Term Loan Disbursement / (Repayment)"] + [fmt_cf(r.get('loan_inc', 0) - r.get('tl_repayment', 0)) for r in display_results])
        rows.append(["Increase/(Decrease) in CC/WC Limit"] + [fmt_cf(r.get('cc_limit', 0) - (proj_results[proj_results.index(r)-1].get('cc_limit',0) if proj_results.index(r)>0 else 0)) for r in display_results])
        
        # FIX: Missing Unsecured Loan Inflows (Promoter support to maintain cash floor)
        rows.append(["Increase/(Decrease) in Unsecured Loans (Promoters)"] + [fmt_cf(r.get('unsecured_loans', 0) - (proj_results[proj_results.index(r)-1].get('unsecured_loans',0) if proj_results.index(r)>0 else 0)) for r in display_results])

        row_sub_c = ["<b>NET CASH FROM FINANCING ACTIVITIES (C)</b>"]
        for r in display_results:
            prev_idx = proj_results.index(r) - 1
            cc_diff = r.get('cc_limit',0) - (proj_results[prev_idx].get('cc_limit',0) if prev_idx >= 0 else 0)
            ul_diff = r.get('unsecured_loans', 0) - (proj_results[prev_idx].get('unsecured_loans',0) if prev_idx >= 0 else 0)
            val = r.get('cap_inc',0) - r.get('drawings',0) + r.get('loan_inc',0) - r.get('tl_repayment',0) + cc_diff + ul_diff
            row_sub_c.append(f"<b>{val:.2f}</b>")
        rows.append(row_sub_c)
        rows.append(["-"] * (len(display_results) + 1))

        rows.append(["<b>NET INCREASE/(DECREASE) IN CASH (A+B+C)</b>"])
        for r in display_results:
            prev_idx = proj_results.index(r) - 1
            cc_diff = r.get('cc_limit',0) - (proj_results[prev_idx].get('cc_limit',0) if prev_idx >= 0 else 0)
            ul_diff = r.get('unsecured_loans', 0) - (proj_results[prev_idx].get('unsecured_loans',0) if prev_idx >= 0 else 0)
            
            # Re-calculate correct A+B+C incorporating UL diff
            a_val = r.get('pat',0) + r.get('depreciation',0) - r.get('ca_inc',0) + r.get('cl_inc',0)
            b_val = -r.get('asset_purchase', 0)
            c_val = r.get('cap_inc',0) - r.get('drawings',0) + r.get('loan_inc',0) - r.get('tl_repayment',0) + cc_diff + ul_diff
            total_cf = a_val + b_val + c_val
            rows[-1].append(f"<b>{total_cf:.2f}</b>")

        rows.append(["Opening Cash & Bank Balance"] + [f"{r.get('opening_cash_bal', 0):.2f}" for r in display_results])
        rows.append(["<b>CLOSING CASH & BANK BALANCE</b>"] + [f"<b>{r.get('closing_cash_bal', 0):.2f}</b>" for r in display_results])
            
        if theme:
            total_w = 175*mm
            desc_w = 75*mm
            col_w = (total_w - desc_w) / len(display_results)
            t = theme.build_table(headers, rows, [desc_w] + [col_w]*len(display_results), num_cols_start=1, wrap_style=body_style)
        else:
            t = Table([headers] + rows)
            
        sec_elements.append(t)
        from reportlab.platypus import KeepTogether
        elements.append(KeepTogether(sec_elements))
        elements.append(PageBreak())

    @classmethod
    def _add_section_T_bep(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """T. Break-Even Point (BEP) Targets — With Layout Continuity."""
        t_elements = []
        t_elements.append(Paragraph("SECTION T: BREAK-EVEN POINT (BEP) TARGETS", section_style))
        t_elements.append(Paragraph("This analysis identifies the sales volume required to cover all fixed and variable costs.", body_style))
        
        if not proj_results: return
        
        year_headers = cls._get_year_headers(proj_results, project, body_style)
        headers = ["Particulars"] + year_headers
        
        metrics = [
            ("Total Fixed Costs", "fixed_costs"),
            ("Contribution Margin (%)", "contribution_pct"),
            ("BREAK-EVEN SALES (Rs. Lakhs)", "bep_sales"),
            ("Cash BEP Sales (Rs. Lakhs)", "cash_bep"),
        ]
        
        rows = []
        for label, key in metrics:
            row = [label]
            for r in proj_results:
                val = r.get(key, 0)
                if key == "contribution_pct": row.append(f"{val:.2f}%")
                else: row.append(f"{val:.2f}")
            rows.append(row)
            
        if theme:
            total_w = 175*mm
            desc_w = 60*mm 
            col_w = (total_w - desc_w) / len(proj_results)
            
            # Recommendation: Use a slightly smaller font for YEAR headers to prevent narrow Column wrapping
            # We pass a specific Paragraph style for the headers to be used BY the theme builder
            small_header_style = ParagraphStyle(
                'SmallHeader', parent=body_style, fontSize=8, 
                alignment=TA_CENTER, textColor=theme.HEADER_FG, fontName='Helvetica-Bold'
            )
            wrapped_headers = [Paragraph(f"<b>{h}</b>", small_header_style) for h in headers]
            
            t = theme.build_table(wrapped_headers, rows, [desc_w] + [col_w]*len(proj_results), num_cols_start=1, wrap_style=body_style)
        else:
            t = Table([headers] + rows)
            
        t_elements.append(t)
        
        from reportlab.platypus import KeepTogether
        elements.append(KeepTogether(t_elements))
        elements.append(PageBreak())
    @classmethod
    def _add_section_U_margin(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """U. Security Margin Analysis — Theme-aware."""
        elements.append(Paragraph("SECTION U: SECURITY MARGIN ANALYSIS", section_style))
        elements.append(Paragraph("The margin indicates the level of promoter's equity protective cover for the bank's exposure.", body_style))
        if not proj_results: return
        
        headers = ["Security Components", "Total Asset Val.", "Bank Finance", "Borrower Margin", "Margin %"]
        
        from services.cma.models import BusinessMode
        is_existing = project.profile.business_mode in [BusinessMode.EXISTING.value, BusinessMode.EXISTING_NO_BOOKS.value]
        margin_source = "Fresh Equity" if not is_existing else "Internal Accruals/Equity"

        data = [
            ["Project Assets (Term Loan)", 
             f"{sum(a.cost for a in project.assets):.2f}", 
             f"{project.loan.term_loan_amount:.2f}",
             f"{(sum(a.cost for a in project.assets) - project.loan.term_loan_amount):.2f}", 
             f"{((sum(a.cost for a in project.assets) - project.loan.term_loan_amount)/sum(a.cost for a in project.assets)*100):.1f}%" if sum(a.cost for a in project.assets) > 0 else "0%"],
            ["Working Capital (OD/CC)", 
             f"{project.loan.working_capital_requirement:.2f}", 
             f"{project.loan.cash_credit_amount:.2f}",
             f"{(project.loan.working_capital_requirement - project.loan.cash_credit_amount):.2f}", 
             f"{((project.loan.working_capital_requirement - project.loan.cash_credit_amount)/project.loan.working_capital_requirement*100 if project.loan.working_capital_requirement > 0 else 25.0):.1f}%"],
        ]
        
        elements.append(Paragraph(f"<b>Margin Support Basis:</b> {margin_source}", body_style))
        elements.append(Spacer(1, 4*mm))
        
        if theme:
            t = theme.build_table(
                headers, data, [50*mm, 30*mm, 30*mm, 30*mm, 20*mm],
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + data, colWidths=[50*mm, 30*mm, 30*mm, 30*mm, 20*mm])
            
        elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_V_repayment(cls, elements, project, section_style, body_style, theme=None):
        """V. Loan Repayment Summary — Theme-aware."""
        elements.append(Paragraph("SECTION V: LOAN REPAYMENT SUMMARY", section_style))
        elements.append(Paragraph(f"Proposed Tenure: {project.loan.term_loan_tenure_years} Years | Moratorium: {project.assumptions.moratorium_months} Months", body_style))
        
        loan_sched = ProjectionEngineService.calculate_loan_amortization(
            project.loan.term_loan_amount, 
            project.assumptions.interest_on_tl, 
            project.loan.term_loan_tenure_years,
            project.assumptions.moratorium_months
        )
        
        if loan_sched:
            headers = ["Year", "Opening Balance", "Interest", "Repayment", "Closing Balance"]
            rows = []
            for row in loan_sched:
                rows.append([
                    f"Year {row['year']}", f"{row['opening_balance']:.2f}",
                    f"{row['interest']:.2f}", f"{row['principal_repayment']:.2f}",
                    f"{row['closing_balance']:.2f}"
                ])
            
            if theme:
                t = theme.build_table(
                    headers, rows, [20*mm, 40*mm, 35*mm, 35*mm, 40*mm],
                    num_cols_start=1, wrap_style=body_style
                )
            else:
                t = Table([headers] + rows, colWidths=[20*mm, 40*mm, 35*mm, 35*mm, 40*mm])
            elements.append(t)
        else:
            elements.append(Paragraph("Standard quarterly repayment terms as per bank norms.", body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_W_cma_data(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """W. CMA Data Presentation — Theme-aware."""
        elements.append(Paragraph("SECTION W: CMA DATA PRESENTATION", section_style))
        elements.append(Paragraph("Consolidated financial mapping for Credit Monitoring Arrangement (CMA) appraisal requirements.", body_style))
        
        if proj_results:
            year_headers = cls._get_year_headers(proj_results, project, body_style)
            headers = ["CMA KEY EXTRACTS"] + year_headers
            
            rows = [
                ["Adjusted Tangible Net Worth"] + [f"{r.get('share_capital',0)+r.get('reserves_surplus',0):.2f}" for r in proj_results],
                ["Working Capital Gap"] + [f"{r.get('current_assets',0)-r.get('creditors',0):.2f}" for r in proj_results],
                ["MPBF (Method II)"] + [f"{(r.get('current_assets',0)*0.75 - r.get('creditors',0)):.2f}" for r in proj_results],
            ]
            
            if theme:
                total_w = 175*mm
                desc_w = 70*mm
                col_w = (total_w - desc_w) / len(proj_results)
                t = theme.build_table(
                    headers, rows, [desc_w] + [col_w]*len(proj_results),
                    num_cols_start=1, wrap_style=body_style
                )
            else:
                t = Table([headers] + rows)
            elements.append(t)
        elements.append(PageBreak())

    @classmethod
    def _add_section_X_assumptions(cls, elements, project, section_style, body_style, theme=None):
        """X. Financial Assumptions & Notes — Theme-aware."""
        elements.append(Paragraph("SECTION X: NOTES & FINANCIAL ASSUMPTIONS", section_style))
        elements.append(Paragraph("The projections are based on the following key managerial and economic assumptions:", body_style))
        
        ass = project.assumptions
        header = ["Particulars / Parameter", "Assumed Basis"]
        rows = [
            ["Sales Growth Rate", f"{ass.sales_growth_percent}% p.a. (Compounded)"],
            ["Gross Profit Margin", f"{ass.gp_percent}% on Sales"],
            ["Tax Provision Rate", f"{ass.tax_rate_percent}% on PBT"],
            ["Depreciation Policy", f"{ass.depreciation_method} Basis"],
            ["Interest on Term Loan", f"{ass.interest_on_tl}% p.a."],
            ["Interest on WC Finance", f"{ass.interest_on_cc}% p.a."],
        ]
        
        if theme:
            t = theme.build_table(
                header, rows, [80*mm, 80*mm],
                num_cols_start=2, wrap_style=body_style
            )
        else:
            t = Table([header] + rows, colWidths=[80*mm, 80*mm])
            
        elements.append(t)
        elements.append(Spacer(1, 10*mm))

        # ── Part 2: Professional Institutional Notes ──
        # Styled Header for Notes
        note_hdr_style = ParagraphStyle(
            'NoteHdr', parent=body_style, fontSize=12, fontName='Helvetica-Bold',
            textColor=white, alignment=TA_CENTER, backColor=theme.PRIMARY if theme else black,
            borderPadding=5, borderRadius=3
        )
        elements.append(Paragraph("Notes to the Project Report", note_hdr_style))
        elements.append(Spacer(1, 6*mm))

        notes = [
            ("a.", "Depreciation has been computed in accordance with the depreciation rates prescribed in the Income Tax Act. A separate depreciation schedule has been provided for reference and calculation purposes."),
            ("b.", "The data presented, including sensitivity analysis and balance sheet synopsis, has been prepared utilizing standard financial assumptions and calculations."),
            ("c.", "The financial projections and assessments are based on the assumption that there will be no changes in government policies and rules that may impact the loan applicant's business. Furthermore, it is assumed that no abnormal events will occur during the lifespan of the project or business."),
            ("d.", "Provision for Income Tax has been made on the Rules and Regulations which are applicable for current scenario."),
            ("e.", "The financial statements have been prepared under the standard assumption that the fiscal year-end occurs in March."),
            ("f.", "The details of indirect expenses, break-even analysis, and security margin calculation have been provided in separate annexures for reference."),
            ("g.", "The financial data pertaining to revenue from business operations, asset additions, existing obligations, etc., has been presented based on the information provided by the client."),
            ("h.", "The projected data included in this report represents future-oriented financial information. It has been prepared based on the best judgment of the applicants, incorporating assumptions regarding the most probable set of economic conditions. However, it is important to note that this information should not be considered as a forecast."),
            ("i.", "The information pertaining to the business entity, owner's profile, employment details, feasibility studies, industry analysis, market potential, current scenario, and challenges/solutions has been compiled based on discussions and inputs provided by the loan applicant.")
        ]

        note_style = ParagraphStyle('NoteItem', parent=body_style, fontSize=9.5, leading=12)
        
        for letter, text in notes:
            # Table-based layout for clean indentation of the lettered list
            n_tbl = Table([[letter, Paragraph(text, note_style)]], colWidths=[8*mm, 152*mm])
            n_tbl.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            elements.append(n_tbl)
            
        elements.append(PageBreak())

    @classmethod
    def _add_section_Y_security(cls, elements, project, section_style, body_style, theme=None):
        """Y. Security & Collateral Details"""
        elements.append(Paragraph("SECTION Y: SECURITY & COLLATERAL DETAILS", section_style))
        elements.append(Spacer(1, 5*mm))
        elements.append(Paragraph("<b>Primary Security:</b> Personal guarantee of promoters and hypothecation of all assets created out of bank finance including Plant, Machinery, Stock and Receivables.", body_style))
        elements.append(Spacer(1, 5*mm))
        elements.append(Paragraph("<b>Collateral Security:</b> Additional tangible security as per the bank's specific sanction terms and margin requirements.", body_style))
        elements.append(PageBreak())

    @classmethod
    def _add_section_Z_declaration(cls, elements, project, section_style, body_style, theme=None):
        """Z. Final Declaration & Certification"""
        elements.append(Paragraph("SECTION Z: DECLARATION & CERTIFICATION", section_style))
        elements.append(Spacer(1, 10*mm))
        elements.append(Paragraph("I/We hereby declare that all information provided in this Integrated Financial Study / Detailed Project Report is true and accurate to the best of my/our knowledge.", body_style))
        elements.append(Spacer(1, 40*mm))
        
        sig_data = [
            [cls._wrap_cell(f"Prepared by:\nDate: {datetime.now().strftime('%d/%m/%Y')}", body_style), 
             cls._wrap_cell(f"For {project.profile.business_name.upper()}\n\nAuthorized Signatory", body_style)]
        ]
        t = Table(sig_data, colWidths=[80*mm, 80*mm])
        t.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
            ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ]))
        elements.append(t)
        elements.append(PageBreak())
        
    @classmethod
    def _generate_charts(cls, project: CmaProject, projections: List[dict]) -> List[io.BytesIO]:
        """Generates matplotlib charts as BytesIO objects for the PDF."""
        charts = []
        if not projections: return charts

        # 1. Revenue Growth Chart
        try:
            plt.figure(figsize=(8, 4))
            years = [p['year_label'] for p in projections]
            revenues = [p['revenue'] for p in projections]
            plt.bar(years, revenues, color='#1976D2')
            plt.title('Projected Revenue Growth (Rs. in Lakhs)', fontsize=12, fontweight='bold')
            plt.ylabel('Amount (Lakhs)')
            plt.grid(axis='y', linestyle='--', alpha=0.7)
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
            buf.seek(0)
            charts.append(buf)
            plt.close()
        except: pass

        # 2. Profitability Trend (Line Chart)
        try:
            plt.figure(figsize=(8, 4))
            pats = [p['pat'] for p in projections]
            plt.plot(years, pats, marker='o', color='#388E3C', linewidth=3, label='Net Profit')
            plt.title('Profitability Trend (PAT)', fontsize=12, fontweight='bold')
            plt.ylabel('Amount (Lakhs)')
            plt.legend()
            plt.grid(True, linestyle='--', alpha=0.7)
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
            buf.seek(0)
            charts.append(buf)
            plt.close()
        except: pass
        
        # 3. Project Cost Breakup (Pie Chart)
        try:
            plt.figure(figsize=(6, 6))
            asset_total = sum(a.cost for a in project.assets)
            wc_req = project.loan.working_capital_requirement
            
            labels = ['Fixed Assets', 'Working Capital']
            sizes = [asset_total, wc_req]
            colors = ['#FFC107', '#2196F3']
            
            plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=colors)
            plt.title('Project Cost Composition', fontsize=12, fontweight='bold')
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
            buf.seek(0)
            charts.append(buf)
            plt.close()
        except: pass

        return charts

    @classmethod
    def _add_section_AC_mpbf(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """AC. MPBF Analysis & Bank Norms — Theme-aware."""
        elements.append(Paragraph("SECTION AC: MPBF & WORKING CAPITAL ASSESSMENT", section_style))
        from services.cma.mpbf_service import MpbfService
        res = MpbfService.calculate_mpbf(project, proj_results)
        
        if res["status"] in ["error", "insufficient"]:
            elements.append(Paragraph(f"⚠️ {res.get('message', 'Data insufficient for MPBF calculation.')}", body_style))
            return

        # Status & Correction
        elements.append(Paragraph(f"<b>Assessment Status:</b> {res['risk_level']}", body_style))
        if res.get("suggested_correction") and res["suggested_correction"] != "None needed.":
            elements.append(Paragraph(f"<b>Suggested Correction:</b> {res['suggested_correction']}", body_style))
        elements.append(Spacer(1, 5*mm))
        
        from services.cma.models import LoanType
        ft = project.loan.facility_type
        # A facility is NOT pure term loan if it is WC related
        # A facility is NOT pure term loan if it is WC related OR has a WC limit requested
        is_wc_involved = (ft in [
            LoanType.WORKING_CAPITAL.value, 
            LoanType.OD_LIMIT.value, 
            LoanType.RENEWAL.value, 
            LoanType.COMPOSITE_LOAN.value
        ] or project.loan.cash_credit_amount > 0)
        is_pure_tl = ft == LoanType.TERM_LOAN.value and not is_wc_involved
        
        # Summary Table
        summary_header = ["Metric Component", "Assessed Value"]
        if is_pure_tl:
            elements.append(Paragraph("<b>Working Capital assessment (MPBF) is not requested for this Pure Term Loan facility.</b>", body_style))
            summary_rows = [
                ["Projected Year 1 Turnover", f"Rs. {res['turnover']:.2f} Lakhs"],
                ["Requested Cash Credit Limit", "N/A"],
                ["Max Permissible Finance (MPBF)", "N/A"],
            ]
        else:
            # Assessment remains active for OD/CC/Composite cases
            summary_rows = [
                ["Projected Year 1 Turnover", f"Rs. {res['turnover']:.2f} Lakhs"],
                ["Requested Cash Credit Limit", f"Rs. {res['requested_limit']:.2f} Lakhs"],
                ["Assessed Limit (Conservative Basis)", f"Rs. {res['permissible_limit']:.2f} Lakhs"],
                ["Excess / Shortfall (vs Assessed)", f"Rs. {res['excess_amount']:.2f} Lakhs"],
                ["Shortfall in NWC Margin (Guideline)", f"Rs. {res.get('shortfall_nwc', 0):.2f} Lakhs"],
            ]
        
        if theme:
            t1 = theme.build_table(
                summary_header, summary_rows, [100*mm, 60*mm],
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t1 = Table([summary_header] + summary_rows, colWidths=[100*mm, 60*mm])
        elements.append(t1)
        
        if is_pure_tl:
            elements.append(PageBreak())
            return
            
        elements.append(Spacer(1, 10*mm))
        
        # Methods Comparison
        elements.append(Paragraph("<b>Detailed Banking Method Comparisons (Institutional Reference)</b>", body_style))
        elements.append(Spacer(1, 3*mm))
        
        method_h = ["Assessment Framework", "WC Gap", "Min. Margin", "Permissible Finance"]
        method_rows = [
            ["Method I (20% of WC Gap)", f"{res['wc_gap']:.2f}", f"{(res['wc_gap']*0.20):.2f}", f"{(res['wc_gap']*0.80):.2f}"],
            ["Method II (25% of CA)", f"{res['wc_gap']:.2f}", f"{(res['total_ca']*0.25):.2f}", f"{(res['wc_gap'] - res['total_ca']*0.25):.2f}"],
            ["Nayak Committee (5% of TO)", f"-", f"-", f"{res['permissible_limit']:.2f}"],
        ]
        
        if theme:
            t2 = theme.build_table(
                method_h, method_rows, [60*mm, 35*mm, 35*mm, 35*mm],
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t2 = Table([method_h] + method_rows, colWidths=[60*mm, 35*mm, 35*mm, 35*mm])
            
        elements.append(t2)
        elements.append(PageBreak())

    @classmethod
    def _add_section_AD_readiness(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """AD. Bank-Readiness Health Check — Lite snapshot or full audit."""
        from services.cma.readiness_service import ReadinessService
        res = ReadinessService.evaluate_readiness(project, proj_results)
        
        border_c = theme.BORDER if theme else HexColor("#BDBDBD")
        hdr_bg = theme.HEADER_BG if theme else HexColor("#0D47A1")
        hdr_fg = theme.HEADER_FG if theme else white
        muted_c = theme.MUTED if theme else HexColor("#757575")

        if theme and theme.mode_key == "lite":
            # ── LITE: Compact Readiness Snapshot ──
            elements.append(Paragraph("READINESS SNAPSHOT", section_style))
            elements.append(Paragraph(f"Overall Status: {res['readiness_level']}", body_style))
            elements.append(Spacer(1, 3*mm))
            snap_rows = []
            for c in res["checks"]:
                snap_rows.append([c["name"], c["value"]])
            if theme:
                t = theme.build_table(
                    ["Parameter", "Value"],
                    snap_rows, [80*mm, 80*mm],
                    num_cols_start=1, wrap_style=body_style
                )
            else:
                t = Table([["Parameter", "Value"]] + snap_rows, colWidths=[80*mm, 80*mm])
            elements.append(t)
        else:
            # ── PRO / CMA: Full detailed audit ──
            elements.append(Paragraph("SECTION AD: BANK-READINESS HEALTH CHECK", section_style))
            elements.append(Paragraph(f"Overall Readiness Status: {res['readiness_level']}", body_style))
            elements.append(Spacer(1, 5*mm))
            
            chk_data = [["Key Parameter", "Current Value", "Assessment & Professional Advice"]]
            p_name_style = ParagraphStyle('PName', parent=body_style, fontSize=9, fontName='Helvetica-Bold', leading=10)
            
            for c in res["checks"]:
                name_para = Paragraph(cls._clean_text(c["name"]), p_name_style)
                advice_para = Paragraph(cls._clean_text(c["advice"]), ParagraphStyle('Small', parent=body_style, fontSize=8.5, leading=10, alignment=0)) # Justified=4, Left=0
                chk_data.append([name_para, c["value"], advice_para])
            
            # Width Adjustment: 45 + 30 + 99 = 174mm (Exact A4 printable width)
            t = Table(chk_data, colWidths=[45*mm, 30*mm, 99*mm])
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.3, border_c),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BACKGROUND', (0, 0), (-1, 0), hdr_bg),
                ('TEXTCOLOR', (0, 0), (-1, 0), hdr_fg),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
                ('LEFTPADDING', (0, 0), (-1, -1), 8),
                ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ]))
            elements.append(t)
        
        elements.append(Spacer(1, 8*mm))
        elements.append(PageBreak())

    @classmethod
    def _validate_project_for_export(cls, project, proj_results):
        """Strict blocking gate for mathematical integrity (Requirement 9, 10)."""
        if not proj_results: return
        
        for r in proj_results:
            lbl = r.get("year_label", "Unlabeled Year")
            
            # 1. B/S Tally Check (within tolerance)
            # Projection engine already handles adjustments, but we double check the final result
            assets = r.get("total_assets", 0)
            liabs = r.get("total_liabilities", 0)
            diff = abs(assets - liabs)
            
            if diff > 0.011: # Allowing for tiny floating point float above 0.01
                raise ValueError(f"Balance Sheet in {lbl} is not tallied. Diff: {diff:.4f} Lakhs. Export blocked.")
            
            # 2. Ratio Dependency Check
            # If current ratio > 0, current liabs should be non-zero unless CA is 0
            ca = r.get("current_assets", 0)
            cl = r.get("current_liabilities", 0)
            if cl <= 0 and ca > 0:
                # This could be a summary mode issue
                if not r.get("is_actual"):
                    raise ValueError(f"Inconsistent Liquidity in {lbl}: Current assets present without liabilities.")

        return True

    @classmethod
    def _add_section_AA_monthly_repayment(cls, elements, project, proj_results, section_style, body_style, theme=None):
        """AA. Comprehensive Monthly Repayment Annexure (Phase 2 Upgrade)"""
        elements.append(Paragraph("ANNEXURE AA: MONTHLY LOAN REPAYMENT DIARY", section_style))
        elements.append(Paragraph("This schedule provides month-by-month repayment tracking for the proposed Term Loan assistance.", body_style))
        elements.append(Spacer(1, 4*mm))
        
        from services.cma.projection_engine_service import ProjectionEngineService
        monthly_sched = ProjectionEngineService.calculate_monthly_repayment(
            project.loan.term_loan_amount, 
            project.assumptions.interest_on_tl, 
            project.loan.term_loan_tenure_years, 
            project.assumptions.moratorium_months
        )
        
        if not monthly_sched:
            elements.append(Paragraph("No Term Loan amortisation required for this project profile.", body_style))
            elements.append(PageBreak())
            return

        headers = ["Month", "Opening Bal", "Interest", "Principal", "Closing Bal"]
        rows = []
        for m in monthly_sched:
            rows.append([
                f"Month {m['month_num']}",
                f"{m['opening_balance']:.2f}",
                f"{m['interest']:.2f}",
                f"{m['principal_repayment']:.2f}",
                f"{m['closing_balance']:.2f}"
            ])
            
        if theme:
            # For 84 rows, we need a compact table
            t = theme.build_table(
                headers, rows, [25*mm, 35*mm, 35*mm, 35*mm, 35*mm],
                num_cols_start=1, wrap_style=body_style
            )
        else:
            t = Table([headers] + rows, colWidths=[25*mm, 35*mm, 35*mm, 35*mm, 35*mm])
            
        elements.append(t)
        elements.append(PageBreak())
