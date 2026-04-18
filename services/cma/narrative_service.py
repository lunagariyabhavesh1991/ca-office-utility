"""
Narrative Service for CMA / DPR Builder.
Generates automated narrative sections based on party details and templates.
"""

from typing import Dict, Any
from services.cma.models import CmaProject, BusinessCategory

class NarrativeService:
    """Service to generate professional text sections for CMA/DPR reports."""

    TEMPLATES = {
        "executive_summary": {
            "default": (
                "EXECUTIVE OVERVIEW: This report presents the project appraisal for {business_name}, a {entity_type} based in {address}. "
                "The core objective of the project is the operational expansion and optimization of technical infrastructure. "
                "The total project cost is appraised at Rs. {total_cost:.2f} Lakhs, funded through a combination of Rs. {promoter_margin:.2f} Lakhs "
                "as promoter contribution and a requested bank limit of Rs. {loan_amount:.2f} Lakhs. "
                "The project demonstrates sound financial viability with a robust DSCR and satisfactory working capital turnover."
            ),
            BusinessCategory.BAKERY.value: (
                "{business_name} specializes in high-margin artisanal food production in the {address} region. "
                "With a requested credit limit of Rs. {loan_amount:.2f} Lakhs, the unit aims to modernize its baking infrastructure "
                "and scale capacity to meet the growing hygiene-conscious consumer segment. Led by {promoters}, "
                "the venture maintains a strong operational margin and contributes significantly to local employment."
            ),
            BusinessCategory.MACHINERY.value: (
                "The engineering unit {business_name} is positioned as a key fabrication partner in the industrial sector. "
                "The requested financial assistance of Rs. {loan_amount:.2f} Lakhs is designated for high-precision machinery acquisition "
                "and working capital stabilization. The project leverages {promoters}'s technical expertise to ensure "
                "world-class quality standards and timely delivery for core infrastructure infrastructure clients."
            ),
        },
        "promoter_profile": (
            "The project is anchored by the visionary leadership of {promoters}. The management team brings a comprehensive "
            "understanding of local market dynamics in {address} and possesses technical proficiency critical for success. "
            "The promoters have committed substantial personal capital, reflecting their confidence in the project's long-term "
            "solvency and strategic growth potential."
        ),
        "business_overview": (
            "Operating in the {category} industry, {business_name} focuses on delivering value through quality, reliability, and "
            "competitive positioning. The business model is structured to maintain lean operations while scaling capacity to meet "
            "market demand. The registered office at {address} serves as a central hub for logistics and operational management."
        ),
        "employment_details": (
            "As an MSME-compliant enterprise, the project creates sustainable livelihoods for {employee_count} individuals. "
            "This encompasses a balanced mix of skilled technical staff, semi-skilled laborers, and administrative professionals. "
            "The enterprise fosters a productive working environment, ensuring high retention and consistent output quality."
        ),
        "project_rationale": (
            "The rationale for this project is rooted in the significant demand-supply gap currently observed in the {category} market. "
            "By implementing and expanding operations at {address}, {business_name} is positioned to capture a significant market share "
            "while maintaining a healthy margin profile. The project is designed for scalability and high resource optimization."
        ),
        "means_of_finance_narrative": (
            "The Means of Finance is structured with a prudent Debt-Equity ratio, ensuring that the enterprise maintains adequate "
            "serviceability of all bank obligations. The promoter's contribution reflects a strong skin-in-the-game, which "
            "underpins the financial stability and creditworthiness of the project."
        ),
        "projection_rationale": (
            "The financial projections for {business_name} are based on a projected annual growth rate of {growth_rate:.1f}% per annum. "
            "{projection_base_note}"
            "{stabilization_note}{wc_reset_note}"
        ),
        "scheme_background": {
            "default": "The project is submitted for institutional credit appraisal under general banking norms.",
            "MUDRA": "This project is proposed under the Pradhan Mantri Mudra Yojana (PMMY), specifically targeting {mudra_type} category. The scheme aims to provide credit facilities to non-corporate, non-farm small/micro enterprises to foster local entrepreneurship.",
            "PMEGP": "This project is submitted under the Prime Minister's Employment Generation Programme (PMEGP). The proposal accounts for the applicable margin money subsidy (15-35%) and focuses on sustainable employment generation in {address}.",
            "Standup India": "The proposal is developed under the Standup India Scheme, aimed at promoting entrepreneurship among SC/ST and women entrepreneurs. The project fulfills the scheme's requirement of a composite loan between Rs. 10 Lakhs and Rs. 1 Crore."
        }
    }

    @classmethod
    def generate_section(cls, section_key: str, project: CmaProject) -> str:
        """
        Generates a narrative section for the report with data-driven insights.
        """
        from services.cma.projection_engine_service import ProjectionEngineService
        profile = project.profile
        loan = project.loan
        
        # Calculate base outlay
        total_assets_cost = sum(a.cost for a in project.assets)
        total_project_cost = total_assets_cost + loan.working_capital_requirement
        total_loan = loan.term_loan_amount + loan.cash_credit_amount
        promoter_margin = max(0, total_project_cost - total_loan)
        
        # Get financial snapshot for analytical narrative
        ratios = ProjectionEngineService.get_summary_ratios(project)
        dscr = ratios.get("avg_dscr", 0.0)
        curr_ratio = ratios.get("current_ratio", 0.0)
        der = ratios.get("debt_equity_ratio", 0.0)
        
        # Determine sentiment-aware phrases
        dscr_note = "a robust debt-servicing capability" if dscr > 1.25 else "adequate repayment comfort" if dscr > 1.15 else "a tight cash-flow position requiring close monitoring"
        liquidity_note = "excellent liquidity" if curr_ratio > 1.33 else "satisfactory working capital cycle" if curr_ratio > 1.1 else "tight liquidity levels"
        solvency_note = "sound capital structure" if der < 2.0 else "standard leverage" if der < 3.5 else "high leverage which requires prudent operational management"

        # Growth & Base logic
        ass = project.assumptions
        growth_rate = ass.sales_growth_percent
        base_year = project.audited_history[-1].year_label if project.audited_history else "N/A"
        
        projection_base_note = (
            f"For existing operations, revenue is projected to grow from the latest audited base of {base_year}, "
            "ensuring the appraisal accounts for immediate operational momentum. "
        ) if base_year != "N/A" else ""
        from services.cma.models import BusinessMode
        stabilization_note = ""
        if project.profile.business_mode == BusinessMode.NEW.value:
            stabilization_note = "As this is a new project, Year 1 serves as the stabilization phase with targeted base revenue, followed by standardized growth. "
        
        # Working Capital Reset Note (Point 3)
        wc_reset_note = ""
        projections = ProjectionEngineService.generate_full_projections(project)
        if len(projections) > 1:
            # Check for sharp release of cash in Year 1 of projections
            hist_years = [p for p in projections if p.get("is_actual")]
            proj_years = [p for p in projections if not p.get("is_actual")]
            if hist_years and proj_years:
                last_h = hist_years[-1]
                first_p = proj_years[0]
                # Compare Days or absolute levels
                h_debtors = last_h.get("debtors", 0)
                p_debtors = first_p.get("debtors", 0)
                if h_debtors > p_debtors * 1.5: # 50% drop
                    wc_reset_note = (
                        "The working capital cycle in the first projected year reflects a strategic shift towards normative industry benchmarks, "
                        "resulting in an optimization of trapped liquidity and improved cash accruals. "
                    )
        
        # Narrative Cleanup: Prioritize exact activity over broad category (Point 1, 2, 7 & 8)
        raw_desc = profile.description or ""
        display_category = profile.manual_business_category or profile.business_category
        
        # Two-layer output logic:
        # If the user provided a detailed description, we blend it or prefer it.
        if len(raw_desc) > 10:
            if "Manufacturing" in display_category or "Other" in display_category:
                display_category = f"{display_category} ({raw_desc[:40]}...)"
            
        # Target specific industry themes
        # Target specific industry themes
        industry_theme = "industrial manufacturing"
        infrastructure_term = "operational infrastructure development"
        output_term = "consistent quality output"
        ecosystem_term = "local and regional supply chains"
        
        lower_desc = raw_desc.lower()
        if "wind turbine" in lower_desc:
            industry_theme = "renewable energy engineering and spare parts manufacturing"
            ecosystem_term = "renewable energy sector"
        elif "cnc" in lower_desc or "fabrication" in lower_desc:
            industry_theme = "precision engineering and metal fabrication"
            infrastructure_term = "precision infrastructure development"
            output_term = "high-precision output"
            ecosystem_term = "renewable and industrial ecosystems"
        elif any(x in lower_desc for x in ["plastic", "household", "toys", "mould"]):
            industry_theme = "plastic and polymer products manufacturing"
            infrastructure_term = "production capacity expansion"
            output_term = "durable consumer-grade products"
            ecosystem_term = "consumer goods supply chain"

        # Context for templates
        context = {
            "business_name": profile.business_name,
            "entity_type": profile.entity_type,
            "address": profile.address,
            "promoters": profile.promoters,
            "category": display_category,
            "industry_theme": industry_theme,
            "employee_count": profile.employee_count,
            "total_cost": total_project_cost,
            "loan_amount": total_loan,
            "tl_amt": loan.term_loan_amount,
            "wc_limit": loan.cash_credit_amount,
            "promoter_margin": promoter_margin,
            "dscr_note": dscr_note,
            "liquidity_note": liquidity_note,
            "solvency_note": solvency_note,
            "dscr": f"{dscr:.2f}",
            "der": f"{der:.2f}",
            "growth_rate": growth_rate,
            "base_year": base_year,
            "projection_base_note": projection_base_note,
            "stabilization_note": stabilization_note,
            "wc_reset_note": wc_reset_note,
            "infrastructure_term": infrastructure_term,
            "output_term": output_term,
            "ecosystem_term": ecosystem_term,
            "description": raw_desc if len(raw_desc) > 5 else "specialized industrial operations",
            "mudra_type": "Tarun" if loan.term_loan_amount > 5 else "Kishor" if loan.term_loan_amount > 0.5 else "Shishu"
        }

        if section_key == "executive_summary":
            if loan.term_loan_amount > 0 and loan.cash_credit_amount > 0:
                facility_desc = f"a Term Loan of Rs. {loan.term_loan_amount:.2f} Lakhs and a Working Capital Limit (CC/OD) of Rs. {loan.cash_credit_amount:.2f} Lakhs"
            elif loan.cash_credit_amount > 0:
                facility_desc = f"a Working Capital Limit (CC/OD) of Rs. {loan.cash_credit_amount:.2f} Lakhs"
            else:
                facility_desc = f"a Term Loan of Rs. {loan.term_loan_amount:.2f} Lakhs"
            
            context["facility_desc"] = facility_desc
            template = (
                "EXECUTIVE SUMMARY: This project report details the strategic appraisal for {business_name}, "
                "an enterprise specialized in the {description}. Operating within the {industry_theme} sector at {address}, "
                "the project centers on operational scaling and {infrastructure_term}. "
                "The total appraised outlay of Rs. {total_cost:.2f} Lakhs is structured via {facility_desc}, "
                "supported by a promoter equity/margin of Rs. {promoter_margin:.2f} Lakhs. "
                "The proposal is underpinned by {dscr_note} (Avg. DSCR: {dscr}) and a {solvency_note}."
            )
            return template.format(**context)

        if section_key == "project_rationale":
            template = (
                "PROJECT RATIONALE: The venture is strategically focused on the {industry_theme} segment. "
                "The rationale is driven by the consistent market demand for {description} and the critical need for "
                "reliable supply chain partners in the {ecosystem_term}. By leveraging "
                "technical expertise and the strategic location at {address}, {business_name} targets {output_term} "
                "and recurring demand, ensuring long-term operational sustainability."
            )
            return template.format(**context)

        template = cls.TEMPLATES.get(section_key, "")
        if isinstance(template, dict):
            template = template.get(profile.business_category, template.get("default", ""))

        if not template:
            return ""

        try:
            return template.format(**context)
        except Exception:
            return template
