"""
Data models for CMA / DPR Builder.
Structured data classes for handling party-wise loan reports.
"""

from dataclasses import dataclass, field
from typing import List, Optional
from enum import Enum
from datetime import datetime


class EntityType(Enum):
    PROPRIETOR = "Proprietor"
    PARTNERSHIP = "Partnership"
    COMPANY = "Company"
    LLP = "LLP"
    OTHER = "Other"

class BusinessCategory(Enum):
    BAKERY = "Bakery / Food Production"
    MACHINERY = "Machinery / Engineering"
    CNC = "CNC Workshop / Fabrication"
    HVAC = "HVAC / Ducting"
    MANUFACTURING = "General Manufacturing"
    TRADING = "Retail / Trading"
    PHARMA = "Pharmaceuticals"
    DAIRY = "Dairy / Agriculture"
    SERVICE = "Service Business"
    GENERIC = "Other Business"

class DepreciationMethod(Enum):
    WDV = "WDV (Written Down Value)"
    SLM = "SLM (Straight Line Method)"

class ProjectionScenario(Enum):
    CONSERVATIVE = "Conservative"
    REALISTIC = "Realistic"
    OPTIMISTIC = "Optimistic"


class BusinessMode(Enum):
    NEW = "New Business / Startup"
    EXISTING = "Existing Business (Audited)"
    EXISTING_NO_BOOKS = "Existing Business (No Ready Books)"

class LoanType(Enum):
    TERM_LOAN = "Term Loan"
    MACHINERY_PURCHASE = "Machinery Purchase Loan"
    PROJECT_LOAN = "Project Loan"
    COMPOSITE_LOAN = "Composite Loan"
    WORKING_CAPITAL = "Working Capital / CC"
    OD_LIMIT = "OD Limit"
    RENEWAL = "CC / OD Renewal"
    EXPANSION = "Business Expansion Loan"
    SCHEME_LOAN = "Scheme Loan (Mudra/PMEGP)"
    TERM_LOAN_PLUS_WC = "Term Loan + CC/OD"

class SchemeType(Enum):
    MUDRA = "Mudra Scheme"
    PMEGP = "PMEGP Scheme"
    CGTMSE = "CGTMSE Covered"
    MSME = "MSME General"
    STANDUP_INDIA = "Stand-Up India"
    GENERAL = "General Bank Loan"

class ReportMode(Enum):
    LITE = "Lite Project Report (Compact)"
    PRO = "Pro Project Report (Professional)"
    CMA = "CMA Detailed Analysis (Banker-style)"

class DataStatus(Enum):
    AUDITED = "Audited"
    PROVISIONAL = "Provisional"
    UNAUDITED = "Unaudited / Estimate"
    NON_AUDITED = "Non-Audited"


@dataclass
class AssetItem:
    """Represents a single asset being purchased as part of the project."""
    name: str = ""
    cost: float = 0.0
    group: str = "General Assets"
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "cost": self.cost,
            "group": self.group,
            "notes": self.notes,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "AssetItem":
        return cls(**data)


@dataclass
class PartyProfile:
    """Core details of the business and promoters."""
    business_name: str = ""
    pan: str = ""
    promoters: str = ""
    entity_type: str = EntityType.PROPRIETOR.value
    establishment_date: str = ""
    address: str = ""
    description: str = ""
    business_category: str = BusinessCategory.GENERIC.value
    manual_business_category: str = "" # New field for manual override
    employee_count: int = 0
    security_type: str = ""
    security_value: float = 0.0
    
    # New Premium Metadata
    business_mode: str = BusinessMode.NEW.value
    loan_type: str = LoanType.TERM_LOAN.value
    scheme_type: str = SchemeType.GENERAL.value
    report_mode: str = ReportMode.PRO.value
    user_overrode_report_mode: bool = False

    def to_dict(self) -> dict:
        return {
            "business_name": self.business_name,
            "pan": self.pan,
            "promoters": self.promoters,
            "entity_type": self.entity_type,
            "establishment_date": self.establishment_date,
            "address": self.address,
            "description": self.description,
            "business_category": self.business_category,
            "manual_business_category": self.manual_business_category,
            "employee_count": self.employee_count,
            "security_type": self.security_type,
            "security_value": self.security_value,
            "business_mode": self.business_mode,
            "loan_type": self.loan_type,
            "scheme_type": self.scheme_type,
            "report_mode": self.report_mode,
            "user_overrode_report_mode": self.user_overrode_report_mode,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "PartyProfile":
        return cls(**{k: data.get(k, cls.__dataclass_fields__[k].default) 
                      if k in ["employee_count", "security_value"] or k.endswith("_mode") or k == "loan_type" or k == "scheme_type" or k == "user_overrode_report_mode" or k == "manual_business_category"
                      else data.get(k, "") for k in cls.__dataclass_fields__})


@dataclass
class LoanProfile:
    """Details regarding the loan requirement."""
    purpose: str = ""
    term_loan_amount: float = 0.0
    term_loan_tenure_years: int = 5
    tl_interest_rate: float = 10.0
    cc_interest_rate: float = 11.0
    working_capital_requirement: float = 0.0
    cash_credit_amount: float = 0.0
    facility_type: str = "Term Loan"

    def to_dict(self) -> dict:
        return {
            "purpose": self.purpose,
            "term_loan_amount": self.term_loan_amount,
            "term_loan_tenure_years": self.term_loan_tenure_years,
            "tl_interest_rate": self.tl_interest_rate,
            "cc_interest_rate": self.cc_interest_rate,
            "working_capital_requirement": self.working_capital_requirement,
            "cash_credit_amount": self.cash_credit_amount,
            "facility_type": self.facility_type,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "LoanProfile":
        # Handle migration from old interest_rate field
        if "interest_rate" in data and "tl_interest_rate" not in data:
            data["tl_interest_rate"] = data.get("interest_rate")
            
        return cls(**{k: data.get(k, cls.__dataclass_fields__[k].default) 
                      for k in cls.__dataclass_fields__})


@dataclass
class ReportVersion:
    """Metadata for a generated report version."""
    version_id: str = ""
    mode: str = "Standard"
    output_pdf_path: str = ""
    generated_on: str = ""
    remarks: str = ""

    def to_dict(self) -> dict:
        return {
            "version_id": self.version_id,
            "mode": self.mode,
            "output_pdf_path": self.output_pdf_path,
            "generated_on": self.generated_on,
            "remarks": self.remarks,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "ReportVersion":
        return cls(**{k: data.get(k, "") for k in cls.__dataclass_fields__})


@dataclass
class FinancialAssumptions:
    projection_years: int = 5
    depreciation_method: str = DepreciationMethod.WDV.value
    selected_scenario: str = ProjectionScenario.REALISTIC.value
    
    # Growth and Margins
    sales_growth_percent: float = 10.0
    revenue_base_override: float = 0.0 # New field for Startup manual revenue target
    gp_percent: float = 20.0
    indirect_expense_percent: float = 5.0
    
    # Working Capital Cycle (Days)
    debtor_days: int = 45
    creditor_days: int = 30
    stock_days: int = 60
    
    # Tax and Loan
    tax_rate_percent: float = 25.0
    interest_on_cc: float = 11.0
    interest_on_tl: float = 10.5
    moratorium_months: int = 6
    
    # Depreciation Rates (Standard Income Tax / Companies Act)
    depr_plant_machinery: float = 15.0
    depr_building: float = 10.0
    depr_furniture: float = 10.0
    depr_others: float = 15.0

    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items()}

    @classmethod
    def from_dict(cls, data: dict) -> "FinancialAssumptions":
        return cls(**{k: data.get(k, cls.__dataclass_fields__[k].default) for k in cls.__dataclass_fields__})


@dataclass
class AuditedData:
    year_label: str = "" # e.g. "2023-24 (A)"
    data_type: str = DataStatus.AUDITED.value
    revenue: float = 0.0
    net_profit: float = 0.0
    depreciation: float = 0.0
    interest_paid: float = 0.0
    share_capital: float = 0.0
    reserves_surplus: float = 0.0
    term_loans: float = 0.0
    current_liabilities: float = 0.0
    net_block: float = 0.0
    current_assets: float = 0.0
    cash_bank: float = 0.0
    
    # New Breakup Fields for Detailed Mode
    inventory: float = 0.0
    debtors: float = 0.0
    creditors: float = 0.0
    other_current_assets: float = 0.0
    other_current_liabilities: float = 0.0
    
    # Requirement: Granular Liabilities & Assets
    unsecured_loan: float = 0.0
    bank_od: float = 0.0
    other_loans_liabilities: float = 0.0
    provisions: float = 0.0
    loans_advances: float = 0.0
    deposits: float = 0.0
    investments: float = 0.0
    
    # New Granular P&L Fields (Requirement 11)
    opening_stock: float = 0.0
    cogs: float = 0.0
    gross_profit: float = 0.0
    salary_wages: float = 0.0
    labour_expenses: float = 0.0
    power_fuel: float = 0.0
    rent_rates: float = 0.0
    admin_expenses: float = 0.0
    other_direct_expenses: float = 0.0
    interest_exp: float = 0.0
    tax_amt: float = 0.0
    
    is_detailed: bool = False
    
    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items()}

    @classmethod
    def from_dict(cls, data: dict) -> "AuditedData":
        # Auto-infer is_detailed if breakups exist
        if not data.get("is_detailed"):
            if data.get("inventory") or data.get("debtors") or data.get("creditors"):
                data["is_detailed"] = True
                
        return cls(**{k: data.get(k, cls.__dataclass_fields__[k].default) 
                      for k in cls.__dataclass_fields__})
@dataclass
class SimplifiedData:
    """Used for 'Existing Business without ready balance sheet' mode."""
    approx_turnover: float = 0.0
    gp_percent: float = 15.0
    np_percent: float = 5.0
    receivables_estimate: float = 0.0
    inventory_estimate: float = 0.0
    creditors_estimate: float = 0.0
    fixed_assets_estimate: float = 0.0
    borrowing_estimate: float = 0.0
    cash_bank_estimate: float = 0.0

    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items()}

    @classmethod
    def from_dict(cls, data: dict) -> "SimplifiedData":
        return cls(**{k: data.get(k, cls.__dataclass_fields__[k].default) for k in cls.__dataclass_fields__})



@dataclass
class BrandingDetails:
    """Optional CA Firm branding for the report."""
    firm_name: str = ""
    prepared_by: str = ""
    contact_line: str = ""
    disclaimer: str = ""
    logo_path: str = ""

    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items()}

    @classmethod
    def from_dict(cls, data: dict) -> "BrandingDetails":
        return cls(**{k: data.get(k, "") for k in cls.__dataclass_fields__})


@dataclass
class CmaProject:
    """Top-level container for a CMA / DPR project."""
    party_id: str = ""
    is_new_project: bool = True
    profile: PartyProfile = field(default_factory=PartyProfile)
    loan: LoanProfile = field(default_factory=LoanProfile)
    assets: List[AssetItem] = field(default_factory=list)
    assumptions: FinancialAssumptions = field(default_factory=FinancialAssumptions)
    audited_history: List[AuditedData] = field(default_factory=list)
    past_years_count: int = 2
    simplified_data: SimplifiedData = field(default_factory=SimplifiedData)
    branding: BrandingDetails = field(default_factory=BrandingDetails)
    history: List[ReportVersion] = field(default_factory=list)
    created_at: str = ""
    updated_at: str = ""

    def to_dict(self) -> dict:
        return {
            "party_id": self.party_id,
            "is_new_project": self.is_new_project,
            "profile": self.profile.to_dict(),
            "loan": self.loan.to_dict(),
            "assets": [a.to_dict() for a in self.assets],
            "assumptions": self.assumptions.to_dict(),
            "audited_history": [ad.to_dict() for ad in self.audited_history],
            "past_years_count": self.past_years_count,
            "simplified_data": self.simplified_data.to_dict(),
            "branding": self.branding.to_dict(),
            "history": [v.to_dict() for v in self.history],
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "CmaProject":
        proj = cls()
        proj.party_id = data.get("party_id", "")
        proj.is_new_project = data.get("is_new_project", True)
        proj.profile = PartyProfile.from_dict(data.get("profile", {}))
        proj.loan = LoanProfile.from_dict(data.get("loan", {}))
        proj.assets = [AssetItem.from_dict(a) for a in data.get("assets", [])]
        proj.assumptions = FinancialAssumptions.from_dict(data.get("assumptions", {}))
        proj.audited_history = [AuditedData.from_dict(ad) for ad in data.get("audited_history", [])]
        proj.past_years_count = data.get("past_years_count", len(proj.audited_history) or 2)
        proj.simplified_data = SimplifiedData.from_dict(data.get("simplified_data", {}))
        proj.branding = BrandingDetails.from_dict(data.get("branding", {}))
        proj.history = [ReportVersion.from_dict(v) for v in data.get("history", [])]
        proj.created_at = data.get("created_at", "")
        proj.updated_at = data.get("updated_at", "")
        
        # Backward Compatibility: Sync is_new_project with business_mode if mode is missing
        profile_data = data.get("profile", {})
        if "business_mode" not in profile_data:
            if proj.is_new_project:
                proj.profile.business_mode = BusinessMode.NEW.value
            else:
                proj.profile.business_mode = BusinessMode.EXISTING.value
        
        # Map old report mode strings to new ones
        old_mode = profile_data.get("report_mode", "")
        if old_mode == "Draft Mode":
            proj.profile.report_mode = ReportMode.LITE.value
        elif old_mode in ["Professional Mode", "Bank-Specific Mode"]:
            proj.profile.report_mode = ReportMode.PRO.value
                
        return proj
