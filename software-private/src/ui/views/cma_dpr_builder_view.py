"""
CMA / DPR Builder — Main UI View
Premium tabbed workflow for building loan project reports.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import os
import threading
import logging
import uuid
from datetime import datetime

from ui.theme import Theme
from ui.components import DragDropArea
from services.cma.models import (
    CmaProject, AssetItem, ProjectionScenario, LoanType,
    DepreciationMethod, DataStatus, BusinessMode, ReportMode, BusinessCategory, SchemeType, EntityType
)
from services.cma.party_service import PartyMasterService
from services.cma.image_mapping_service import ImageMappingService
from services.cma.report_generator_service import ReportGeneratorService
from services.cma.word_generator_service import WordGeneratorService

logger = logging.getLogger(__name__)

# ─── Shared styling constants ───
_CARD_KW = dict(fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS,
                border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR)
_HDR_FONT = lambda sz=14: ctk.CTkFont(family=Theme.FONT_FAMILY, size=sz, weight="bold")
_BODY_FONT = lambda sz=12, bold=False: ctk.CTkFont(family=Theme.FONT_FAMILY, size=sz, weight="bold" if bold else "normal")
_MUTED = Theme.TEXT_MUTED
_ACCENT = Theme.ACCENT_BLUE
_HOVER = Theme.ACCENT_HOVER
_GREEN = Theme.ACCENT_GREEN
_AMBER = Theme.ACCENT_AMBER
_BG1 = Theme.BG_PRIMARY
_BG2 = Theme.BG_SECONDARY


class CmaDprBuilderView(ctk.CTkFrame):
    """Full CMA / DPR Builder module UI."""

    TAB_NAMES = ["Dashboard", "1. Setup", "2. Past Financials", "3. Project & Finance", "4. Assumptions", "5. Projections", "6. MPBF Analysis", "7. Bank Readiness", "8. Generate"]

    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color=_BG1, **kwargs)
        self.app_window = app_window
        # Set a temporary title to verify version
        self.app_window.title("CA Office PDF Utility - Premium CMA / DPR Suite")
        
        self.project = CmaProject()
        self._current_tab = 0

        # Main layout
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_tab_area()
        self._build_bottom_bar()
        self._show_tab(0) # Show Dashboard first

    # ──────────────────────── HEADER ────────────────────────
    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(0, 10))
        hdr.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(hdr, text="📊", font=ctk.CTkFont(size=32)).grid(row=0, column=0, padx=(0, 12))
        ctk.CTkLabel(hdr, text="CMA / DPR Builder", font=_HDR_FONT(24), anchor="w").grid(row=0, column=1, sticky="w")
        
        self.party_info_lbl = ctk.CTkLabel(hdr, text="New Project", font=_BODY_FONT(13), text_color=_AMBER)
        self.party_info_lbl.grid(row=0, column=2, padx=10)

        # Tab navigation buttons
        tabs_fr = ctk.CTkFrame(hdr, fg_color="transparent")
        tabs_fr.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self._tab_buttons = []
        for i, name in enumerate(self.TAB_NAMES):
            b = ctk.CTkButton(tabs_fr, text=name, width=130, height=36, corner_radius=10,
                              fg_color=_BG2, hover_color=_HOVER, text_color=_MUTED,
                              font=_BODY_FONT(12, bold=True), command=lambda idx=i: self._show_tab(idx))
            b.pack(side="left", padx=3)
            self._tab_buttons.append(b)

    # ──────────────────────── TAB AREA ────────────────────────
    def _build_tab_area(self):
        self._tab_container = ctk.CTkFrame(self, fg_color="transparent")
        self._tab_container.grid(row=1, column=0, sticky="nsew", padx=10)
        self._tab_container.grid_rowconfigure(0, weight=1)
        self._tab_container.grid_columnconfigure(0, weight=1)

    def _show_tab(self, idx):
        self._current_tab = idx
        for i, b in enumerate(self._tab_buttons):
            if i == idx:
                b.configure(fg_color=_ACCENT, text_color="white")
            else:
                b.configure(fg_color=_BG2, text_color=_MUTED)
        
        for w in self._tab_container.winfo_children():
            w.destroy()

        tab_builders = [
            self._tab_dashboard, self._tab_setup, 
            self._tab_audited_data, self._tab_project_assets, self._tab_assumptions, 
            self._tab_projections, self._tab_mpbf_analysis, self._tab_readiness_check,
            self._tab_generate
        ]
        tab_builders[idx]()
        
        # Refresh dashboard if we just entered it
        if idx == 0:
            self._refresh_project_list()

    # ──────────────────────── BOTTOM BAR ────────────────────────
    def _build_bottom_bar(self):
        bar = ctk.CTkFrame(self, fg_color=_BG2, corner_radius=10, height=52)
        bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(10, 0))
        bar.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(bar, text="💾 Save Project", width=140, height=36, corner_radius=8,
                      fg_color=_ACCENT, hover_color=_HOVER, font=_BODY_FONT(12, bold=True), 
                      command=self._save_project).grid(row=0, column=0, padx=(12, 6), pady=8)
        
        ctk.CTkButton(bar, text="🆕 New Party", width=140, height=36, corner_radius=8,
                      fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                      hover_color=_BG1, font=_BODY_FONT(12), command=self._new_project
                      ).grid(row=0, column=1, padx=6, pady=8)

        self._status_lbl = ctk.CTkLabel(bar, text="Ready", font=_BODY_FONT(12), text_color=_MUTED, anchor="e")
        self._status_lbl.grid(row=0, column=2, sticky="e", padx=(10, 20))

    # ════════════════════════════════════════════════════════════
    # TAB 1 — SETUP (NEW)
    # ════════════════════════════════════════════════════════════
    def _tab_setup(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)

        # Main Setup Card
        card = ctk.CTkFrame(container, **_CARD_KW)
        card.place(relx=0.5, rely=0.48, anchor="center", relwidth=0.75)
        
        ctk.CTkLabel(card, text="⚙️ Premium Project Configuration", font=_HDR_FONT(22)).pack(pady=(25, 5))
        ctk.CTkLabel(card, text="Guided workflow to optimize your project for bank sanction.", 
                      font=_BODY_FONT(13), text_color=_MUTED).pack(pady=(0, 25))
        
        from services.cma.models import BusinessMode, LoanType, ReportMode, SchemeType
        
        # Grid for selections
        sel_fr = ctk.CTkFrame(card, fg_color="transparent")
        sel_fr.pack(fill="x", padx=40)
        sel_fr.grid_columnconfigure(1, weight=1)

        def _add_sel(row, label, var, values, cmd=None):
            ctk.CTkLabel(sel_fr, text=label, font=_BODY_FONT(13, bold=True)).grid(row=row, column=0, sticky="w", pady=10)
            menu = ctk.CTkOptionMenu(sel_fr, variable=var, values=values,
                                     height=40, corner_radius=8, fg_color=_BG1, button_color=_ACCENT,
                                     command=cmd if cmd else lambda _: self._on_setup_change())
            menu.grid(row=row, column=1, sticky="ew", padx=(20, 0), pady=10)
            return menu

        self.mode_var = tk.StringVar(value=self.project.profile.business_mode)
        _add_sel(0, "Business Mode", self.mode_var, [m.value for m in BusinessMode])

        self.loan_type_var = tk.StringVar(value=self.project.profile.loan_type)
        _add_sel(1, "Loan Type", self.loan_type_var, [l.value for l in LoanType])

        self.scheme_var = tk.StringVar(value=self.project.profile.scheme_type)
        _add_sel(2, "Bank Scheme", self.scheme_var, [s.value for s in SchemeType])

        self.report_mode_var = tk.StringVar(value=self.project.profile.report_mode)
        _add_sel(3, "Report Mode", self.report_mode_var, [r.value for r in ReportMode], cmd=self._on_report_mode_manual_change)

        # Recommendation Panel
        self.reco_fr = ctk.CTkFrame(card, fg_color=_BG1, corner_radius=10, border_width=1, border_color=_ACCENT)
        self.reco_fr.pack(fill="x", padx=40, pady=20)
        
        self.reco_lbl = ctk.CTkLabel(self.reco_fr, text="", font=_BODY_FONT(12), text_color=_ACCENT, wraplength=500)
        self.reco_lbl.pack(pady=(12, 4), padx=15)
        
        self.reco_btn = ctk.CTkButton(self.reco_fr, text="↺ Use Recommended Mode", font=_BODY_FONT(11), 
                                       fg_color="transparent", text_color=_ACCENT, hover_color=_BG2, 
                                       width=150, height=24, command=self._use_recommended_mode)
        self.reco_btn.pack(pady=(0, 10))
        
        self._on_setup_change() # Trigger initial reco

        ctk.CTkButton(card, text="Build Project →", font=_HDR_FONT(18), height=52, corner_radius=12, 
                      fg_color=_GREEN, hover_color="#059669", 
                      command=lambda: self._show_tab(2)).pack(fill="x", padx=40, pady=(0, 30))

    def _on_setup_change(self):
        """Smarter logic to recommend report mode based on business profile."""
        from services.cma.models import ReportMode, BusinessMode, LoanType, SchemeType
        mode = self.mode_var.get()
        lt = self.loan_type_var.get()
        scheme = self.scheme_var.get()
        
        reco = ReportMode.PRO.value
        msg = "Recommended: Pro Mode (Flagship report for detailed project appraisal)."
        
        # Logic 1: CMA for existing business or WC renewals
        if mode == BusinessMode.EXISTING.value or any(x in lt for x in ["Renewal", "CC", "OD", "Working Capital"]):
            reco = ReportMode.CMA.value
            msg = "Recommended: CMA Mode (Banker-style analytical assessment for existing units)."
        
        # Logic 2: Lite for small schemes or new mudra/pmegp
        elif any(x in scheme for x in ["Mudra", "PMEGP"]) or any(x in lt for x in ["Small", "Composite"]):
            reco = ReportMode.LITE.value
            msg = "Recommended: Lite Mode (Compact 8-10 page report for simplified bank schemes)."
        
        # Logic 3: New Machinery / Term Loans favor Pro
        elif "Machinery" in lt or "Term Loan" in lt:
            reco = ReportMode.PRO.value
            msg = "Recommended: Pro Mode (Premium flagship presentation for capital funding)."

        if not self.project.profile.user_overrode_report_mode:
            self.report_mode_var.set(reco)
            self.reco_btn.pack_forget() # Hide button if already using recommended
        else:
            self.reco_btn.pack(pady=(0, 10)) # Show button if overrode
            
        self.reco_lbl.configure(text=f"💡 {msg}")
        self._sync_all_data()

    def _on_report_mode_manual_change(self, choice):
        """Triggered when user manually clicks the Report Mode dropdown."""
        self.project.profile.user_overrode_report_mode = True
        self.reco_btn.pack(pady=(0, 10))
        self._sync_all_data()

    def _use_recommended_mode(self):
        """Resets the override and triggers recommendation again."""
        self.project.profile.user_overrode_report_mode = False
        self._on_setup_change() # This will re-set the dropdown to recommended
    # ════════════════════════════════════════════════════════════
    def _tab_dashboard(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(2, weight=1)

        # 1. Active Project Card (if project open)
        if self.project.party_id:
            active_fr = ctk.CTkFrame(container, **_CARD_KW)
            active_fr.grid(row=0, column=0, sticky="ew", pady=(0, 15))
            
            # Left: Info
            info_fr = ctk.CTkFrame(active_fr, fg_color="transparent")
            info_fr.pack(side="left", fill="both", expand=True, padx=20, pady=15)
            
            ctk.CTkLabel(info_fr, text=f"Active Project: {self.project.profile.business_name}", 
                         font=_HDR_FONT(18), anchor="w").pack(anchor="w")
            ctk.CTkLabel(info_fr, text=f"PAN: {self.project.profile.pan} | Type: {self.project.profile.entity_type}", 
                         font=_BODY_FONT(12), text_color=_MUTED, anchor="w").pack(anchor="w", pady=(5, 0))
            
            # Right: Actions
            act_fr = ctk.CTkFrame(active_fr, fg_color="transparent")
            act_fr.pack(side="right", padx=20)
            
            ctk.CTkButton(act_fr, text="✏️ Edit Party Details", width=160, height=38, corner_radius=8,
                          fg_color=_ACCENT, hover_color=_HOVER, font=_BODY_FONT(12, bold=True),
                          command=self._edit_party_details).pack(pady=5)
            
            ctk.CTkButton(act_fr, text="📊 Open Setup", width=160, height=38, corner_radius=8,
                          fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                          font=_BODY_FONT(12), command=lambda: self._show_tab(1)).pack(pady=5)
            

        # Dashboard Top Actions
        top_fr = ctk.CTkFrame(container, fg_color="transparent")
        top_fr.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        ctk.CTkLabel(top_fr, text="Recent Parties / Projects", font=_HDR_FONT(18)).pack(side="left")

        # Proactive "Add New" Button on Dashboard
        ctk.CTkButton(top_fr, text="+ Add New Party", width=140, height=34, corner_radius=8,
                      fg_color=_GREEN, hover_color="#059669", font=_BODY_FONT(12, bold=True),
                      command=self._new_project).pack(side="right", padx=10)
        
        
        # Project List Scroll
        self.list_scroll = ctk.CTkScrollableFrame(container, fg_color=_BG1, corner_radius=10,
                                                   border_width=1, border_color=Theme.BORDER_COLOR)
        self.list_scroll.grid(row=2, column=0, sticky="nsew")
        self._refresh_project_list()

    def _refresh_project_list(self):
        if not hasattr(self, 'list_scroll') or not self.list_scroll.winfo_exists():
            return
            
        for w in self.list_scroll.winfo_children():
            w.destroy()
        
        projects = PartyMasterService.list_projects()
        if not projects:
            ctk.CTkLabel(self.list_scroll, text="No saved projects found.", 
                         font=_BODY_FONT(14), text_color=_MUTED).pack(pady=100)
            return

        for p in projects:
            self._build_project_row(p)

    def _build_project_row(self, p):
        row = ctk.CTkFrame(self.list_scroll, **_CARD_KW, height=60)
        row.pack(fill="x", pady=4, padx=5)
        row.pack_propagate(False)
        
        info_fr = ctk.CTkFrame(row, fg_color="transparent")
        info_fr.pack(side="left", fill="both", expand=True, padx=15)
        
        ctk.CTkLabel(info_fr, text=p["business_name"], font=_HDR_FONT(14), anchor="w").pack(side="top", anchor="w", pady=(8, 0))
        ctk.CTkLabel(info_fr, text=f"PAN: {p['pan']} | {p['entity_type']} | {p['category']}", 
                      font=_BODY_FONT(11), text_color=_MUTED, anchor="w").pack(side="top", anchor="w")
        
        btn_fr = ctk.CTkFrame(row, fg_color="transparent")
        btn_fr.pack(side="right", padx=10)
        
        ctk.CTkButton(btn_fr, text="📂 Open", width=80, height=30, corner_radius=6,
                      fg_color=_ACCENT, hover_color=_HOVER, command=lambda fp=p["file_path"]: self._load_project_by_path(fp)).pack(side="left", padx=4)
        
        ctk.CTkButton(btn_fr, text="🗑️", width=35, height=30, corner_radius=6,
                      fg_color="transparent", hover_color="#e74c3c", text_color=_MUTED,
                      command=lambda pid=p["party_id"]: self._delete_project(pid)).pack(side="left", padx=4)

    def _edit_party_details(self):
        """Switches to the Party Details view (overlay within Dashboard)."""
        for w in self._tab_container.winfo_children():
            w.destroy()
        self._tab_party_details()

    # ════════════════════════════════════════════════════════════
    # PARTY DETAILS VIEW (Accessible from Dashboard)
    # ════════════════════════════════════════════════════════════
    def _tab_party_details(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        
        # Header with Back button
        hdr_fr = ctk.CTkFrame(container, fg_color="transparent")
        hdr_fr.pack(fill="x", pady=(0, 15))
        
        ctk.CTkButton(hdr_fr, text="← Back to Dashboard", width=120, height=32, corner_radius=6,
                      fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                      command=lambda: self._show_tab(0)).pack(side="left")
        
        ctk.CTkButton(hdr_fr, text="💾 Save Changes", width=120, height=32, corner_radius=6,
                      fg_color=_ACCENT, command=self._save_party_details).pack(side="right", padx=10)

        scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        scroll.pack(fill="both", expand=True)
        scroll.grid_columnconfigure(1, weight=1)

        fields = [
            ("business_name", "Business Name *"),
            ("pan", "PAN (Business)"),
            ("description", "Business Description (Activity) *"),
            ("promoters", "Promoter Names"),
            ("address", "Registered Address"),
            ("establishment_date", "Date of Establishment"),
            ("employee_count", "Expected Employees"),
            ("security_type", "Primary Security Type"),
            ("security_value", "Security Value (Lakhs)"),
        ]

        ctk.CTkLabel(scroll, text="Basic Business Details", font=_HDR_FONT(18)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(10, 15))

        self._party_entries = {}
        for i, (key, label) in enumerate(fields, start=1):
            ctk.CTkLabel(scroll, text=label, font=_BODY_FONT(12)).grid(row=i, column=0, sticky="e", padx=(20, 10), pady=8)
            ent = ctk.CTkEntry(scroll, height=36, corner_radius=8, fg_color=_BG1, border_color=Theme.BORDER_COLOR)
            ent.grid(row=i, column=1, sticky="ew", padx=(0, 20), pady=8)
            val = getattr(self.project.profile, key, "")
            ent.insert(0, str(val) if val is not None else "")
            self._party_entries[key] = ent

        # Business Category Dropdown (New)
        current_row = len(fields) + 1
        ctk.CTkLabel(scroll, text="Business Category *", font=_BODY_FONT(12)).grid(row=current_row, column=0, sticky="e", padx=(20, 10), pady=8)
        self.business_category_var = tk.StringVar(value=self.project.profile.business_category)
        from services.cma.models import BusinessCategory
        bc_menu = ctk.CTkOptionMenu(scroll, variable=self.business_category_var, 
                                     values=[b.value for b in BusinessCategory],
                                     height=36, corner_radius=8, fg_color=_BG1, button_color=_ACCENT)
        bc_menu.grid(row=current_row, column=1, sticky="ew", padx=(0, 20), pady=8)
        
        # New: Manual Business category Entry
        current_row += 1
        ctk.CTkLabel(scroll, text="Or write Category manually", font=_BODY_FONT(10), text_color=_MUTED).grid(row=current_row, column=0, sticky="e", padx=(20, 10), pady=0)
        self.manual_category_var = tk.StringVar(value=self.project.profile.manual_business_category)
        cat_entry = ctk.CTkEntry(scroll, textvariable=self.manual_category_var, placeholder_text="e.g. Plastic Household Items", height=32, corner_radius=8)
        cat_entry.grid(row=current_row, column=1, sticky="ew", padx=(0, 20), pady=2)

        # Entity Type Dropdown
        current_row += 2 # Adjusted for manual category
        ctk.CTkLabel(scroll, text="Constitution *", font=_BODY_FONT(12)).grid(row=current_row, column=0, sticky="e", padx=(20, 10), pady=8)
        self.entity_type_var = tk.StringVar(value=self.project.profile.entity_type)
        et_menu = ctk.CTkOptionMenu(scroll, variable=self.entity_type_var, 
                                     values=[e.value for e in EntityType],
                                     height=36, corner_radius=8, fg_color=_BG1, button_color=_ACCENT)
        et_menu.grid(row=current_row, column=1, sticky="ew", padx=(0, 20), pady=8)

    # ════════════════════════════════════════════════════════════
    # TAB 3 — AUDITED DATA (Historical)
    # ════════════════════════════════════════════════════════════
    def _tab_audited_data(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)

        from services.cma.models import BusinessMode

        # Handle different modes
        mode = self.project.profile.business_mode
        
        if mode == BusinessMode.NEW.value:
            ctk.CTkLabel(container, text="🚀 New Business / Startup Mode", font=_HDR_FONT(18)).pack(pady=(50, 10))
            ctk.CTkLabel(container, text="No historical audited data is required for a new project.\nProceed directly to 'Project & Finance' to enter your startup costs.", 
                         font=_BODY_FONT(14), text_color=_MUTED, justify="center").pack(pady=20)
            ctk.CTkButton(container, text="Go to Project Cost →", width=220, height=40, font=_BODY_FONT(12, bold=True),
                          command=lambda: self._show_tab(4)).pack(pady=20)
            return

        if mode == BusinessMode.EXISTING_NO_BOOKS.value:
            self._tab_audited_simplified(container)
            return

        # EXISTING (Audited) Mode
        scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 1. Configuration Bar
        config_fr = ctk.CTkFrame(scroll, fg_color=_BG2, corner_radius=8)
        config_fr.pack(fill="x", pady=(0, 20))
        
        ctk.CTkLabel(config_fr, text="Historical Data Configuration", font=_HDR_FONT(14)).pack(side="left", padx=15, pady=10)
        ctk.CTkLabel(config_fr, text="(All figures below in Rs. Lakhs)", font=_BODY_FONT(11), text_color=_ACCENT).pack(side="left", padx=5)
        
        ctk.CTkLabel(config_fr, text="No. of Past Years:", font=_BODY_FONT(12)).pack(side="left", padx=(20, 5))
        self.past_years_var = tk.StringVar(value=str(self.project.past_years_count))
        years_menu = ctk.CTkOptionMenu(config_fr, variable=self.past_years_var, values=["1", "2", "3"],
                                       width=80, height=30, corner_radius=6, fg_color=_BG1, button_color=_ACCENT,
                                       command=lambda _: self._update_past_years_count())
        years_menu.pack(side="left", padx=5)

        from services.cma.models import AuditedData, DataStatus
        
        # Ensure audited history matches the required count
        while len(self.project.audited_history) < self.project.past_years_count:
            year_num = -(len(self.project.audited_history) + 1)
            self.project.audited_history.insert(0, AuditedData(year_label=f"Year {year_num}"))
        while len(self.project.audited_history) > self.project.past_years_count:
            self.project.audited_history.pop(0)

        self._audited_entries = []
        self._audited_meta = [] 
        self._tally_lbls = [] # Requirement: Track reconcile labels for all years
        
        for i, ad in enumerate(self.project.audited_history):
            card = ctk.CTkFrame(scroll, **_CARD_KW)
            card.pack(fill="x", pady=5)
            
            hdr_fr = ctk.CTkFrame(card, fg_color="transparent")
            hdr_fr.pack(fill="x", padx=15, pady=10)
            
            ctk.CTkLabel(hdr_fr, text="Year Label:", font=_BODY_FONT(11), text_color=_MUTED).pack(side="left")
            lbl_ent = ctk.CTkEntry(hdr_fr, width=120, height=28, corner_radius=4)
            lbl_ent.insert(0, ad.year_label)
            lbl_ent.pack(side="left", padx=(5, 20))
            
            ctk.CTkLabel(hdr_fr, text="Data Status:", font=_BODY_FONT(11), text_color=_MUTED).pack(side="left")
            status_var = tk.StringVar(value=ad.data_type)
            status_menu = ctk.CTkOptionMenu(hdr_fr, variable=status_var, values=[s.value for s in DataStatus],
                                            height=28, width=140, corner_radius=4, fg_color=_BG1, button_color=_ACCENT)
            status_menu.pack(side="left", padx=5)
            
            # Dual Import Buttons (Requirement 11)
            ctk.CTkButton(hdr_fr, text="📑 Import P&L", width=110, height=28, 
                          corner_radius=6, fg_color=_ACCENT, hover_color=Theme.ACTIVATION_HOVER,
                          font=_BODY_FONT(10, bold=True), 
                          command=lambda target=ad: self._import_financials_wizard(target, "Profit & Loss")).pack(side="right", padx=5)
            ctk.CTkButton(hdr_fr, text="📑 Import B/S", width=110, height=28, 
                          corner_radius=6, fg_color=_GREEN, hover_color=Theme.ACTIVATION_HOVER,
                          font=_BODY_FONT(10, bold=True), 
                          command=lambda target=ad: self._import_financials_wizard(target, "Balance Sheet")).pack(side="right", padx=5)
            
            self._audited_meta.append((lbl_ent, status_var))
            
            # Split View: Balance Sheet (Left) vs P&L (Right)
            grid_fr = ctk.CTkFrame(card, fg_color="transparent")
            grid_fr.pack(fill="x", padx=10, pady=5)
            grid_fr.columnconfigure((0, 1), weight=1)
            
            bs_col = ctk.CTkFrame(grid_fr, fg_color=_BG2, corner_radius=6)
            bs_col.grid(row=0, column=0, padx=5, sticky="nsew")
            pl_col = ctk.CTkFrame(grid_fr, fg_color=_BG2, corner_radius=6)
            pl_col.grid(row=0, column=1, padx=5, sticky="nsew")
            
            row_entries = {}
            
            # Balance Sheet Column
            ctk.CTkLabel(bs_col, text="BALANCE SHEET DETAIL", font=_BODY_FONT(11, bold=True), text_color=_ACCENT).pack(pady=5)
            bs_fields = [
                ("share_capital", "Capital / Equity"), ("reserves_surplus", "Reserves & Surplus"),
                ("term_loans", "Term Loans (Bank)"), ("unsecured_loan", "Unsecured Loans"),
                ("bank_od", "Bank OD A/c"),
                ("creditors", "Sundry Creditors"), ("other_loans_liabilities", "Other Loans & Liabs"),
                ("provisions", "Provisions"), ("other_current_liabilities", "Other Current Liabs"),
                ("net_block", "Fixed Assets (Net)"), ("investments", "Investments"),
                ("inventory", "Inventory / Stock"), ("debtors", "Debt / Receivables"), 
                ("loans_advances", "Loans & Advances"), ("deposits", "Deposits (Assets)"),
                ("other_current_assets", "Other Current Assets"), ("cash_bank", "Cash & Bank Balances")
            ]
            for key, label in bs_fields:
                f = ctk.CTkFrame(bs_col, fg_color="transparent")
                f.pack(fill="x", padx=10, pady=2)
                ctk.CTkLabel(f, text=label, font=_BODY_FONT(10), text_color=_MUTED).pack(side="left")
                ent = ctk.CTkEntry(f, height=28, width=120, corner_radius=4)
                try:
                    val = getattr(ad, key)
                except AttributeError:
                    val = 0.0
                ent.insert(0, str(val))
                ent.pack(side="right")
                row_entries[key] = ent
                # Requirement: Tally Check real-time tracking for B/S fields
                ent.bind("<KeyRelease>", lambda e, idx=i: self._update_tally_display(idx))
                
            # Profit & Loss Column
            ctk.CTkLabel(pl_col, text="PROFIT & LOSS DETAIL", font=_BODY_FONT(11, bold=True), text_color=_ACCENT).pack(pady=5)
            pl_fields = [
                ("revenue", "Revenue / Sales"), ("opening_stock", "Opening Stock"), 
                ("cogs", "Cost of Goods Sold (COGS)"), ("gross_profit", "Gross Profit (GP)"),
                ("salary_wages", "Salary & Wages"), ("labour_expenses", "Labour Exp."),
                ("power_fuel", "Power & Fuel"), ("rent_rates", "Rent & Rates"), 
                ("admin_expenses", "Admin Expenses"), ("other_direct_expenses", "Other Direct Exp."),
                ("interest_exp", "Interest Exp."), ("depreciation", "Depreciation"), 
                ("net_profit", "Net Profit (PAT)"), ("tax_amt", "Provision for Tax")
            ]
            for key, label in pl_fields:
                f = ctk.CTkFrame(pl_col, fg_color="transparent")
                f.pack(fill="x", padx=10, pady=2)
                ctk.CTkLabel(f, text=label, font=_BODY_FONT(10), text_color=_MUTED).pack(side="left")
                ent = ctk.CTkEntry(f, height=28, width=120, corner_radius=4)
                try:
                    val = getattr(ad, key)
                except AttributeError:
                    val = 0.0
                ent.insert(0, str(val))
                ent.pack(side="right")
                row_entries[key] = ent
                # Requirement: Tally Check real-time tracking
                ent.bind("<KeyRelease>", lambda e, idx=i: self._update_tally_display(idx))
            
            tally_fr = ctk.CTkFrame(card, fg_color=_BG1, height=40, corner_radius=6)
            tally_fr.pack(fill="x", padx=10, pady=(5, 10))
            
            t_set = {}
            
            ctk.CTkLabel(tally_fr, text="B/S RECONCILIATION:", font=_BODY_FONT(10, bold=True)).pack(side="left", padx=15)
            
            t_set["liabs"] = ctk.CTkLabel(tally_fr, text="Total Liabs: 0.00", font=_BODY_FONT(10))
            t_set["liabs"].pack(side="left", padx=10)
            
            t_set["assets"] = ctk.CTkLabel(tally_fr, text="Total Assets: 0.00", font=_BODY_FONT(10))
            t_set["assets"].pack(side="left", padx=10)
            
            t_set["diff"] = ctk.CTkLabel(tally_fr, text="Diff: 0.00", font=_BODY_FONT(11, bold=True))
            t_set["diff"].pack(side="left", padx=20)
            
            self._tally_lbls.append(t_set)
            
            self._audited_entries.append(row_entries)
            # Initial tally calc
            self.after(100, lambda idx=i: self._update_tally_display(idx))

    def _update_tally_display(self, idx):
        """Calculates and updates the tally bar for a specific year card."""
        if not hasattr(self, '_audited_entries') or idx >= len(self._audited_entries):
            return
            
        entries = self._audited_entries[idx]
        t_set = self._tally_lbls[idx]
        
        def gv(k):
            try: 
                val_str = entries[k].get().replace(",", "").strip()
                return float(val_str or 0.0)
            except: return 0.0
            
        # Tally Logic: Automated based on user feedback
        l_keys = ["share_capital", "reserves_surplus", "term_loans", "unsecured_loan", "bank_od", 
                  "creditors", "other_loans_liabilities", "provisions", "other_current_liabilities"]
        a_keys = ["net_block", "investments", "inventory", "debtors", 
                  "loans_advances", "deposits", "other_current_assets", "cash_bank"]
        
        total_l = sum(gv(k) for k in l_keys)
        total_a = sum(gv(k) for k in a_keys)
        diff = round(total_l - total_a, 2)
        
        t_set["liabs"].configure(text=f"Total Liabs: {total_l:.2f}")
        t_set["assets"].configure(text=f"Total Assets: {total_a:.2f}")
        
        if abs(diff) < 0.01:
            t_set["diff"].configure(text="✅ Tally Match", text_color=_GREEN)
        else:
            t_set["diff"].configure(text=f"⚠️ Difference: {diff:.2f}", text_color="#e74c3c")

    def _update_past_years_count(self):
        try:
            self.project.past_years_count = int(self.past_years_var.get())
            self._sync_all_data()
            self._tab_audited_data()
        except: pass

    def _import_financials_wizard(self, target_ad, import_type: str):
        """Step-by-step import wizard for specific year and document type."""
        file_path = filedialog.askopenfilename(
            title=f"Select {import_type} Statement for {target_ad.year_label}",
            filetypes=[("PDF Documents", "*.pdf")]
        )
        if not file_path: return
        
        logger.info(f"Analysing {target_ad.year_label} {import_type}...")
        
        from services.cma.extraction_engine_service import ExtractionEngineService
        result = ExtractionEngineService.extract_from_pdf(file_path)
        
        if result["status"] == "error":
            messagebox.showerror("Extraction Error", result["message"])
            return

        if not result["data"]:
            messagebox.showwarning("No Data Found", "No numeric data detected.")
            return

        self._show_mapping_review(result["data"], result["units_description"], target_ad, import_type)

    def _show_mapping_review(self, extracted_data: dict, detected_units: str, target_ad, import_type: str):
        """User review screen for imported values."""
        dialog = tk.Toplevel(self.app_window)
        dialog.title(f"Review & Map {import_type} Data")
        dialog.geometry("600x700")
        dialog.configure(bg=_BG1)
        dialog.transient(self.app_window)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text=f"Review extracted {import_type}", font=_HDR_FONT(18)).pack(pady=15)
        ctk.CTkLabel(dialog, text="Note: All figures converted to Rs. Lakhs", font=_BODY_FONT(12), text_color=_ACCENT).pack()
        
        # Filter fields based on import type (Requirement 11)
        bs_fields = [
            "share_capital", "reserves_surplus", "term_loans", "unsecured_loan", "bank_od",
            "creditors", "other_loans_liabilities", "provisions", "other_current_liabilities", 
            "net_block", "investments", 
            "inventory", "debtors", "loans_advances", "deposits", "other_current_assets", "cash_bank"
        ]
        pl_fields = [
            "revenue", "opening_stock", "cogs", "gross_profit", "salary_wages", "labour_expenses",
            "power_fuel", "rent_rates", "admin_expenses", "other_direct_expenses", 
            "interest_exp", "depreciation", "net_profit", "tax_amt"
        ]
        
        target_fields = bs_fields if import_type == "Balance Sheet" else pl_fields
        
        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=10)

        mapping_entries = {}
        for key in target_fields:
            val = extracted_data.get(key, 0.0)
            
            # Requirement: Show ALL fields for comprehensive review
            fr = ctk.CTkFrame(scroll, fg_color="transparent")
            fr.pack(fill="x", pady=2)
            display_name = key.replace("_", " ").title()
            if key == "tax_amt": display_name = "Provision for Tax"
            if key == "net_profit": display_name = "Net Profit (PAT)"
            if key == "other_loans_liabilities": display_name = "Other Loans & Liabilities"
            
            ctk.CTkLabel(fr, text=display_name, width=204, anchor="w", font=_BODY_FONT(12)).pack(side="left")
            ent = ctk.CTkEntry(fr, height=30, width=150)
            ent.insert(0, str(val))
            ent.pack(side="right", padx=10)
            mapping_entries[key] = ent

        if not mapping_entries:
            ctk.CTkLabel(scroll, text="No numeric adjustments proposed.", font=_BODY_FONT(12), text_color=_MUTED).pack(pady=20)

        def confirm():
            for key, ent in mapping_entries.items():
                try:
                    if hasattr(target_ad, key):
                        setattr(target_ad, key, float(ent.get()))
                except: pass
            
            # Auto-infer detailed mode
            if any(getattr(target_ad, k, 0) > 0 for k in ["inventory", "debtors", "creditors"]):
                target_ad.is_detailed = True
                
            # Refresh UI
            self._tab_audited_data()
            # Use messagebox for success since show_status is missing
            messagebox.showinfo("Import Success", f"Updated {len(mapping_entries)} fields in {target_ad.year_label}.\nRemaining year data was preserved.")
            dialog.destroy()

        ctk.CTkButton(dialog, text="Confirm & Import", command=confirm, fg_color=_ACCENT).pack(pady=20)

    def _tab_audited_simplified(self, container):
        """Simplified input for businesses without fixed books."""
        scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=10)
        scroll.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(scroll, text="Estimated Business Summary (Last Year)", font=_HDR_FONT(18)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(10, 20))
        
        fields = [
            ("approx_turnover", "Approx. Annual Turnover (Lakhs)"),
            ("gp_percent", "Estimated Gross Profit (%)"),
            ("np_percent", "Estimated Net Profit (%)"),
            ("fixed_assets_estimate", "Estimate Fixed Assets (Lakhs)"),
            ("receivables_estimate", "Estimate Debtors (O/s)"),
            ("inventory_estimate", "Estimate Stock on Hand"),
            ("creditors_estimate", "Estimate Creditors"),
            ("borrowing_estimate", "Existing Bank Borrowing"),
            ("cash_bank_estimate", "Approx. Cash/Bank Balance"),
        ]
        
        self._simplified_entries = {}
        for i, (key, label) in enumerate(fields, start=1):
            ctk.CTkLabel(scroll, text=label, font=_BODY_FONT(12)).grid(row=i, column=0, sticky="e", padx=10, pady=8)
            ent = ctk.CTkEntry(scroll, height=36, corner_radius=8, fg_color=_BG1)
            val = getattr(self.project.simplified_data, key)
            ent.insert(0, str(val))
            ent.grid(row=i, column=1, sticky="ew", padx=10, pady=8)
            self._simplified_entries[key] = ent
        
        ctk.CTkLabel(scroll, text="Note: This data will be used to construct a 'Draft' historical base for projections.", 
                      font=_BODY_FONT(11), text_color=_AMBER).grid(row=len(fields)+1, column=0, columnspan=2, pady=20)

    def _toggle_project_type(self):
        self.project.is_new_project = not self.project.is_new_project
        self._show_tab(self._current_tab)

    # ════════════════════════════════════════════════════════════
    # TAB 4 — PROJECT & ASSETS
    # ════════════════════════════════════════════════════════════
    def _tab_project_assets(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(1, weight=1)

        # Loan Details Section
        loan_fr = ctk.CTkFrame(container, **_CARD_KW)
        loan_fr.grid(row=0, column=0, sticky="ew", pady=(0, 10), padx=2)
        loan_fr.grid_columnconfigure((1, 3), weight=1)
        
        ctk.CTkLabel(loan_fr, text="Loan Requirement Details", font=_HDR_FONT(14)).grid(row=0, column=0, columnspan=4, sticky="w", padx=15, pady=10)
        
        # Purpose
        ctk.CTkLabel(loan_fr, text="Purpose:", font=_BODY_FONT(11)).grid(row=1, column=0, padx=(15, 5))
        self.loan_purpose_ent = ctk.CTkEntry(loan_fr, height=32, corner_radius=6, fg_color=_BG1)
        self.loan_purpose_ent.insert(0, self.project.loan.purpose)
        self.loan_purpose_ent.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(0, 15), pady=5)
        
        # Amount & Tenure
        fields = [
            ("term_loan_amount", "Term Loan (Lakhs)"),
            ("term_loan_tenure_years", "Tenure (Years)"),
            ("tl_interest_rate", "TL Int. Rate %"),
            ("cc_interest_rate", "CC/OD Int. Rate %"),
            ("working_capital_requirement", "WC Req. (Lakhs)"),
            ("cash_credit_amount", "Requested CC/OD Limit (Lakhs)"),
        ]
        self.loan_entries = {}
        for i, (key, label) in enumerate(fields):
            r, c = (2 if i < 2 else 3 if i < 4 else 4 if i < 6 else 5), (0 if i % 2 == 0 else 2)
            # Adjust grid for the new field
            if key == "cash_credit_amount":
                r, c = 4, 2
                ctk.CTkLabel(loan_fr, text=label, font=_BODY_FONT(11)).grid(row=r, column=c, padx=(15, 5), pady=5)
                ent = ctk.CTkEntry(loan_fr, height=32, corner_radius=6, fg_color=_BG1)
                ent.insert(0, str(getattr(self.project.loan, key)))
                ent.grid(row=r, column=c+1, sticky="ew", padx=(0, 5), pady=5)
                self.loan_entries[key] = ent
                
                # Add a helper button for OD cases
                ctk.CTkButton(loan_fr, text="Copy from Req.", width=80, height=24, corner_radius=4,
                              fg_color=_ACCENT, font=_BODY_FONT(9),
                              command=lambda e=ent: self._copy_wc_req(e)).grid(row=r, column=2, padx=5)
                continue

            ctk.CTkLabel(loan_fr, text=label, font=_BODY_FONT(11)).grid(row=r, column=c, padx=(15, 5), pady=5)
            ent = ctk.CTkEntry(loan_fr, height=32, corner_radius=6, fg_color=_BG1)
            ent.insert(0, str(getattr(self.project.loan, key)))
            ent.grid(row=r, column=c+1, sticky="ew", padx=(0, 15 if c == 2 else 5), pady=5)
            self.loan_entries[key] = ent

        # Facility Sub-Type (Logic Switch)
        ctk.CTkLabel(loan_fr, text="Proposed Facility Type:", font=_BODY_FONT(11)).grid(row=5, column=0, padx=(15, 5), pady=10)
        self.facility_type_var = tk.StringVar(value=getattr(self.project.loan, 'facility_type', LoanType.TERM_LOAN.value))
        facility_menu = ctk.CTkOptionMenu(loan_fr, variable=self.facility_type_var, 
                                          values=[l.value for l in LoanType],
                                          height=32, corner_radius=6, fg_color=_BG1, button_color=_ACCENT)
        facility_menu.grid(row=5, column=1, columnspan=3, sticky="ew", padx=(0, 15), pady=10)

        # Assets Section
        assets_hdr = ctk.CTkFrame(container, fg_color="transparent")
        assets_hdr.grid(row=1, column=0, sticky="ew", pady=(10, 5))
        ctk.CTkLabel(assets_hdr, text="Project Assets / Cost Breakdown", font=_HDR_FONT(14)).pack(side="left")
        ctk.CTkButton(assets_hdr, text="+ Add Asset", width=100, height=30, corner_radius=6,
                      fg_color=_GREEN, hover_color="#059669", font=_BODY_FONT(11, bold=True),
                      command=self._add_asset).pack(side="right")

        self.assets_scroll = ctk.CTkScrollableFrame(container, fg_color=_BG1, corner_radius=10, 
                                                   border_width=1, border_color=Theme.BORDER_COLOR)
        self.assets_scroll.grid(row=2, column=0, sticky="nsew", pady=5)
        self._refresh_assets_list()

    def _copy_wc_req(self, ent):
        """Helper to copy WC Requirement value to Requested CC Limit."""
        try:
            req_ent = self.loan_entries.get("working_capital_requirement")
            if req_ent:
                val = req_ent.get()
                ent.delete(0, "end")
                ent.insert(0, val)
                self._status_lbl.configure(text="Copied WC Requirement to CC Limit")
        except: pass

    def _add_asset(self):
        self.project.assets.append(AssetItem(name="New Asset", cost=0.0))
        self._refresh_assets_list()

    def _refresh_assets_list(self):
        for w in self.assets_scroll.winfo_children():
            w.destroy()
        
        self._asset_widgets = [] # Track widgets for manual sync
        
        if not self.project.assets:
            ctk.CTkLabel(self.assets_scroll, text="No assets added yet.", font=_BODY_FONT(12), text_color=_MUTED).pack(pady=20)
            return

        for i, asset in enumerate(self.project.assets):
            row = ctk.CTkFrame(self.assets_scroll, fg_color=_BG2, height=45, corner_radius=6)
            row.pack(fill="x", pady=2, padx=2)
            
            name_ent = ctk.CTkEntry(row, width=250, height=30, corner_radius=4, fg_color=_BG1)
            name_ent.insert(0, asset.name)
            name_ent.pack(side="left", padx=5, pady=5)
            
            cost_ent = ctk.CTkEntry(row, width=100, height=30, corner_radius=4, fg_color=_BG1)
            cost_ent.insert(0, str(asset.cost))
            cost_ent.pack(side="left", padx=5, pady=5)
            
            ctk.CTkLabel(row, text="Lakhs", font=_BODY_FONT(10), text_color=_MUTED).pack(side="left")
            
            self._asset_widgets.append((name_ent, cost_ent, asset))

            ctk.CTkButton(row, text="✕", width=30, height=30, corner_radius=4, 
                          fg_color="transparent", text_color=_MUTED, hover_color="#e74c3c",
                          command=lambda idx=i: self._remove_asset(idx)).pack(side="right", padx=5)

    def _update_asset_cost(self, asset, val):
        try:
            asset.cost = float(val)
        except: asset.cost = 0.0

    def _remove_asset(self, idx):
        self.project.assets.pop(idx)
        self._refresh_assets_list()

    # ════════════════════════════════════════════════════════════
    # TAB 4 — FINANCIAL ASSUMPTIONS
    # ════════════════════════════════════════════════════════════
    def _tab_assumptions(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew", padx=2)
        # Initialize Entry Storage
        self._ass_entries = {}
        curr_row = 0

        # 1. Report Settings
        ctk.CTkLabel(scroll, text="Report Settings & Depreciation", font=_HDR_FONT(16)).grid(row=curr_row, column=0, columnspan=2, sticky="w", pady=(10, 15))
        curr_row += 1
        
        # Years Dropdown
        ctk.CTkLabel(scroll, text="Projection Duration:", font=_BODY_FONT(12)).grid(row=curr_row, column=0, sticky="e", padx=10, pady=8)
        self.years_var = tk.StringVar(value=str(self.project.assumptions.projection_years))
        years_menu = ctk.CTkOptionMenu(scroll, variable=self.years_var, values=["3", "4", "5", "6", "7", "8", "9", "10"],
                                       height=36, corner_radius=8, fg_color=_BG1, button_color=_ACCENT)
        years_menu.grid(row=curr_row, column=1, sticky="w", padx=10, pady=8)
        curr_row += 1

        # Depr Method
        ctk.CTkLabel(scroll, text="Depreciation Method:", font=_BODY_FONT(12)).grid(row=curr_row, column=0, sticky="e", padx=10, pady=8)
        from services.cma.models import DepreciationMethod
        self.depr_meth_var = tk.StringVar(value=self.project.assumptions.depreciation_method)
        depr_menu = ctk.CTkOptionMenu(scroll, variable=self.depr_meth_var, values=[m.value for m in DepreciationMethod],
                                       height=36, corner_radius=8, fg_color=_BG1, button_color=_ACCENT)
        depr_menu.grid(row=curr_row, column=1, sticky="w", padx=10, pady=8)
        curr_row += 1

        # Projection Scenario
        ctk.CTkLabel(scroll, text="Projection Scenario:", font=_BODY_FONT(12)).grid(row=curr_row, column=0, sticky="e", padx=10, pady=8)
        self.scenario_var = tk.StringVar(value=self.project.assumptions.selected_scenario)
        self.scenario_var.trace_add("write", lambda *args: self._on_scenario_change())
        scenario_menu = ctk.CTkOptionMenu(scroll, variable=self.scenario_var, values=[s.value for s in ProjectionScenario],
                                       height=36, corner_radius=8, fg_color=_BG1, button_color=_ACCENT)
        scenario_menu.grid(row=curr_row, column=1, sticky="w", padx=10, pady=8)
        curr_row += 1

        # 2. Growth & Margins
        ctk.CTkLabel(scroll, text="Growth & Margins (%)", font=_HDR_FONT(16)).grid(row=curr_row, column=0, columnspan=2, sticky="w", pady=(20, 15))
        curr_row += 1
        
        from services.cma.models import BusinessMode
        # If it's a new business, add the manual revenue target option prominently
        if self.project.profile.business_mode == BusinessMode.NEW.value:
            ctk.CTkLabel(scroll, text="Base Year Revenue (Optional Target):", font=_BODY_FONT(12, bold=True), text_color=_ACCENT).grid(row=curr_row, column=0, sticky="e", padx=10, pady=8)
            ent = ctk.CTkEntry(scroll, height=36, corner_radius=8, fg_color=_BG1, border_width=1, border_color=_ACCENT)
            val = getattr(self.project.assumptions, "revenue_base_override", 0.0)
            ent.insert(0, str(val))
            ent.grid(row=curr_row, column=1, sticky="w", padx=10, pady=8)
            self._ass_entries["revenue_base_override"] = ent
            ctk.CTkLabel(scroll, text="(Set Year 1 Revenue Manual, or leave 0.0 for Auto)", font=_BODY_FONT(10), text_color=_MUTED).grid(row=curr_row, column=1, sticky="w", padx=(170, 0))
            curr_row += 1

        margin_fields = [
            ("sales_growth_percent", "Sales Growth Rate (%)"),
            ("gp_percent", "Gross Profit Margin (%)"),
            ("indirect_expense_percent", "Indirect Expenses (%)"),
            ("tax_rate_percent", "Income Tax Rate (%)"),
        ]
        
        for k, label in margin_fields:
            ctk.CTkLabel(scroll, text=label, font=_BODY_FONT(12)).grid(row=curr_row, column=0, sticky="e", padx=10, pady=8)
            ent = ctk.CTkEntry(scroll, height=36, corner_radius=8, fg_color=_BG1)
            val = getattr(self.project.assumptions, k, 0.0)
            ent.insert(0, str(val))
            ent.grid(row=curr_row, column=1, sticky="w", padx=10, pady=8)
            self._ass_entries[k] = ent
            curr_row += 1

        # 3. Working Capital Cycle
        ctk.CTkLabel(scroll, text="Working Capital Cycle (Days)", font=_HDR_FONT(16)).grid(row=curr_row, column=0, columnspan=2, sticky="w", pady=(20, 15))
        curr_row += 1
        
        wc_fields = [
            ("debtor_days", "Debtor Collection Days"),
            ("stock_days", "Stock Holding Days"),
            ("creditor_days", "Creditor Payment Days"),
        ]
        for k, label in wc_fields:
            ctk.CTkLabel(scroll, text=label, font=_BODY_FONT(12)).grid(row=curr_row, column=0, sticky="e", padx=10, pady=8)
            ent = ctk.CTkEntry(scroll, height=36, corner_radius=8, fg_color=_BG1)
            val = getattr(self.project.assumptions, k, 0)
            ent.insert(0, str(val))
            ent.grid(row=curr_row, column=1, sticky="w", padx=10, pady=8)
            self._ass_entries[k] = ent
            curr_row += 1

        # Real-time Stats Card (Floating side card)
        stats_fr = ctk.CTkFrame(container, **_CARD_KW, width=280)
        stats_fr.grid(row=0, column=1, sticky="ns", padx=(10, 0), pady=2)
        stats_fr.grid(row=0, column=1, sticky="ns", padx=(10, 0), pady=2)
        stats_fr.grid_propagate(False)
        
        ctk.CTkLabel(stats_fr, text="Project Ratios", font=_HDR_FONT(16)).pack(pady=15)
        
        self.ratio_lbls = {}
        ratio_defs = [
            ("total_loan", "Total Loan Requirement", "Lakhs"),
            ("promoter_contribution", "Promoter Contribution", "Lakhs"),
            ("margin_percent", "Margin Percentage", "%"),
            ("debt_equity_ratio", "Debt-Equity Ratio", ":1"),
        ]
        
        for key, name, unit in ratio_defs:
            ctk.CTkLabel(stats_fr, text=name, font=_BODY_FONT(11), text_color=_MUTED).pack(anchor="w", padx=20)
            lbl = ctk.CTkLabel(stats_fr, text="0.00", font=_HDR_FONT(20), text_color=_ACCENT)
            lbl.pack(anchor="w", padx=20, pady=(0, 10))
            self.ratio_lbls[key] = lbl

        ctk.CTkButton(stats_fr, text="🔄 Recalculate Ratios", height=36, corner_radius=8, 
                      command=self._refresh_ratios).pack(pady=20, padx=20, fill="x")
        self._refresh_ratios()

    def _refresh_ratios(self):
        self._sync_all_data()
        from services.cma.projection_engine_service import ProjectionEngineService
        stats = ProjectionEngineService.get_summary_ratios(self.project)
        
        is_valid = stats.get("is_valid", False)
        
        for key, lbl in self.ratio_lbls.items():
            if not is_valid and key != "total_loan":
                lbl.configure(text="Pending", text_color=_MUTED)
                continue

            val = stats.get(key, 0.0)
            lbl.configure(text_color=_ACCENT)
            if key == "debt_equity_ratio":
                lbl.configure(text=f"{val:.2f}")
            elif key == "margin_percent":
                lbl.configure(text=f"{val:.1f}%")
            else:
                lbl.configure(text=f"Rs. {val:.2f}")
        
        if not is_valid:
            self._status_lbl.configure(text="Add assets/project cost to calculate ratios.")
        else:
            self._status_lbl.configure(text="Ratios calculated successfully.")

    def _on_scenario_change(self):
        """Applies scenario presets to the assumption UI fields."""
        scenario = self.scenario_var.get()
        presets = {
            ProjectionScenario.CONSERVATIVE.value: {
                "sales_growth_percent": 5.0,
                "gp_percent": 15.0,
                "indirect_expense_percent": 7.0,
                "tax_rate_percent": 30.0
            },
            ProjectionScenario.REALISTIC.value: {
                "sales_growth_percent": 10.0,
                "gp_percent": 20.0,
                "indirect_expense_percent": 5.0,
                "tax_rate_percent": 25.0
            },
            ProjectionScenario.OPTIMISTIC.value: {
                "sales_growth_percent": 20.0,
                "gp_percent": 30.0,
                "indirect_expense_percent": 4.0,
                "tax_rate_percent": 22.0
            }
        }
        
        if scenario in presets:
            data = presets[scenario]
            if hasattr(self, '_ass_entries'):
                for k, v in data.items():
                    if k in self._ass_entries:
                        ent = self._ass_entries[k]
                        ent.delete(0, "end")
                        ent.insert(0, str(v))
            
            # Sync back to model immediately to reflect in preview
            self._sync_all_data()
            self._refresh_ratios()

    # ════════════════════════════════════════════════════════════
    # TAB 6 — FINANCIAL PROJECTIONS
    # ════════════════════════════════════════════════════════════
    def _tab_projections(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        try:
            self._sync_all_data()
            from services.cma.projection_engine_service import ProjectionEngineService
            results = ProjectionEngineService.generate_full_projections(self.project)
            
            if not results:
                raise ValueError("No projection data available. Ensure Past Financials and Assumptions are filled.")
        except Exception as e:
            err_fr = ctk.CTkFrame(container, fg_color="#3e1b1b", border_color="#e74c3c", border_width=1)
            err_fr.pack(fill="x", pady=50, padx=20)
            ctk.CTkLabel(err_fr, text="⚠️ CALCULATION ERROR", font=_BODY_FONT(14, bold=True), text_color="#ff5555").pack(pady=(15, 5))
            msg = f"An error occurred while generating projections: {str(e)}\n\nPlease ensure your Cost of Project and Past Financial data are realistic and complete."
            ctk.CTkLabel(err_fr, text=msg, font=_BODY_FONT(12), text_color="white", wraplength=800).pack(pady=(0, 20), padx=20)
            logger.exception("Projections tab error:")
            return

        # Sub-Tab Switcher (Internal)
        self.p_view_var = tk.StringVar(value="Summary")
        top_bar = ctk.CTkFrame(container, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ctk.CTkLabel(top_bar, text="Projected Financial Drill-down", font=_HDR_FONT(18)).pack(side="left")
        
        seg = ctk.CTkSegmentedButton(top_bar, values=["Summary", "P&L", "Balance Sheet", "Cash Flow"],
                                      command=lambda v: self._refresh_projection_view(results),
                                      variable=self.p_view_var, height=35)
        seg.pack(side="right", padx=10)

        self.p_scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        self.p_scroll.grid(row=1, column=0, sticky="nsew")

        self._refresh_projection_view(results)

    def _refresh_projection_view(self, results):
        try:
            for w in self.p_scroll.winfo_children(): w.destroy()
            view = self.p_view_var.get()

            # --- Negative Net Worth Warning (Bankability Alert) ---
            risk_year = self._check_net_worth_risk(results)
            if risk_year:
                warn_fr = ctk.CTkFrame(self.p_scroll, fg_color="#442222", border_color="#e74c3c", border_width=1)
                warn_fr.pack(fill="x", pady=(5, 15), padx=5)
            
                ctk.CTkLabel(warn_fr, text="🚨 CRITICAL BANKABILITY ALERT", 
                              font=_BODY_FONT(12, bold=True), text_color="#ff5555").pack(anchor="w", padx=15, pady=(10, 2))
            
                msg = f"Projected Net Worth becomes negative from {risk_year}. This is a major bankability concern and may adversely impact sanction of both term loan and working capital limits. Review project viability, repayment burden, and funding structure."
            
                ctk.CTkLabel(warn_fr, text=msg, font=_BODY_FONT(13), text_color="#FFD700", 
                              wraplength=1000, justify="left").pack(anchor="w", padx=15, pady=(0, 10))
        
            # Perfect Alignment Strategy: Fixed Column Widths
            LBL_W = 280
            VAL_W = 120
        
            # Table Headers
            thr = ctk.CTkFrame(self.p_scroll, fg_color=_BG2, corner_radius=8)
            thr.pack(fill="x", pady=(5, 10), padx=5)
        
            header_row = ctk.CTkFrame(thr, fg_color="transparent")
            header_row.pack(fill="x", padx=10, pady=8)
        
            ctk.CTkLabel(header_row, text="FINANCIAL PARAMETER", font=_BODY_FONT(11, bold=True), width=LBL_W, anchor="w").pack(side="left")
            for r in results:
                label = r["year_label"]
                badge = " (ACTUAL)" if r.get("is_actual") else " (PROJ)"
                ctk.CTkLabel(header_row, text=label + badge, font=_BODY_FONT(10, bold=True), width=VAL_W, text_color=_ACCENT if not r.get("is_actual") else _AMBER).pack(side="left")

            if view == "Summary":
                rows = [
                    ("SUMMARY DASHBOARD", "HEADER"),
                    ("💰 Total Sales / Revenue", "revenue"), 
                    ("⚡ EBITDA", "ebitda", True), 
                    ("💎 Net Profit (PAT)", "pat", True), 
                    ("-", "-"),
                    ("BANKABILITY INDICATORS", "HEADER"),
                    ("📊 Debt Service Coverage (DSCR)", "dscr"), 
                    ("⚖️ Current Ratio", "current_ratio"), 
                    ("-", "-"),
                    ("POSITIONAL METRICS", "HEADER"),
                    ("👤 Net Worth (Equity + Res)", "share_capital", True),
                    ("🏛️ Total Term Loan Bal", "tl_bal"), 
                    ("🏗️ Total Assets", "total_assets", True),
                ]
            elif view == "P&L":
                rows = [
                    ("OPERATING PERFORMANCE", "HEADER"),
                    ("💰 Gross Revenue", "revenue"), 
                    ("📦 COGS / Material Cost", "cogs"), 
                    ("📈 Gross Profit (GP)", "gp_amt"), 
                    ("-", "-"),
                    ("ESTABLISHMENT COST", "HEADER"),
                    ("🏢 Indirect Expenses", "ind_exp"), 
                    ("⚡ EBITDA", "ebitda", True),
                    ("📉 Depreciation", "depreciation"), 
                    ("🏦 Interest", "total_int"), 
                    ("-", "-"),
                    ("NET BOTTOM-LINE", "HEADER"),
                    ("📊 PBT", "pbt"), 
                    ("🧾 Tax Provision", "tax_amt"), 
                    ("💎 Net Profit (PAT)", "pat", True)
                ]
            elif view == "Balance Sheet":
                rows = [
                    ("LIABILITIES & CAPITAL", "HEADER"),
                    ("👤 Owner Capital", "share_capital"), 
                    ("💰 Reserves & Surplus", "reserves_surplus"), 
                    ("🏛️ Term Loan Bal", "tl_bal"), 
                    ("💳 WC / CC Loan", "wc_loan_bal"), 
                    ("🤝 Creditors", "creditors"), 
                    ("-", "-"),
                    ("FIXED & CURRENT ASSETS", "HEADER"),
                    ("🏗️ Net Fixed Assets", "net_fixed_assets"), 
                    ("📦 Current Assets", "current_assets"), 
                    ("🏦 Cash & Bank", "cash_bal"), 
                    ("🏁 TOTAL ASSETS", "total_assets", True)
                ]
            else: # Cash Flow
                rows = [
                    ("CASH SOURCES", "HEADER"),
                    ("📥 Sources: Net Profit", "pat"), 
                    ("🔄 Sources: Depreciation", "depreciation"), 
                    ("🏦 Sources: Loan/Cap", "loan_inc"), 
                    ("📊 TOTAL SOURCES", "total_sources", True),
                    ("-", "-"),
                    ("CASH USES", "HEADER"),
                    ("🏗️ Uses: Asset Purchase", "asset_purchase"), 
                    ("🏛️ Uses: Loan Repay", "tl_repayment"), 
                    ("📦 Uses: WC Increase", "ca_inc"), 
                    ("📊 TOTAL USES", "total_uses", True),
                    ("-", "-"),
                    ("CASH POSITION", "HEADER"),
                    ("💎 NET CASH FLOW", "net_cash_flow", True)
                ]

            # Explicit column setup
            LBL_W = 260
            VAL_W = 120
        
            # Guard definitions
            COLOR_MAP = {"PROJ": _ACCENT, "ACT": _AMBER, "BOLD_UP": _GREEN, "BOLD_DN": "#e74c3c"}

            for idx, r_item in enumerate(rows):
                try:
                    # 1. Unpack row data safely
                    if not r_item or not isinstance(r_item, (tuple, list)): continue
                
                    label = r_item[0]
                    key = r_item[1]
                    is_bold = r_item[2] if len(r_item) > 2 else False
                
                    # 2. Handle Separators
                    if label == "-":
                        ctk.CTkFrame(self.p_scroll, height=1, fg_color=_MUTED).pack(fill="x", pady=4, padx=20)
                        continue

                    # 3. Handle Headers
                    if key == "HEADER":
                        h_fr = ctk.CTkFrame(self.p_scroll, fg_color="transparent")
                        h_fr.pack(fill="x", pady=(8, 2))
                        ctk.CTkLabel(h_fr, text=label, font=_BODY_FONT(12, bold=True), text_color=_MUTED).pack(side="left", padx=10)
                        ctk.CTkFrame(h_fr, height=1, fg_color=_MUTED).pack(side="left", fill="x", expand=True, padx=10)
                        continue

                    # 4. Render Data Row
                    row_bg = _BG2 if idx % 2 == 0 else "transparent"
                    if is_bold: row_bg = _BG2 
                
                    fr = ctk.CTkFrame(self.p_scroll, fg_color=row_bg, corner_radius=8 if is_bold else 0)
                    fr.pack(fill="x", pady=0.5, padx=5)
                
                    # Perfect Alignment Strategy: Use a fixed 30px margin for labels
                    if is_bold:
                        # Sleek fixed-height accent bar (positioned within the margin)
                        bar = ctk.CTkFrame(fr, width=4, height=22, fg_color=_ACCENT, corner_radius=2)
                        bar.pack(side="left", padx=(8, 0), pady=6) 
                        bar.pack_propagate(False)
                        label_padx = (18, 10) # 8 (bar pad) + 4 (bar width) + 18 = 30
                    else:
                        label_padx = (30, 10) # Standard 30px margin

                    lbl_font = _BODY_FONT(13, bold=is_bold)
                    val_font = _BODY_FONT(12, bold=is_bold)
                
                    ctk.CTkLabel(fr, text=label, font=lbl_font, width=LBL_W, anchor="w").pack(side="left", padx=label_padx, pady=3)
                
                    for res in results:
                        if not isinstance(res, dict): continue
                    
                        # Custom aggregations
                        if key == "share_capital" and view == "Summary":
                            val = float(res.get("share_capital", 0.0) or 0.0) + float(res.get("reserves_surplus", 0.0) or 0.0)
                        else:
                            val = res.get(key, 0.0)
                    
                        if val is None: val = 0.0
                        try: 
                            num_val = float(val)
                        except: 
                            num_val = 0.0
                    
                        txt = f"{num_val:.2f}"
                        color = "white"
                        if is_bold:
                            color = COLOR_MAP["BOLD_UP"] if num_val >= 0 else COLOR_MAP["BOLD_DN"]
                        elif res.get("is_actual"):
                            color = COLOR_MAP["ACT"]
                        else:
                            color = "white"
                    
                        ctk.CTkLabel(fr, text=txt, font=val_font, width=VAL_W, text_color=color).pack(side="left")
                    
                except Exception as row_err:
                    # If a row fails, show a small error label instead of going blank
                    err_fr = ctk.CTkFrame(self.p_scroll, fg_color="#3e1b1b")
                    err_fr.pack(fill="x", pady=1, padx=20)
                    ctk.CTkLabel(err_fr, text=f"⚠️ Error in row '{r_item[0]}: {str(row_err)}'", 
                                 font=_BODY_FONT(10), text_color="#ff5555").pack(side="left", padx=10)
                    logger.error(f"Projection rendering error: {row_err}")

        except Exception as e:
            # Clean up and show error
            for w in self.p_scroll.winfo_children(): w.destroy()
            err_fr = ctk.CTkFrame(self.p_scroll, fg_color="#3e1b1b", border_color="#e74c3c", border_width=1)
            err_fr.pack(fill="x", pady=20, padx=20)
            
            ctk.CTkLabel(err_fr, text="⚠️ CALCULATION ENGINE ERROR", 
                          font=_BODY_FONT(14, bold=True), text_color="#ff5555").pack(pady=(15, 5))
            
            msg = f"An error occurred while generating projections: {str(e)}\n\nPossible Causes:\n- Missing Revenue or Profit inputs in Past Financials.\n- Zero tenure or interest rates in Project/Loan data.\n- Corrupted project file data.\n\nTry checking your inputs and switching back to this tab."
            
            ctk.CTkLabel(err_fr, text=msg, font=_BODY_FONT(12), text_color="white", 
                          wraplength=800, justify="left").pack(pady=(0, 20), padx=20)
            logger.exception("Projections tab crash:")

    def _check_net_worth_risk(self, results):
        """Identifies the first year where Tangible Net Worth falls below zero."""
        for res in results:
            if res.get("is_actual"): continue
            
            nw = float(res.get("share_capital", 0.0) or 0.0) + float(res.get("reserves_surplus", 0.0) or 0.0)
            if nw <= 0:
                return res.get("year_label", "First Year")
        return None

    # ════════════════════════════════════════════════════════════
    # TAB 7 — NARRATIVE & STYLE
    # ════════════════════════════════════════════════════════════
    def _tab_narrative(self):
        scroll = ctk.CTkScrollableFrame(self._tab_container, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        # Industry Category Mapping
        cat_fr = ctk.CTkFrame(scroll, **_CARD_KW)
        cat_fr.pack(fill="x", pady=10, padx=2)
        
        ctk.CTkLabel(cat_fr, text="Business Category & Auto-Image", font=_HDR_FONT(14)).pack(anchor="w", padx=15, pady=(10, 5))
        ctk.CTkLabel(cat_fr, text="This determines the narrative templates and cover images.", 
                      font=_BODY_FONT(11), text_color=_MUTED).pack(anchor="w", padx=15)
        
        self.cat_var = tk.StringVar(value=self.project.profile.business_category)
        cat_menu = ctk.CTkOptionMenu(cat_fr, variable=self.cat_var, 
                                     values=[c.value for c in BusinessCategory],
                                     height=38, corner_radius=8, fg_color=_BG1, button_color=_ACCENT, 
                                     command=self._on_category_change)
        cat_menu.pack(fill="x", padx=15, pady=15)

        # Image Preview
        self.img_preview_lbl = ctk.CTkLabel(cat_fr, text="Image Preview", height=200, fg_color=_BG1, corner_radius=10)
        self.img_preview_lbl.pack(fill="x", padx=15, pady=(0, 15))
        self._update_image_preview()

        # Business Description
        ctk.CTkLabel(scroll, text="Custom Business Description (Optional)", font=_HDR_FONT(14)).pack(anchor="w", pady=(15, 5))
        self.desc_text = tk.Text(scroll, height=6, font=(Theme.FONT_FAMILY, 11), bg=_BG2, fg="white", 
                                 insertbackground="white", borderwidth=1, relief="flat")
        self.desc_text.pack(fill="x", pady=5)
        self.desc_text.insert("1.0", self.project.profile.description)

    def _on_category_change(self, val):
        self.project.profile.business_category = val
        self._update_image_preview()
        
        # Phase 5: Smart Narrative Defaults
        defaults = {
            "Bakery / Food Production": "State-of-the-art bakery unit focusing on high-quality baked goods, cakes, and confectionery items utilizing modern ovens and hygiene-first production lines.",
            "CNC Workshop / Fabrication": "Precision engineering workshop equipped with advanced CNC machinery for high-accuracy metal fabrication, components, and industrial assemblies.",
            "HVAC / Ducting": "Specialized HVAC fabrication unit providing comprehensive ducting solutions, air handling systems, and climate control infrastructure for commercial projects.",
            "Retail / Trading": "Modern retail establishment focused on efficient supply chain management and customer-centric product distribution in the local market.",
            "Pharmaceuticals": "Compliance-driven pharmaceutical manufacturing/distribution unit ensuring rigid quality standards and efficient logistics for healthcare products.",
        }
        
        current_desc = self.desc_text.get("1.0", "end-1c").strip()
        if not current_desc or current_desc.startswith("State-of-the-art") or current_desc.startswith("Precision") or current_desc.startswith("Specialized"):
            self.desc_text.delete("1.0", "end")
            self.desc_text.insert("1.0", defaults.get(val, f"Professional {val} enterprise focused on quality, growth, and operational excellence in the target market."))

    def _update_image_preview(self):
        img_path = ImageMappingService.get_image_for_category(self.cat_var.get())
        if os.path.exists(img_path):
            from PIL import Image
            try:
                ctk_img = ctk.CTkImage(light_image=Image.open(img_path),
                                      dark_image=Image.open(img_path),
                                      size=(500, 200))
                self.img_preview_lbl.configure(image=ctk_img, text="")
            except:
                self.img_preview_lbl.configure(text="Preview Not Available", image=None)
        else:
            self.img_preview_lbl.configure(text=f"Default Image will be used\n({os.path.basename(img_path)})", image=None)

    # ════════════════════════════════════════════════════════════
    # TAB 8 & 9 — ANALYTICS (NEW)
    # ════════════════════════════════════════════════════════════
    def _tab_mpbf_analysis(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)

        try:
            self._sync_all_data()
            from services.cma.projection_engine_service import ProjectionEngineService
            from services.cma.mpbf_service import MpbfService
            
            projections = ProjectionEngineService.generate_full_projections(self.project)
            res = MpbfService.calculate_mpbf(self.project, projections)
            
            if res.get("status") == "error":
                raise ValueError(res.get("message", "Unknown error in MPBF service."))

            hdr = ctk.CTkFrame(container, **_CARD_KW)
            hdr.pack(fill="x", pady=10)
            
            ctk.CTkLabel(hdr, text="MPBF / Working Capital Assessment", font=_HDR_FONT(18)).pack(pady=(15, 5))
            risk_color = Theme.ACCENT_GREEN if res["risk_status"] == "Green" else (Theme.ACCENT_AMBER if res["risk_status"] == "Yellow" else "#e74c3c")
            ctk.CTkLabel(hdr, text=res["risk_message"], font=_BODY_FONT(12), text_color=risk_color).pack(pady=(0, 15))

            # Comparison Grid
            grid = ctk.CTkFrame(container, fg_color="transparent")
            grid.pack(fill="x", pady=10)
            grid.grid_columnconfigure((0, 1, 2), weight=1)

            # Logic: Headroom/Variance
            headroom = res['permissible_limit'] - res['requested_limit']
            if headroom >= 0:
                var_label = "Headroom Available"
                var_color = _GREEN
            else:
                var_label = "Gap / Over-limit"
                var_color = "#e74c3c" # Red
                
            # Extract short method name
            m_tag = res.get('method_used', 'Nayak').split('(')[0].strip()

        except Exception as e:
            err_fr = ctk.CTkFrame(container, fg_color="#3e1b1b", border_color="#e74c3c", border_width=1)
            err_fr.pack(fill="x", pady=50, padx=20)
            ctk.CTkLabel(err_fr, text="⚠️ MPBF ASSESSMENT ERROR", font=_BODY_FONT(14, bold=True), text_color="#ff5555").pack(pady=(15, 5))
            msg = f"Failed to perform MPBF analysis: {str(e)}\n\nThis usually happens if Projections could not be calculated first. Ensure Turnover and Asset values are correctly populated in previous tabs."
            ctk.CTkLabel(err_fr, text=msg, font=_BODY_FONT(12), text_color="white", wraplength=800).pack(pady=(0, 20), padx=20)
            logger.exception("MPBF tab crash:")
            return

        metrics = [
            ("Requested CC Limit", f"Rs. {res['requested_limit']:.2f}", _ACCENT),
            ("Adopted Limit for Assessment", f"Rs. {res['permissible_limit']:.2f}\n(Adopted: {m_tag})", _GREEN),
            (var_label, f"Rs. {abs(headroom):.2f}", var_color)
        ]
        
        for i, (l, v, c) in enumerate(metrics):
            f = ctk.CTkFrame(grid, **_CARD_KW)
            f.grid(row=0, column=i, padx=5, sticky="nsew")
            ctk.CTkLabel(f, text=l, font=_BODY_FONT(11), text_color=_MUTED).pack(pady=(10, 0))
            ctk.CTkLabel(f, text=v, font=_HDR_FONT(20), text_color=c).pack(pady=(0, 10))

        # Method Breakdown
        scroll = ctk.CTkScrollableFrame(container, fg_color=_BG1, border_width=1, border_color=Theme.BORDER_COLOR)
        scroll.pack(fill="both", expand=True, pady=10)
        
        for k, m in res["methods"].items():
            row = ctk.CTkFrame(scroll, fg_color=_BG2, corner_radius=10)
            row.pack(fill="x", pady=5, padx=5)
            ctk.CTkLabel(row, text=m["name"], font=_BODY_FONT(13, bold=True)).grid(row=0, column=0, padx=15, pady=10, sticky="w")
            
            ctk.CTkLabel(row, text=f"Gap: {m['gap']:.2f} | Margin: {m['margin']:.2f} | Permissible: {m['limit']:.2f}", 
                          font=_BODY_FONT(12), text_color=_MUTED).grid(row=0, column=1, padx=20, pady=10, sticky="e")

    def _tab_readiness_check(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)

        self._sync_all_data()
        from services.cma.projection_engine_service import ProjectionEngineService
        from services.cma.readiness_service import ReadinessService
        
        projections = ProjectionEngineService.generate_full_projections(self.project)
        # Upgrade: Using the new evaluate_readiness logic
        res = ReadinessService.evaluate_readiness(self.project, projections)

        # Premium Status Header
        score_fr = ctk.CTkFrame(container, **_CARD_KW)
        score_fr.pack(fill="x", pady=5)
        
        # Color Mapping for the new levels
        final_color = Theme.ACCENT_GREEN
        if "HIGH RISK" in res["readiness_level"]: final_color = "#e74c3c"
        elif "MODERATE" in res["readiness_level"]: final_color = Theme.ACCENT_AMBER
        
        hdr_fr = ctk.CTkFrame(score_fr, fg_color="transparent")
        hdr_fr.pack(fill="x", padx=25, pady=20)
        
        l_fr = ctk.CTkFrame(hdr_fr, fg_color="transparent")
        l_fr.pack(side="left")
        ctk.CTkLabel(l_fr, text="Bank-Readiness Health Check", font=_BODY_FONT(13), text_color=_MUTED).pack(anchor="w")
        ctk.CTkLabel(l_fr, text=res["readiness_level"], font=_HDR_FONT(26), text_color=final_color).pack(anchor="w")
        
        r_fr = ctk.CTkFrame(hdr_fr, fg_color="transparent")
        r_fr.pack(side="right")
        ctk.CTkButton(r_fr, text="🔄 Refresh Audit", width=160, height=38, corner_radius=8,
                      fg_color=_ACCENT, hover_color=_HOVER,
                      command=self._tab_readiness_check).pack()

        # Checks List
        scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        scroll.pack(fill="both", expand=True, pady=10)

        ctk.CTkLabel(scroll, text="DETAILED CREDIT PARAMETERS & ADVICE", font=_BODY_FONT(11, bold=True), text_color=_MUTED).pack(anchor="w", padx=10, pady=(0, 10))

        if not res["checks"]:
            ctk.CTkLabel(scroll, text="No critical data found to audit. Please fill Setup and Projections.", 
                         font=_BODY_FONT(14), text_color=_MUTED).pack(pady=40)
        else:
            for c in res["checks"]:
                row = ctk.CTkFrame(scroll, **_CARD_KW)
                row.pack(fill="x", pady=4, padx=5)
                
                c_color = Theme.ACCENT_GREEN if c["level"] == "PASS" else (Theme.ACCENT_AMBER if c["level"] == "WARNING" else "#e74c3c")
                
                row.grid_columnconfigure(0, minsize=180) # Label
                row.grid_columnconfigure(1, minsize=140) # Value
                row.grid_columnconfigure(2, weight=1)    # Advice
                
                l_info = ctk.CTkFrame(row, fg_color="transparent")
                l_info.grid(row=0, column=0, sticky="nsw", padx=(15, 0), pady=12)
                ctk.CTkLabel(l_info, text="●", text_color=c_color, font=_HDR_FONT(20)).pack(side="left", padx=(0, 5))
                ctk.CTkLabel(l_info, text=c["name"], font=_BODY_FONT(13, bold=True), anchor="w", wraplength=140).pack(side="left")
                
                v_fr = ctk.CTkFrame(row, fg_color="transparent")
                v_fr.grid(row=0, column=1, sticky="nsew", padx=10)
                ctk.CTkLabel(v_fr, text=c["value"], font=_HDR_FONT(16), text_color=_ACCENT).pack(expand=True)
                
                adv_fr = ctk.CTkFrame(row, fg_color="transparent")
                adv_fr.grid(row=0, column=2, sticky="nsew", padx=(10, 20), pady=12)
                ctk.CTkLabel(adv_fr, text="Improvement Advice:", font=_BODY_FONT(10, bold=True), text_color=_MUTED).pack(anchor="w")
                ctk.CTkLabel(adv_fr, text=c["advice"], font=_BODY_FONT(11), text_color="white", wraplength=450, justify="left").pack(anchor="w", fill="x")

        # Legal Note
        ctk.CTkLabel(container, text="⚠️ Note: Audit parameters are based on standard banking benchmarks and do not guarantee sanction.", 
                     font=_BODY_FONT(11), text_color=_MUTED, wraplength=800).pack(pady=10)
    # ════════════════════════════════════════════════════════════
    def _tab_generate(self):
        container = ctk.CTkFrame(self._tab_container, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")

        # Top Section
        top = ctk.CTkFrame(scroll, fg_color="transparent")
        top.pack(pady=20)
        ctk.CTkLabel(top, text="🚀 Premium Report Publishing", font=_HDR_FONT(26)).pack()
        ctk.CTkLabel(top, text="Finalize branding and export your bank-ready document.", font=_BODY_FONT(13), text_color=_MUTED).pack(pady=5)
        
        # Branding Panel (New)
        brand_fr = ctk.CTkFrame(scroll, **_CARD_KW)
        brand_fr.pack(fill="x", padx=40, pady=10)
        
        ctk.CTkLabel(brand_fr, text="🏢 Optional CA Branding Panel", font=_HDR_FONT(16)).pack(anchor="w", padx=20, pady=(15, 5))
        ctk.CTkLabel(brand_fr, text="Configure how your firm appears on the cover and footer.", font=_BODY_FONT(11), text_color=_MUTED).pack(anchor="w", padx=20)
        
        branding_grid = ctk.CTkFrame(brand_fr, fg_color="transparent")
        branding_grid.pack(fill="x", padx=20, pady=15)
        branding_grid.grid_columnconfigure((0, 2), weight=0)
        branding_grid.grid_columnconfigure((1, 3), weight=1)
        
        self.brand_entries = {}
        fields = [
            ("CA Firm Name", "firm_name"), ("Prepared By", "prepared_by"),
            ("Contact Line", "contact_line")
        ]
        
        for i, (lbl, key) in enumerate(fields):
            r, c = divmod(i, 2)
            ctk.CTkLabel(branding_grid, text=lbl, font=_BODY_FONT(12, bold=True)).grid(row=r, column=c*2, sticky="w", padx=10, pady=8)
            
            ent_fr = ctk.CTkFrame(branding_grid, fg_color="transparent")
            ent_fr.grid(row=r, column=c*2+1, sticky="ew", padx=10, pady=8)
            
            ent = ctk.CTkEntry(ent_fr, height=36, corner_radius=8, fg_color=_BG1)
            ent.pack(side="left", fill="x", expand=True)
            ent.insert(0, getattr(self.project.branding, key, ""))
            self.brand_entries[key] = ent
            
            if key == "logo_path":
                btn = ctk.CTkButton(ent_fr, text="📁 Pick...", width=80, height=36, corner_radius=8,
                                    fg_color=_BG2, hover_color=_ACCENT, 
                                    command=lambda: self._on_pick_logo(self.brand_entries["logo_path"]))
                btn.pack(side="left", padx=(5, 0))

        dis_fr = ctk.CTkFrame(brand_fr, fg_color="transparent")
        dis_fr.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkLabel(dis_fr, text="Custom Disclaimer / Note", font=_BODY_FONT(12, bold=True)).pack(anchor="w", padx=10)
        self.brand_disclaimer = ctk.CTkEntry(dis_fr, height=36, corner_radius=8, fg_color=_BG1, placeholder_text="e.g. For Internal Banking Use Only")
        self.brand_disclaimer.pack(fill="x", padx=10, pady=5)
        self.brand_disclaimer.insert(0, self.project.branding.disclaimer)

        # Action Section
        act_fr = ctk.CTkFrame(scroll, fg_color=_BG2, corner_radius=15, border_width=1, border_color=Theme.BORDER_COLOR)
        act_fr.pack(pady=20, padx=40, fill="x")
        
        self.gen_btn = ctk.CTkButton(act_fr, text="📊 Generate Premium PDF", width=300, height=54,
                                     corner_radius=12, fg_color=_GREEN, hover_color="#059669",
                                     font=_HDR_FONT(18), command=self._start_generation)
        self.gen_btn.pack(pady=(25, 8))
        
        self.word_gen_btn = ctk.CTkButton(act_fr, text="📝 Generate Editable Word", width=300, height=54,
                                          corner_radius=12, fg_color="#3B82F6", hover_color="#2563EB",
                                          font=_HDR_FONT(18), command=self._start_word_generation)
        self.word_gen_btn.pack(pady=(8, 25))
        
        self.gen_status = ctk.CTkLabel(scroll, text="Ready for publishing", font=_BODY_FONT(12), text_color=_MUTED)
        self.gen_status.pack(pady=10)

    def _on_pick_logo(self, entry):
        path = filedialog.askopenfilename(
            title="Select Firm Logo",
            filetypes=[("Images", "*.jpg *.jpeg *.png"), ("All Files", "*.*")]
        )
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _start_generation(self):
        self._sync_all_data()
        
        if not self.project.profile.business_name:
            messagebox.showwarning("Incomplete Data", "Please enter a Business Name in 'Party Master' tab.")
            self._show_tab(1)
            return

        # --- MANDATORY VALIDATION GATE (New) ---
        try:
            from services.cma.projection_engine_service import ProjectionEngineService
            from services.cma.models import ReportMode, LoanType
            
            proj_data = ProjectionEngineService.generate_full_projections(self.project)
            ProjectionEngineService.validate_projections(proj_data)
            
            # Additional Mode Logic Check (Point 12)
            rm = self.project.profile.report_mode
            lt = self.project.profile.loan_type
            if rm == ReportMode.CMA.value:
                is_wc_only = (lt in [LoanType.OD_LIMIT.value, LoanType.RENEWAL.value, LoanType.WORKING_CAPITAL.value])
                if is_wc_only and self.project.loan.term_loan_amount > 0:
                    messagebox.showerror("Mode Consistency Error", 
                                        "You have selected 'CMA Mode' for a Working Capital facility, but the project contains 'Term Loan' amounts.\n\n"
                                        "Please either change Report Mode to 'Pro' (Composite) or set Term Loan amount to 0 in 'Project & Finance' tab.")
                    return
        except ValueError as ve:
            messagebox.showerror("Financial Inconsistency Detected", 
                                f"Unable to generate report due to data mismatches:\n\n{str(ve)}\n\n"
                                "Please review your Assets, Loans, and Assumptions to ensure the Balance Sheet tallies.")
            return
            # Bankability Soft-Warning Gate (New)
            from services.cma.readiness_service import ReadinessService
            readiness = ReadinessService.evaluate_readiness(self.project, proj_data)
            if readiness.get("critical_count", 0) > 0:
                msg = (f"🚨 CRITICAL BANKABILITY RISKS DETECTED\n\n"
                       f"Your project has {readiness['critical_count']} critical credit flags (e.g., low DSCR or negative net worth).\n\n"
                       f"While the software allows generation, banks are likely to REJECT this proposal as currently structured.\n\n"
                       f"Are you sure you want to proceed with 'High Risk' report generation?")
                if not messagebox.askyesno("Bankability Alert", msg):
                    self._show_tab(7) # Redirect to Readiness Tab
                    return

        out_path = filedialog.asksaveasfilename(
            title="Export Premium PDF Report",
            defaultextension=".pdf",
            filetypes=[("Professional PDF", "*.pdf")],
            initialfile=f"DPR_{self.project.profile.business_name.replace(' ', '_')}_Final.pdf"
        )
        if not out_path: return

        self.gen_btn.configure(state="disabled", text="🏗️ Building PDF...")
        self.gen_status.configure(text="Generating multi-pass high-fidelity report...", text_color=_ACCENT)
        threading.Thread(target=self._gen_thread, args=(out_path, "pdf"), daemon=True).start()

    def _start_word_generation(self):
        self._sync_all_data()
        
        if not self.project.profile.business_name:
            messagebox.showwarning("Incomplete Data", "Please enter a Business Name in 'Party Master' tab.")
            self._show_tab(1)
            return

        out_path = filedialog.asksaveasfilename(
            title="Export Editable Word Report",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialfile=f"DPR_{self.project.profile.business_name.replace(' ', '_')}.docx"
        )
        if not out_path: return

        self.word_gen_btn.configure(state="disabled", text="Generating Word...")
        threading.Thread(target=self._gen_thread, args=(out_path, "word"), daemon=True).start()

    def _gen_thread(self, path, mode):
        try:
            if mode == "pdf":
                ReportGeneratorService.generate_pdf(self.project, path)
            else:
                WordGeneratorService.generate_docx(self.project, path)

            # Add to history
            version = {
                "version_id": str(uuid.uuid4())[:6],
                "mode": mode.upper(),
                "output_pdf_path": path,
                "generated_on": datetime.now().strftime("%Y-%m-%d %H:%M")
            }
            # Save project state
            self.after(0, lambda: self._on_gen_success(path, mode))
        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._on_gen_error(err_msg))

    def _on_gen_success(self, path, mode):
        self.gen_btn.configure(state="normal", text="⚡ Generate PDF Report")
        self.word_gen_btn.configure(state="normal", text="📝 Generate Word Document")
        
        self.gen_status.configure(text=f"{mode.upper()} Generated Successfully!\nSaved to: {os.path.basename(path)}", text_color=_GREEN)
        self._save_project()
        os.startfile(path)

    def _on_gen_error(self, msg):
        self.gen_btn.configure(state="normal", text="⚡ Generate PDF Report")
        self.word_gen_btn.configure(state="normal", text="📝 Generate Word Document")
        if "Permission denied" in msg or "PermissionError" in msg:
            messagebox.showerror("File Locked", "Wait! The file is open in another program (Acrobat or Word).\n\nPlease CLOSE the file and try again.")
        else:
            messagebox.showerror("Generation Error", f"Failed to generate report: {msg}")

    # ════════════════════════════════════════════════════════════
    # CORE LOGIC
    # ════════════════════════════════════════════════════════════
    def _sync_all_data(self):
        """Syncs all UI entry fields into the project model safely."""
        # Sync Setup Tab (New)
        try:
            if hasattr(self, 'mode_var'):
                self.project.profile.business_mode = self.mode_var.get()
            if hasattr(self, 'loan_type_var'):
                self.project.profile.loan_type = self.loan_type_var.get()
            if hasattr(self, 'scheme_var'):
                self.project.profile.scheme_type = self.scheme_var.get()
            if hasattr(self, 'report_mode_var'):
                self.project.profile.report_mode = self.report_mode_var.get()
        except: pass

        # Tab 2: Party Master
        try:
            if hasattr(self, '_party_entries'):
                for k, ent in self._party_entries.items():
                    try:
                        val = ent.get()
                        if k == "employee_count":
                            try: setattr(self.project.profile, k, int(val))
                            except: pass
                        elif k == "security_value":
                            try: setattr(self.project.profile, k, float(val))
                            except: pass
                        else:
                            setattr(self.project.profile, k, val)
                    except Exception: pass # Widget destroyed
                if hasattr(self, 'entity_type_var'):
                    self.project.profile.entity_type = self.entity_type_var.get()
                if hasattr(self, 'business_category_var'):
                    self.project.profile.business_category = self.business_category_var.get()
                if hasattr(self, 'manual_category_var'):
                    self.project.profile.manual_business_category = self.manual_category_var.get()
        except Exception: pass

        # Tab: Audited Data (Simplified)
        try:
            if hasattr(self, '_simplified_entries'):
                for k, ent in self._simplified_entries.items():
                    try: setattr(self.project.simplified_data, k, float(ent.get() or 0))
                    except: pass
        except: pass

        # Tab 3: Loan & Assets
        try:
            if hasattr(self, 'loan_purpose_ent'):
                try: self.project.loan.purpose = self.loan_purpose_ent.get()
                except Exception: pass
            if hasattr(self, 'loan_entries'):
                for k, ent in self.loan_entries.items():
                    try:
                        val = ent.get()
                        try:
                            if "year" in k: setattr(self.project.loan, k, int(float(val)))
                            else: setattr(self.project.loan, k, float(val))
                        except: pass
                    except Exception: pass
            
            if hasattr(self, 'facility_type_var'):
                self.project.loan.facility_type = self.facility_type_var.get()
        except Exception: pass
        
        # Sync Dynamic Assets
        try:
            if hasattr(self, '_asset_widgets'):
                for name_ent, cost_ent, asset in self._asset_widgets:
                    try:
                        asset.name = name_ent.get()
                        asset.cost = float(cost_ent.get() or 0.0)
                    except Exception: pass
        except Exception: pass
        
        # Sync Audited Data
        try:
            if hasattr(self, '_audited_entries'):
                for i, row_map in enumerate(self._audited_entries):
                    ad = self.project.audited_history[i]
                    # Numeric fields
                    for k, ent in row_map.items():
                        try: setattr(ad, k, float(ent.get() or 0.0))
                        except: pass
                    # Meta fields (Year Label, Status) from _audited_meta
                    if hasattr(self, '_audited_meta') and i < len(self._audited_meta):
                        lbl_ent, status_var = self._audited_meta[i]
                        ad.year_label = lbl_ent.get()
                        ad.data_type = status_var.get()
                    
                    # Automate totals calculations for model consistency
                    l_keys = ["share_capital", "reserves_surplus", "term_loans", "unsecured_loan", "bank_od",
                              "creditors", "other_loans_liabilities", "provisions", "other_current_liabilities"]
                    a_keys = ["net_block", "investments", "inventory", "debtors", 
                              "loans_advances", "deposits", "other_current_assets", "cash_bank"]
                    
                    ad.current_liabilities = sum(getattr(ad, k, 0) for k in ["creditors", "other_current_liabilities", "provisions"])
                    ad.current_assets = sum(getattr(ad, k, 0) for k in ["inventory", "debtors", "other_current_assets", "loans_advances", "deposits", "cash_bank"])
        except Exception: pass

        # Tab 4: Assumptions & Scenarios
        try:
            self.project.assumptions.projection_years = int(self.years_var.get())
            self.project.assumptions.depreciation_method = self.depr_meth_var.get()
            self.project.assumptions.selected_scenario = self.scenario_var.get()
            if hasattr(self, '_ass_entries'):
                for k, ent in self._ass_entries.items():
                    try:
                        val = ent.get()
                        if "days" in k: setattr(self.project.assumptions, k, int(float(val)))
                        else: setattr(self.project.assumptions, k, float(val))
                    except Exception: pass
        except Exception: pass

        # Tab 5: Narrative
        try:
            if hasattr(self, 'cat_var'):
                self.project.profile.business_category = self.cat_var.get()
            if hasattr(self, 'desc_text'):
                try: self.project.profile.description = self.desc_text.get("1.0", "end-1c")
                except Exception: pass
        except Exception: pass

        # Branding Panel
        try:
            if hasattr(self, 'brand_entries'):
                for k, ent in self.brand_entries.items():
                    setattr(self.project.branding, k, ent.get())
                self.project.branding.disclaimer = self.brand_disclaimer.get()
        except: pass

    def _save_project(self):
        self._sync_all_data()
        try:
            PartyMasterService.save_project(self.project)
            self._status_lbl.configure(text=f"Project Saved: {datetime.now().strftime('%H:%M:%S')}")
            self.party_info_lbl.configure(text=f"Party: {self.project.profile.business_name}")
            self._refresh_project_list()
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def _save_party_details(self):
        """Syncs the party profile fields and saves them, returning to dashboard."""
        if hasattr(self, '_party_entries'):
            for k, ent in self._party_entries.items():
                try:
                    val = ent.get()
                    if k == "employee_count":
                        setattr(self.project.profile, k, int(val or 0))
                    elif k == "security_value":
                        setattr(self.project.profile, k, float(val or 0.0))
                    else:
                        setattr(self.project.profile, k, val)
                except Exception: pass
            
            if hasattr(self, 'entity_type_var'):
                self.project.profile.entity_type = self.entity_type_var.get()
        
        self._save_project()
        self._show_tab(0) # Return to dashboard showing the updated project card

    def _new_project(self):
        if self.project.profile.business_name and not messagebox.askyesno("Confirm", "Create new project? Unsaved changes will be lost."):
            return
        self.project = CmaProject()
        self.party_info_lbl.configure(text="New Project")
        self._edit_party_details() # Show party master details entry immediately

    def _load_project_by_path(self, path):
        try:
            self.project = PartyMasterService.load_project(path)
            self.party_info_lbl.configure(text=f"Party: {self.project.profile.business_name}")
            self._show_tab(0) # Stay on dashboard to show Edit/Setup buttons
            self._status_lbl.configure(text="Project Loaded")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def _delete_project(self, pid):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this project?"):
            if PartyMasterService.delete_project(pid):
                self._refresh_project_list()
                self._status_lbl.configure(text="Project Deleted")

    # ════════════════════════════════════════════════════════════
    # CONNECTIVITY LOGIC (Phase 4)
    # ════════════════════════════════════════════════════════════
