import customtkinter as ctk
import sys
from tkinterdnd2 import TkinterDnD
import tkinter.messagebox as tkmb
import os
from PIL import Image
from utils.license_manager import LicenseManager 
from utils.file_manager import FileManager
from ui.theme import Theme
from ui.components import NavButton, InstructionDialog
import threading
import time

# SET TO False to hide Advanced modules (Standard Version)
IS_PREMIUM_BUILD = True

class TkinterDnDApp(ctk.CTk, TkinterDnD.DnDWrapper):
    """
    Combines CustomTkinter with tkinterdnd2 capabilities for native
    drag-and-drop support.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class AppWindow(TkinterDnDApp):
    def __init__(self):
        super().__init__()
        
        self.title("CA Office PDF Utility")
        self.geometry("1100x750")
        self.minsize(900, 600)
        
        # Protocol for closing
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Initialize attributes
        self.content_frame = None
        self.nav_buttons = {}
        self.seat_manager = None # Replaces license_server and license_client
        
        # Maximize window on startup
        self.after(200, lambda: self._maximize_window())
        
    def _maximize_window(self):
        try:
            self.state("zoomed")
        except:
            pass
        
        # Set theme
        Theme.apply_to_ctk(ctk)
        self.configure(fg_color=Theme.BG_PRIMARY)
        
        # Configure grid exactly to place sidebar left and content right
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0, minsize=240) # Sidebar column (fixed)
        self.grid_columnconfigure(1, weight=1) # Content column (flexible)
        
        # CRITICAL: Setup content frame BEFORE sidebar
        self.setup_content_frame()
        self.setup_sidebar()
        
        # License Check
        self.license_status = LicenseManager.get_status()
        
        # Initialize Peer-to-Peer Seat Manager
        from utils.license_manager import FloatingSeatManager
        
        if self.license_status["is_office"] and self.license_status["is_activated"]:
            # PC is an Activated Office Master
            self.seat_manager = FloatingSeatManager(
                self.license_status["machine_id"],
                self.license_status["member_no"],
                self.license_status["seats"]
            )
        else:
            # PC is a Standalone or Unactivated Client
            # It will try to discover an Office network on the LAN
            self.seat_manager = FloatingSeatManager(self.license_status["machine_id"])
            
        self.seat_manager.start()
        
        # Always show the Office License Server button for network seat monitoring
        self.add_server_mgmt_btn()
        
        # Show default view or activation if expired
        # Only block if EXPIRED AND no floating seat found (Wait a bit for discovery)
        def initial_view_check():
            if self.license_status["expired"] and not self.seat_manager.has_seat:
                self.show_view("activation")
            else:
                self.show_view("merger")
        
        self.after(1500, initial_view_check) # Small delay to allow LAN discovery

    def add_server_mgmt_btn(self):
        """Adds a button to manage the office license seats in the management section."""
        if hasattr(self, "mgmt_fr"):
            from ui.components import NavButton
            btn = NavButton(self.mgmt_fr, text="Office License Server", icon="🏢", 
                             command=lambda: self.show_view("server_mgmt"))
            btn.pack(side="top", fill="x", padx=0, pady=2)
            self.nav_buttons["server_mgmt"] = btn

    def setup_sidebar(self):
        """Sets up the left navigation sidebar with persistent top and bottom elements."""
        self.sidebar_frame = ctk.CTkFrame(self, width=240, corner_radius=0, 
                                          fg_color=Theme.BG_SECONDARY)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.pack_propagate(False) # Keep width fixed at 240
        
        # Border-like separation on the right
        sep = ctk.CTkFrame(self.sidebar_frame, width=1, fg_color=Theme.BORDER_COLOR)
        sep.pack(side="right", fill="y")
        
        # 1. Top Logo Frame
        logo_container = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        logo_container.pack(side="top", fill="x", pady=(20, 10))
        
        try:
            logo_path = FileManager.get_resource_path(os.path.join("assets", "logo.png"))
            if os.path.exists(logo_path):
                logo_image = ctk.CTkImage(light_image=Image.open(logo_path),
                                        dark_image=Image.open(logo_path),
                                        size=(100, 100))
                logo_label = ctk.CTkLabel(logo_container, image=logo_image, text="")
            else:
                raise FileNotFoundError("Logo file not found.")
        except Exception:
            logo_label = ctk.CTkLabel(logo_container, text="CA Office\nPDF Utility", 
                                      font=ctk.CTkFont(size=22, weight="bold"))
        logo_label.pack(pady=5)
        
        # 2. Bottom Frame (Persistent)
        self.bottom_sidebar = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.bottom_sidebar.pack(side="bottom", fill="x", padx=10, pady=(5, 15))
        
        # Helper to create uniform buttons
        self.nav_buttons = {}
        def add_nav_btn(master, name, text, icon, is_activation=False, is_premium=False):
            btn = NavButton(master, text=text, icon=icon, 
                             command=lambda n=name: self.show_view(n),
                             is_activation=is_activation, is_premium=is_premium)
            btn.pack(side="top", fill="x", padx=0, pady=2)
            self.nav_buttons[name] = btn
            return btn
            
        # Activation Button (Fixed at bottom)
        self.btn_activation = add_nav_btn(self.bottom_sidebar, "activation", "Activation Key", "🔑", is_activation=True)
        
        # Branding
        branding_lbl = ctk.CTkLabel(self.bottom_sidebar, 
                                    text="Developed by CA Bhavesh Lunagariya", 
                                    font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=12, weight="bold"),
                                    text_color=Theme.TEXT_PRIMARY)
        branding_lbl.pack(side="top", pady=(10, 0))
        
        # 3. Main Navigation (Scrollable)
        self.nav_scroll = ctk.CTkScrollableFrame(self.sidebar_frame, fg_color="transparent", corner_radius=0)
        self.nav_scroll.pack(side="top", fill="both", expand=True, padx=4)
        
        add_nav_btn(self.nav_scroll, "merger", "PDF Merger", "📄")
        add_nav_btn(self.nav_scroll, "img_to_pdf", "Image to PDF", "🖼️")
        add_nav_btn(self.nav_scroll, "page_mgmt", "Page Management", "📑")
        add_nav_btn(self.nav_scroll, "comp_center", "Compression Center", "🗜️")
        add_nav_btn(self.nav_scroll, "security", "Security & Watermark", "🔒")
        add_nav_btn(self.nav_scroll, "ocr", "OCR Extractor", "🔍")
        add_nav_btn(self.nav_scroll, "pdf_editor", "PDF Editor", "📝")
        
        # Bottom Managed Section (Separated from working modules)
        self.mgmt_sep = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color=Theme.BORDER_COLOR)
        self.mgmt_sep.pack(side="bottom", fill="x", padx=15, pady=(5, 5))
        
        self.mgmt_fr = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.mgmt_fr.pack(side="bottom", fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(self.mgmt_fr, text="SYSTEM & OFFICE", font=ctk.CTkFont(size=10, weight="bold"), 
                     text_color=Theme.TEXT_MUTED).pack(side="top", anchor="w", padx=10, pady=(0, 5))
        
        # Show premium modules in Premium builds
        if IS_PREMIUM_BUILD:
            add_nav_btn(self.nav_scroll, "invoice_parser", "Smart Invoice Parser", "🧾")
                
            add_nav_btn(self.nav_scroll, "sign_stamp", "Sign & Stamp", "🖊️")
            # add_nav_btn(self.nav_scroll, "gst", "GST Pack Builder", "📑")
            add_nav_btn(self.nav_scroll, "cma_dpr", "CMA / DPR Builder", "📊", is_premium=True)
        else:
            # Standard Version only gets basic signing
            add_nav_btn(self.nav_scroll, "sign_stamp", "Sign & Stamp", "🖊️")

    def setup_content_frame(self):
        """Sets up the dynamic content area."""
        self.content_frame = ctk.CTkFrame(self, corner_radius=10, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

    def show_view(self, view_name):
        """Lazily load and switch to a different view."""
        self.current_view_name = view_name
        # Update sidebar highlighting
        for name, btn in self.nav_buttons.items():
            btn.set_active(name == view_name)

        # Safety check: Ensure content frame exists
        if not hasattr(self, "content_frame") or self.content_frame is None:
            return

        # Clear current view
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
        view = None
        try:
            if view_name == "merger":
                from ui.views.merger_view import MergerView
                view = MergerView(self.content_frame, self)
            elif view_name == "img_to_pdf":
                from ui.views.image_to_pdf_view import ImageToPdfView
                view = ImageToPdfView(self.content_frame, self)
            elif view_name == "page_mgmt":
                from ui.views.page_management_view import PageManagementView
                view = PageManagementView(self.content_frame, self)
            elif view_name == "comp_center":
                from ui.views.compression_center_view import CompressionCenterView
                view = CompressionCenterView(self.content_frame, self)
            elif view_name == "security":
                from ui.views.security_view import SecurityView
                view = SecurityView(self.content_frame, self)
            elif view_name == "invoice_parser":
                from ui.views.invoice_parser_view import InvoiceParserView
                view = InvoiceParserView(self.content_frame, self)
            elif view_name == "ocr":
                from ui.views.ocr_view import OcrView
                view = OcrView(self.content_frame, self)
            elif view_name == "pdf_editor":
                from ui.views.pdf_editor_view import PDFEditorView
                view = PDFEditorView(self.content_frame, self)
            elif view_name == "sign_stamp":
                from ui.views.sign_stamp_view import SignStampView
                view = SignStampView(self.content_frame, self)
            elif view_name == "cma_dpr" and IS_PREMIUM_BUILD:
                from ui.views.cma_dpr_builder_view import CmaDprBuilderView
                view = CmaDprBuilderView(self.content_frame, self)
            elif view_name == "gst" and IS_PREMIUM_BUILD:
                from ui.views.gst_pack_view import GstPackView
                view = GstPackView(self.content_frame, self)
            elif view_name == "invoice_parser" and IS_PREMIUM_BUILD:
                from ui.views.invoice_parser_view import InvoiceParserView
                view = InvoiceParserView(self.content_frame, self)
            elif view_name == "bank_statement":
                from ui.views.bank_statement_view import BankStatementView
                view = BankStatementView(self.content_frame, self)
            elif view_name == "activation":
                from ui.views.activation_view import ActivationView
                view = ActivationView(self.content_frame, self, on_success_callback=self.on_activation_success)
            elif view_name == "server_mgmt":
                from ui.views.license_server_view import LicenseServerView
                view = LicenseServerView(self.content_frame, self)
                
            if view:
                view.pack(fill="both", expand=True)
        except Exception as e:
            self.show_toast("Error Loading Module", str(e), is_error=True)

    def on_activation_success(self):
        """Called when user successfully activates."""
        self.license_status = LicenseManager.get_status()
        self.show_view("merger")
        self.show_toast("Activated", "Software is now fully unlocked.")

    def show_support_info(self):
        msg = "Need a key? Contact for activation:\nMob: 8200808507\nEmail: info.maruticonsultancy@gmail.com"
        self.show_toast("Support & Activation", msg)

    def check_operation_allowed(self):
        """
        Checks if the current operation is allowed based on local activation
        or network floating seats.
        """
        # 1. Local Activation (Standalone or Activated Office)
        self.license_status = LicenseManager.get_status()
        if self.license_status["is_activated"]:
            return True
        
        # 2. Office Machine (Client with network seat)
        if self.seat_manager and self.seat_manager.has_seat and self.seat_manager.is_connected:
            return True

        # 3. Trial Mode
        if self.license_status["expired"] and not (self.seat_manager and self.seat_manager.has_seat):
            self.show_view("activation")
            return False
            
        return True # Within trial period and not expired

    def show_toast(self, title, message, is_error=False):
        """Shows a popup message."""
        if is_error:
            tkmb.showerror(title, message)
        else:
            tkmb.showinfo(title, message)
            
    def confirm(self, title, message):
        """Shows a yes/no confirmation dialog."""
        return tkmb.askyesno(title, message)

    def on_closing(self):
        """Handles application shutdown and license cleanup."""
        try:
            if self.seat_manager:
                self.seat_manager.stop()
        except:
            pass
        self.destroy()
