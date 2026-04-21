import customtkinter as ctk
from ui.theme import Theme
from utils.license_manager import LicenseManager

class ActivationView(ctk.CTkScrollableFrame):
    def __init__(self, master, app_window, on_success_callback=None, **kwargs):
        super().__init__(master, fg_color=Theme.BG_PRIMARY, **kwargs)
        self.app = app_window
        self.on_success = on_success_callback
        
        self.grid_columnconfigure(0, weight=1)
        self.status = LicenseManager.get_status()
        
        # Header
        lbl_hdr = ctk.CTkLabel(self, text="Software Activation", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.pack(anchor="w", padx=40, pady=(20, 5))
        
        lbl_subtitle = ctk.CTkLabel(self, text="CA Office PDF Utility - Protection System", text_color="gray70")
        lbl_subtitle.pack(anchor="w", padx=40, pady=(0, 20))
        
        # Info Box
        info_frame = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS, border_width=1, border_color=Theme.BORDER_COLOR)
        info_frame.pack(padx=40, fill="x", pady=10)
        
        if self.status['is_activated']:
            expiry_str = self.status.get('expiry_date', '')
            try:
                from datetime import datetime
                expiry_dt = datetime.strptime(expiry_str, "%Y%m%d")
                formatted_date = expiry_dt.strftime("%d-%b-%Y")
            except:
                formatted_date = "Unknown"
                
            trial_text = f"Subscription Active\nExpires on: {formatted_date}\n({self.status['days_left']} days remaining)"
            color = "#2ecc71"
        elif self.status['expired']:
            trial_text = "Status: EXPIRED\nPlease contact for a new license key."
            color = "#e74c3c"
        else:
            trial_text = f"Trial Status: {self.status['days_left']} Days Remaining"
            color = "#f1c40f" if self.status['days_left'] <= 2 else None
            
        self.lbl_trial = ctk.CTkLabel(info_frame, text=trial_text, 
                                     font=ctk.CTkFont(size=13, weight="bold"),
                                     text_color=color, justify="center")
        self.lbl_trial.pack(pady=15)
        
        # Machine ID Box
        id_frame = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS, border_width=1, border_color=Theme.BORDER_COLOR)
        id_frame.pack(padx=40, fill="x", pady=20)
        
        ctk.CTkLabel(id_frame, text="Your Machine ID (Send this to CA Bhavesh):").pack(pady=(10, 5))
        
        self.ent_id = ctk.CTkEntry(id_frame, justify="center", font=ctk.CTkFont(size=14, weight="bold"))
        self.ent_id.insert(0, self.status['machine_id'])
        self.ent_id.configure(state="readonly")
        self.ent_id.pack(padx=20, pady=(0, 10), fill="x")
        
        btn_copy = ctk.CTkButton(id_frame, text="📋 Copy Machine ID", height=32, corner_radius=6,
                                 fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                                 font=(Theme.FONT_FAMILY, 12),
                                 command=self.copy_id)
        btn_copy.pack(pady=(0, 15))
        
        # Activation Card
        self.card = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CARD_CORNER_RADIUS)
        self.card.pack(padx=40, pady=20, fill="x")

        # CA Identification Toggle
        self.ca_frame = ctk.CTkFrame(self.card, fg_color="transparent")
        self.ca_frame.pack(pady=(10, 0))
        
        self.ca_var = ctk.BooleanVar(value=False)
        self.sw_ca = ctk.CTkSwitch(self.ca_frame, text="Are you a Chartered Accountant?", 
                                   variable=self.ca_var, command=self.toggle_ca_fields,
                                   font=(Theme.FONT_FAMILY, 13))
        self.sw_ca.pack()
        
        # Membership Number Field (Hidden by default)
        self.member_frame = ctk.CTkFrame(self.card, fg_color="transparent")
        # will be packed by toggle_ca_fields
        
        ctk.CTkLabel(self.member_frame, text="CA Membership Number:", font=(Theme.FONT_FAMILY, 13)).pack(side="left", padx=10)
        self.ent_member = ctk.CTkEntry(self.member_frame, placeholder_text="Enter M.No.", width=200,
                                       font=(Theme.FONT_FAMILY, 13), height=34, corner_radius=Theme.BUTTON_CORNER_RADIUS)
        self.ent_member.pack(side="left")

        # Key Entry
        ctk.CTkLabel(self.card, text="Enter License Key:", font=(Theme.FONT_FAMILY, 13)).pack(pady=(20, 5))
        self.ent_key = ctk.CTkEntry(self.card, width=350, height=44, corner_radius=8, 
                                   placeholder_text="XXXX-XXXX-XXXX-XXXX",
                                   border_color=Theme.BORDER_COLOR, fg_color=Theme.BG_PRIMARY, 
                                   font=(Theme.FONT_FAMILY, 14))
        self.ent_key.pack(pady=5)
        
        self.btn_activate = ctk.CTkButton(self.card, text="Activate Software", height=44, corner_radius=10,
                                          fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                          font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=15, weight="bold"), 
                                          command=self.activate)
        self.btn_activate.pack(pady=(20, 30))
        
        # Support Info
        support_frame = ctk.CTkFrame(self, fg_color="transparent")
        support_frame.pack(pady=20)
        
        ctk.CTkLabel(support_frame, text="Need a key? Contact for activation:", font=ctk.CTkFont(weight="bold")).pack()
        
        email_addr = "support.marutitechsolutions@gmail.com"
        btn_copy_email_only = ctk.CTkButton(support_frame, text=email_addr, 
                                            fg_color="transparent", 
                                            text_color="#5dade2", 
                                            font=ctk.CTkFont(size=13, weight="bold", underline=True),
                                            hover_color="#2c3e50", height=25,
                                            command=lambda: self.copy_text(email_addr, "Email copied!"))
        btn_copy_email_only.pack(pady=5)
        
        btn_grid = ctk.CTkFrame(support_frame, fg_color="transparent")
        btn_grid.pack(pady=10)
        
        btn_copy = ctk.CTkButton(btn_grid, text="📋 Copy Purchase Details", 
                                 width=200, height=45, font=ctk.CTkFont(size=13, weight="bold"),
                                 fg_color="gray25", hover_color="#444",
                                 command=self.copy_purchase_details)
        btn_copy.grid(row=0, column=0, padx=5)

        btn_email = ctk.CTkButton(btn_grid, text="📧 Open Email Client", 
                                  width=200, height=45, font=ctk.CTkFont(size=13, weight="bold"),
                                  fg_color="#3498db", hover_color="#2980b9",
                                  command=self.send_purchase_email)
        btn_email.grid(row=0, column=1, padx=5)

        # Disclaimer / Terms of Use
        disclaimer_frame = ctk.CTkFrame(self, fg_color="transparent")
        disclaimer_frame.pack(fill="x", pady=(10, 5))
        
        disclaimer_text = (
            "Disclaimer: This software is for professional office use only. "
            "Developer is not liable for any misuse, fraud, or illegal activity "
            "conducted by the user through this software. Provided 'AS IS' without warranty."
        )
        ctk.CTkLabel(disclaimer_frame, text=disclaimer_text, 
                     font=ctk.CTkFont(size=10, slant="italic"), 
                     text_color="gray60", wraplength=500).pack()

    def get_purchase_text(self):
        """Standardized purchase request text."""
        mid = self.status.get('machine_id', 'Unknown')
        return (
            "Hello,\n\n"
            "I would like to purchase a license for the CA Office PDF Utility.\n\n"
            "My Details:\n"
            f"Machine ID: {mid}\n"
            "Name: [Your Name here]\n"
            "Mobile: [Your Mobile here]\n\n"
            "Please provide payment instructions.\n"
            "Thank you."
        )

    def copy_purchase_details(self):
        """Copies the purchase draft to the clipboard."""
        text = self.get_purchase_text()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.app.show_toast("Copied", "Purchase details copied to clipboard!")

    def send_purchase_email(self):
        """Opens the user's default email client with a simplified pre-filled draft."""
        import os
        import urllib.parse
        import webbrowser
        
        email = "support.marutitechsolutions@gmail.com"
        subject = "License Purchase Request"
        # Only use basic fields to avoid browser length limits
        mailto_url = f"mailto:{email}?subject={urllib.parse.quote(subject)}"
        
        # Also auto-copy the body just in case it doesn't auto-fill
        self.copy_purchase_details()
        
        # os.startfile is more native to Windows
        try:
            os.startfile(mailto_url)
        except:
            webbrowser.open(mailto_url)

    def copy_text(self, text, message="Copied!"):
        """Generic helper to copy text to clipboard."""
        self.clipboard_clear()
        self.clipboard_append(text)
        self.app.show_toast("Copied", message)

    def copy_id(self):
        self.clipboard_clear()
        self.clipboard_append(self.status['machine_id'])
        self.app.show_toast("Copied", "Machine ID copied to clipboard!")

    def toggle_ca_fields(self):
        """Shows or hides the membership number field."""
        if self.ca_var.get():
            self.member_frame.pack(pady=10, after=self.ca_frame)
        else:
            self.member_frame.forget()

    def activate(self):
        key = self.ent_key.get()
        member_no = self.ent_member.get().strip() if self.ca_var.get() else None
        
        if not key:
            self.app.show_toast("Error", "Please enter a key.", is_error=True)
            return
            
        if self.ca_var.get() and not member_no:
            self.app.show_toast("Error", "Please enter your CA Membership Number.", is_error=True)
            return
            
        if LicenseManager.activate(key, member_no):
            self.app.show_toast("Success", "Software activated successfully! Thank you.")
            if self.on_success:
                self.on_success()
        else:
            error_msg = "Invalid License Key."
            if member_no:
                error_msg = "Invalid Key or Mismatched Membership Number."
            self.app.show_toast("Failed", error_msg, is_error=True)
