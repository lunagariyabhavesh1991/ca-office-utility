import customtkinter as ctk
from customtkinter import filedialog
import tkinter as tk
from tkinterdnd2 import DND_FILES
import os
from ui.theme import Theme

class DragDropArea(ctk.CTkFrame):
    """A visual area that accepts files dragged and dropped onto it."""
    def __init__(self, master, on_drop_callback, title="Drag & Drop Files Here", height=150, **kwargs):
        super().__init__(master, height=height, **kwargs)
        self.on_drop_callback = on_drop_callback
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Configure distinctly colored frame
        self.configure(fg_color=Theme.BG_SECONDARY, border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR, corner_radius=Theme.CORNER_RADIUS)
        # Prevent auto-shrinking
        self.pack_propagate(False)
        self.grid_propagate(False)
        
        self.label = ctk.CTkLabel(self, text=title, font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=14, weight="bold"))
        self.label.grid(row=0, column=0, pady=(20, 5), padx=20, sticky="s")
        
        self.btn_browse = ctk.CTkButton(self, text="Select Files", width=120, height=32, corner_radius=8,
                                        fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER, 
                                        font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                        command=self._browse_files)
        self.btn_browse.grid(row=1, column=0, pady=(5, 20), padx=20, sticky="n")
        
        # Register for drag and drop
        # Note: The main window MUST be a TkinterDnD.Tk or equivalent
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._handle_drop)

    def _browse_files(self):
        files = filedialog.askopenfilenames(title="Select Files")
        if files and self.on_drop_callback:
            self.on_drop_callback(list(files))

    def _handle_drop(self, event):
        files = self._parse_drop_files(event.data)
        if self.on_drop_callback:
            self.on_drop_callback(files)

    def _parse_drop_files(self, data: str) -> list:
        """Parses the dropped data string into a list of file paths. Handles spaces in paths."""
        import re
        # tkDND wraps paths with spaces in curly braces
        braced_paths = re.findall(r'\{([^}]+)\}', data)
        for p in braced_paths:
            data = data.replace('{' + p + '}', '')
            
        other_paths = [p for p in data.split(' ') if p.strip()]
        return braced_paths + other_paths

class FileListFrame(ctk.CTkScrollableFrame):
    """A list frame to show added files. Allows deletion."""
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.files = []
        self.rows = []
        
        # Placeholder for empty list
        self.placeholder = ctk.CTkLabel(self, text="No files added yet", 
                                       text_color=Theme.TEXT_MUTED,
                                       font=(Theme.FONT_FAMILY, 12, "italic"))
        self.placeholder.pack(expand=True, pady=40)

    def add_files(self, new_files):
        """Add new files to the list without duplicates."""
        if new_files and self.placeholder:
            self.placeholder.pack_forget()
            
        for f in new_files:
            if f not in self.files:
                self.files.append(f)
                self._add_row(f)

    def clear_files(self):
        """Remove all files."""
        for row in self.rows:
            row.destroy()
        self.rows.clear()
        self.files.clear()
        
        if self.placeholder:
            self.placeholder.pack(expand=True, pady=40)

    def get_files(self):
        """Returns the list of current files."""
        return self.files

    def _add_row(self, file_path):
        """Adds a UI row for a file."""
        row_frame = ctk.CTkFrame(self, fg_color=Theme.BG_PRIMARY, corner_radius=6)
        row_frame.pack(fill="x", pady=2, padx=5)
        
        filename = os.path.basename(file_path)
        lbl = ctk.CTkLabel(row_frame, text=filename, anchor="w", font=(Theme.FONT_FAMILY, 12))
        lbl.pack(side="left", fill="x", expand=True, padx=(10, 10))
        
        btn_del = ctk.CTkButton(row_frame, text="✕", width=28, height=28, corner_radius=6,
                                fg_color="transparent", hover_color="#e74c3c", text_color=Theme.TEXT_MUTED,
                                command=lambda f=file_path, r=row_frame: self._remove_file(f, r))
        btn_del.pack(side="right", padx=2)
        self.rows.append(row_frame)

    def _remove_file(self, file_path, row_frame):
        """Removes a file from the list and destroys its UI row."""
        self.files.remove(file_path)
        self.rows.remove(row_frame)
        row_frame.destroy()

class SmartNamingFrame(ctk.CTkFrame):
    """Provides input fields for Output Location and File Name."""
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS, 
                         border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR, **kwargs)
        
        lbl_title = ctk.CTkLabel(self, text="Output Settings", font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold", size=14))
        lbl_title.grid(row=0, column=0, columnspan=2, pady=(15, 10), padx=15, sticky="w")
        
        # Configurations
        self.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self, text="Output Folder:", font=(Theme.FONT_FAMILY, 12)).grid(row=1, column=0, padx=15, pady=5, sticky="e")
        
        path_frame = ctk.CTkFrame(self, fg_color="transparent")
        path_frame.grid(row=1, column=1, padx=(0, 15), pady=5, sticky="ew")
        path_frame.grid_columnconfigure(0, weight=1)
        
        self.ent_workspace = ctk.CTkEntry(self, placeholder_text="e.g., C:/CA_Output", height=32, corner_radius=8,
                                         border_color=Theme.BORDER_COLOR, fg_color=Theme.BG_PRIMARY)
        self.ent_workspace.grid(row=1, column=1, padx=(0, 80), pady=5, sticky="ew")
        
        self.btn_browse_dir = ctk.CTkButton(self, text="Browse", width=70, height=32, corner_radius=8,
                                           fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                           font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                           command=self._browse_directory)
        self.btn_browse_dir.grid(row=1, column=1, padx=(0, 10), sticky="e")
        
        # Try to load a generic default if none exists
        default_ws = os.path.join(os.path.expanduser("~"), "Documents")
        self.ent_workspace.insert(0, default_ws)
        
        ctk.CTkLabel(self, text="Output File Name:", font=(Theme.FONT_FAMILY, 12)).grid(row=2, column=0, padx=15, pady=(5, 15), sticky="e")
        self.ent_client = ctk.CTkEntry(self, placeholder_text="e.g., Output.pdf", height=32, corner_radius=8,
                                       border_color=Theme.BORDER_COLOR, fg_color=Theme.BG_PRIMARY)
        self.ent_client.grid(row=2, column=1, padx=(0, 15), pady=(5, 15), sticky="ew")

    def _browse_directory(self):
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.ent_workspace.delete(0, "end")
            self.ent_workspace.insert(0, dir_path)

    def get_data(self) -> dict:
        """Returns the inputs as a dictionary."""
        return {
            "output_dir": self.ent_workspace.get().strip(),
            "output_filename": self.ent_client.get().strip(),
        }

class InstructionDialog(ctk.CTkToplevel):
    """A professional modal dialog with a 'Do not show again' checkbox."""
    def __init__(self, master, title, message, setting_key=None):
        super().__init__(master)
        self.title(title)
        self.geometry("500x320")
        self.setting_key = setting_key
        self.configure(fg_color=Theme.BG_PRIMARY)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Content frame
        content = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS,
                               border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR)
        content.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        # Icon and Message
        msg_frame = ctk.CTkFrame(content, fg_color="transparent")
        msg_frame.pack(fill="x", pady=(20, 20), padx=20)
        
        icon_lbl = ctk.CTkLabel(msg_frame, text="ℹ️", font=("", 40))
        icon_lbl.pack(side="left", padx=(0, 20))
        
        msg_lbl = ctk.CTkLabel(msg_frame, text=message, justify="left", wraplength=350, 
                               font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=13))
        msg_lbl.pack(side="left", fill="both", expand=True)
        
        # Checkbox
        self.check_var = ctk.BooleanVar(value=False)
        self.cb = ctk.CTkCheckBox(content, text="Do not show this message again", variable=self.check_var,
                                  font=(Theme.FONT_FAMILY, 12), text_color=Theme.TEXT_MUTED)
        self.cb.pack(pady=10, padx=20, anchor="w")
        
        # OK Button
        btn = ctk.CTkButton(content, text="OK", width=140, height=36, corner_radius=8,
                            fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                            font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                            command=self._on_ok)
        btn.pack(pady=(20, 20))
        
        # Modal logic
        self.grab_set()
        self.focus_force()
        self.protocol("WM_DELETE_WINDOW", self._on_ok)

    def _on_ok(self):
        if self.setting_key and self.check_var.get():
            from utils.settings_manager import SettingsManager
            SettingsManager.set(self.setting_key, False)
        self.destroy()

class NavButton(ctk.CTkFrame):
    """
    A custom navigation button with a strict two-column layout:
    [Fixed Icon Column] [Flexible Text Column]
    """
    def __init__(self, master, text, icon, command, is_activation=False, is_premium=False, **kwargs):
        super().__init__(master, width=220, height=42, corner_radius=8, fg_color="transparent", **kwargs)
        if hasattr(self, "configure"):
            try: self.configure(cursor="hand2")
            except: pass
        self.command = command
        self.is_activation = is_activation
        self.is_premium = is_premium
        self.is_active = False
        
        # Define Colors
        if is_activation:
            self.bg_color = Theme.ACTIVATION_BG
            self.hover_color = Theme.ACTIVATION_HOVER
            self.active_color = Theme.ACTIVATION_BG
            self.text_color = "white"
        elif is_premium:
            self.bg_color = Theme.ACCENT_AMBER
            self.hover_color = "#D97706" # Slightly darker amber for hover
            self.active_color = Theme.ACCENT_AMBER
            self.text_color = "white"
        else:
            self.bg_color = "transparent"
            self.hover_color = Theme.SIDEBAR_HOVER
            self.active_color = Theme.SIDEBAR_ACTIVE
            self.text_color = Theme.TEXT_PRIMARY
        
        if is_premium:
            self.configure(border_width=0)

        self.configure(fg_color=self.bg_color)
        
        # Grid Configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        # Enable propagation so it fills requested width/height
        self.grid_propagate(True) 
        
        # 1. Icon Label (Fixed Width Column)
        self.icon_label = ctk.CTkLabel(self, text=icon, width=40, height=42, 
                                       font=ctk.CTkFont(size=18), text_color=self.text_color)
        self.icon_label.grid(row=0, column=0, sticky="nsw", padx=(4, 0))
        
        # 2. Text Label (Flexible Column)
        self.text_label = ctk.CTkLabel(self, text=text, height=42, anchor="w",
                                       font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=14, weight="bold" if is_premium else "normal"),
                                       text_color=self.text_color)
        self.text_label.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        
        # Bind events to all child widgets for consistent behavior
        for widget in [self, self.icon_label, self.text_label]:
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.bind("<Button-1>", self._on_click)

    def _on_enter(self, event):
        if not self.is_active:
            self.configure(fg_color=self.hover_color)

    def _on_leave(self, event):
        if not self.is_active:
            self.configure(fg_color=self.bg_color)
        else:
            self.configure(fg_color=self.active_color)

    def _on_click(self, event):
        if self.command:
            self.command()

    def set_active(self, is_active):
        """Sets the active state and updates visual highlight."""
        self.is_active = is_active
        if is_active:
            self.configure(fg_color=self.active_color)
        else:
            self.configure(fg_color=self.bg_color)
