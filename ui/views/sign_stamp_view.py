import customtkinter as ctk
from tkinter import Canvas, filedialog, messagebox
from PIL import Image, ImageTk
import os
from ui.theme import Theme
from core.pdf_editor_engine import PDFEditorEngine
from core.pdf_editor_state import (
    PDFEditorState, RotateCommand, DeletePageCommand, 
    DuplicatePageCommand, AddOverlayCommand, 
    UpdateOverlayCommand, DeleteOverlayCommand,
    AddGroupOverlayCommand, RemoveGroupOverlayCommand
)
from .sign_dialogs import SignStampDialog
from .digital_id_dialog import DigitalIDDialog
from core.digital_signature_engine import DigitalSignatureEngine
from utils.settings_manager import SettingsManager
from ui.components import InstructionDialog
import datetime

class SignStampView(ctk.CTkFrame):
    def __init__(self, master, app_instance):
        super().__init__(master, fg_color=Theme.BG_PRIMARY)
        self.app = app_instance
        self.engine = PDFEditorEngine()
        self.state = PDFEditorState()
        
        self.drag_mode = "move" # or "resize"
        self.mode = "select"
        self.selection_rect = None
        self.drag_start_coords = None
        self.temp_bbox = None
        self._insertion_offset = 0
        self.sig_engine = DigitalSignatureEngine()
        
        self.setup_ui()
        self.update_button_states()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure(0, weight=1)

        # 1. Left Panel (Thumbnails)
        self.left_panel = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color=Theme.BG_SECONDARY)
        self.left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 1))
        
        # Border separation
        ctk.CTkFrame(self, width=1, fg_color=Theme.BORDER_COLOR).grid(row=0, column=0, sticky="nse")
        
        self.open_btn = ctk.CTkButton(self.left_panel, text="📂 Open PDF", height=36, corner_radius=8,
                                      fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                                      font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                      command=self.open_pdf)
        self.open_btn.pack(pady=(20, 10), padx=15, fill="x")
        
        self.save_btn = ctk.CTkButton(self.left_panel, text="💾 Save PDF", height=36, corner_radius=8,
                                      fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                      font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                      command=self.save_pdf)
        self.save_btn.pack(pady=(0, 20), padx=15, fill="x")
        
        self.thumb_scroll = ctk.CTkScrollableFrame(self.left_panel, label_text="Pages", 
                                                  fg_color="transparent", label_fg_color="transparent",
                                                  label_font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"))
        self.thumb_scroll.pack(expand=True, fill="both", padx=5, pady=5)

        # 2. Center Panel (PDF Canvas & Toolbar)
        self.center_panel = ctk.CTkFrame(self, corner_radius=0, fg_color=Theme.BG_PRIMARY)
        self.center_panel.grid(row=0, column=1, sticky="nsew")
        
        self.toolbar_container = ctk.CTkFrame(self.center_panel, height=60, fg_color=Theme.BG_SECONDARY, corner_radius=0)
        self.toolbar_container.pack(side="top", fill="x")
        
        # Bottom border for toolbar
        ctk.CTkFrame(self.toolbar_container, height=1, fg_color=Theme.BORDER_COLOR).pack(side="bottom", fill="x")
        
        self.toolbar = ctk.CTkScrollableFrame(self.toolbar_container, orientation="horizontal", height=50, fg_color="transparent")
        self.toolbar.pack(fill="both", expand=True, padx=10)
        
        self.create_toolbar_groups()

        self.canvas_frame = ctk.CTkFrame(self.center_panel, fg_color=Theme.BG_PRIMARY, corner_radius=0)
        self.canvas_frame.pack(expand=True, fill="both")
        
        self.canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas_frame.grid_columnconfigure(0, weight=1)
        
        self.canvas = Canvas(self.canvas_frame, bg=Theme.BG_PRIMARY, highlightthickness=0, takefocus=True)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        
        # Add Scrollbars
        self.v_scrollbar = ctk.CTkScrollbar(self.canvas_frame, orientation="vertical", command=self.canvas.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.h_scrollbar = ctk.CTkScrollbar(self.canvas_frame, orientation="horizontal", command=self.canvas.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        
        # Bindings for Canvas Interaction
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Button-3>", self.show_context_menu)
        self.canvas.bind("<Configure>", lambda e: self._on_resize())
        
        # Mouse Wheel Bindings
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)
        
        # KEY BINDINGS: Bind to root window to catch keys regardless of focus
        self.after(500, self._bind_keys)
        
        # Show "No PDF" placeholder initially
        self.show_empty_state()
        
    def _on_resize(self):
        """Handle canvas resizing and redraw empty state if needed."""
        if not self.state.doc:
            self.show_empty_state()

    def show_empty_state(self):
        """Displays a clean placeholder message when no document is loaded."""
        self.canvas.delete("all")
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 100: w, h = 800, 600 # Fallback for initial render
        
        self.canvas.create_text(w/2, h/2, 
                                text="No PDF Loaded\n\nDrag & Drop a PDF here or use 'Open PDF'", 
                                fill=Theme.TEXT_MUTED, 
                                font=(Theme.FONT_FAMILY, 14, "bold"),
                                justify="center", tags="placeholder")

    def _bind_keys(self):
        try:
            root = self.winfo_toplevel()
            root.bind("<Delete>", lambda e: self.delete_overlay())
            root.bind("<BackSpace>", lambda e: self.delete_overlay())
        except:
            pass

        # Status Bar
        self.status_bar = ctk.CTkFrame(self.center_panel, height=28, corner_radius=0, fg_color=Theme.BG_SECONDARY)
        self.status_bar.pack(side="bottom", fill="x")
        
        # Top border for status bar
        ctk.CTkFrame(self.status_bar, height=1, fg_color=Theme.BORDER_COLOR).place(relx=0, rely=0, relwidth=1)
        
        self.status_lbl = ctk.CTkLabel(self.status_bar, text="Ready", font=(Theme.FONT_FAMILY, 11), text_color=Theme.TEXT_MUTED)
        self.status_lbl.pack(side="left", padx=15)
        self.page_info_lbl = ctk.CTkLabel(self.status_bar, text="Page 0 of 0", font=(Theme.FONT_FAMILY, 11), text_color=Theme.TEXT_MUTED)
        self.page_info_lbl.pack(side="right", padx=15)
        self.zoom_lbl = ctk.CTkLabel(self.status_bar, text="Zoom 100%", font=(Theme.FONT_FAMILY, 11), text_color=Theme.TEXT_MUTED)
        self.zoom_lbl.pack(side="right", padx=15)

        # 3. Right Panel (Properties)
        self.right_panel = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=Theme.BG_SECONDARY)
        self.right_panel.grid(row=0, column=2, sticky="nsew", padx=(1, 0))
        
        # Border separation
        ctk.CTkFrame(self, width=1, fg_color=Theme.BORDER_COLOR).grid(row=0, column=2, sticky="nsw")
        
        ctk.CTkLabel(self.right_panel, text="Properties", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=16, weight="bold")).pack(pady=20)
        self.prop_content = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.prop_content.pack(expand=True, fill="both", padx=15)

    def create_toolbar_groups(self):
        # Group 1: View
        g_view = self.create_group("View")
        self.zoom_in_btn = self.add_tool(g_view, "➕", "Zoom In", self.zoom_in)
        self.zoom_out_btn = self.add_tool(g_view, "➖", "Zoom Out", self.zoom_out)
        
        # Group 2: Signing
        g_sig = self.create_group("Signing")
        self.add_tool(g_sig, "🖋️ Draw", "Draw Signature", lambda: self.open_sign_stamp_tab("🖋️ Draw"))
        self.add_tool(g_sig, "📂 Upload", "Upload Signature", lambda: self.open_sign_stamp_tab("📂 Upload"))
        self.add_tool(g_sig, "🎯 Digital Sign", "Drag area for Digital Sign", self.set_digital_sign_mode, color="#2ecc71")
        
        # Group 3: Stamps
        g_stamp = self.create_group("Stamps")
        self.add_tool(g_stamp, "🏷️ Std Stamp", "Standard Stamps", lambda: self.open_sign_stamp_tab("🏷️ Standard"))
        self.add_tool(g_stamp, "✨ Custom", "Custom Stamp", lambda: self.open_sign_stamp_tab("✨ Custom"))
        self.add_tool(g_stamp, "💾 Saved", "Saved Items", lambda: self.open_sign_stamp_tab("💾 Saved"))
        self.add_tool(g_stamp, "📅 Date", "Current Date/Time", lambda: self.open_sign_stamp_tab("📅 Date/Time"))

    def create_group(self, name):
        frame = ctk.CTkFrame(self.toolbar, fg_color="transparent")
        frame.pack(side="left", padx=5)
        return frame

    def add_tool(self, parent, text, tooltip, command, width=70, color=None, text_color=None):
        btn = ctk.CTkButton(parent, text=text, width=width, height=32, corner_radius=6,
                            fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=12),
                            command=command)
        if color: btn.configure(fg_color=color, border_width=0)
        if text_color: btn.configure(text_color=text_color)
        btn.pack(side="left", padx=4, pady=4)
        btn.bind("<Enter>", lambda e: self.set_status(tooltip))
        btn.bind("<Leave>", lambda e: self.set_status("Ready"))
        return btn

    def set_digital_sign_mode(self):
        # Show instruction if not disabled
        if SettingsManager.get("show_digisign_hint", True):
            msg = ("Using your mouse, click and drag to draw the area where you would like the signature to appear. "
                   "Once you finish dragging out the desired area, you will be taken to the next step of the signing process.")
            InstructionDialog(self.master, title="Acrobat Reader", message=msg, setting_key="show_digisign_hint")
            
        self.mode = "digital_sign"
        self.canvas.configure(cursor="cross")
        self.set_status("DRAW AREA: Click and drag on PDF to place signature")

    def set_status(self, msg):
        self.status_lbl.configure(text=msg)

    def update_button_states(self):
        doc_open = self.state.doc is not None
        btns = [self.save_btn, self.zoom_in_btn, self.zoom_out_btn]
        state = "normal" if doc_open else "disabled"
        for b in btns: b.configure(state=state)

    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            if self.app.check_operation_allowed():
                self.state.load_doc(path)
                # CRITICAL SYNC: Ensure engine uses the exact same document object as state
                self.engine.doc = self.state.doc
                self.load_pdf_data()
                self.set_status(f"Opened: {os.path.basename(path)}")

    def load_pdf_data(self):
        try:
            for widget in self.thumb_scroll.winfo_children(): widget.destroy()
            thumbs = self.engine.get_thumbnails()
            for i, thumb_img in enumerate(thumbs):
                thumb_ctk = ctk.CTkImage(light_image=thumb_img, dark_image=thumb_img, size=(120, 160))
                btn = ctk.CTkButton(self.thumb_scroll, image=thumb_ctk, text=f"Page {i+1}", 
                                    compound="top", fg_color="transparent", corner_radius=8,
                                    font=(Theme.FONT_FAMILY, 11),
                                    command=lambda p=i: self.select_page(p))
                btn.pack(pady=5, padx=5)
                if i == self.state.current_page_index: btn.configure(fg_color=Theme.ACCENT_BLUE, border_width=0)
            self.render_page()
            self.update_button_states()
        except Exception as e:
            self.app.show_toast("Error", str(e), is_error=True)

    def select_page(self, page_num):
        self.state.current_page_index = page_num
        self.load_pdf_data()

    def render_page(self):
        if not self.state.doc: return
        img = self.engine.get_page_image(self.state.current_page_index, zoom=self.state.zoom_level)
        if img:
            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(10, 10, anchor="nw", image=self.photo, tags="pdf")
            self.canvas.config(scrollregion=(0, 0, img.width + 20, img.height + 20))
            
            # Render session overlays
            self.render_overlays()
            
            # Render selection highlight for PDF objects (Signatures)
            self.render_content_highlights()
            
            self.page_info_lbl.configure(text=f"Page {self.state.current_page_index+1} of {len(self.state.doc)}")
            self.zoom_lbl.configure(text=f"Zoom {int(self.state.zoom_level*100)}%")

    def render_content_highlights(self):
        """Draws a highlight box around the selected original PDF object (e.g. Signature)."""
        self.canvas.delete("content_highlight")
        if self.state.selected_object:
            bbox = self.state.selected_object["bbox"]
            zoom = self.state.zoom_level
            x0 = bbox[0] * zoom + 10
            y0 = bbox[1] * zoom + 10
            x1 = bbox[2] * zoom + 10
            y1 = bbox[3] * zoom + 10
            self.canvas.create_rectangle(x0, y0, x1, y1, outline=Theme.ACCENT_BLUE, width=3, dash=(4, 4), tags="content_highlight")

    def render_overlays(self):
        self.canvas.delete("overlay")
        overlays = self.state.page_overlays.get(self.state.current_page_index, [])
        zoom = self.state.zoom_level
        
        for i, ov in enumerate(overlays):
            bbox = ov["bbox"]
            x0 = bbox[0] * zoom + 10
            y0 = bbox[1] * zoom + 10
            x1 = bbox[2] * zoom + 10
            y1 = bbox[3] * zoom + 10
            tags = ("overlay", f"ov_{i}")
            
            if ov["type"] == "image":
                pil_img = Image.open(ov["path"]).convert("RGBA")
                if ov.get("rotation"):
                    pil_img = pil_img.rotate(-ov["rotation"], expand=True, resample=Image.Resampling.BICUBIC)
                if ov.get("opacity", 1.0) < 1.0:
                    alpha = pil_img.split()[3].point(lambda p: p * ov["opacity"])
                    pil_img.putalpha(alpha)
                
                target_w = int(x1 - x0)
                target_h = int(y1 - y0)
                
                # Aspect Ratio Logic
                if ov.get("preserve_aspect") and "aspect" in ov:
                    orig_aspect = ov["aspect"]
                    box_aspect = target_w / target_h if target_h != 0 else 1.0
                    
                    if orig_aspect > box_aspect:
                        # Image is wider than box relatively
                        new_w = target_w
                        new_h = int(target_w / orig_aspect)
                        offset_y = (target_h - new_h) // 2
                        x0_final, y0_final = x0, y0 + offset_y
                    else:
                        # Image is taller than box relatively
                        new_h = target_h
                        new_w = int(target_h * orig_aspect)
                        offset_x = (target_w - new_w) // 2
                        x0_final, y0_final = x0 + offset_x, y0
                else:
                    new_w, new_h = target_w, target_h
                    x0_final, y0_final = x0, y0

                if new_w < 1: new_w = 1
                if new_h < 1: new_h = 1
                
                pil_img = pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(pil_img)
                if not hasattr(self, 'overlay_images'): self.overlay_images = {}
                self.overlay_images[f"ov_{i}"] = photo
                self.canvas.create_image(x0_final, y0_final, anchor="nw", image=photo, tags=tags)
                
            elif ov["type"] == "text":
                # AUTO-FIT SCALING: Derive font size directly from current box height
                box_h_pts = max(5, ov["bbox"][3] - ov["bbox"][1])
                line_count = ov["text"].count("\n") + 1
                if "\n\n\n" in ov["text"]: line_count = 4 # Handle custom stamp signature space
                
                # Professional ratio: fill ~70% of line height
                calc_fontsize = int(box_h_pts / (line_count * 1.35))
                calc_fontsize = max(6, calc_fontsize)
                
                if ov.get("rotation") or ov.get("opacity", 1.0) < 1.0:
                    from PIL import ImageDraw, ImageFont
                    tw, th = int(x1-x0), int(y1-y0)
                    t_img = Image.new("RGBA", (max(1, tw), max(1, th)), (0,0,0,0))
                    draw = ImageDraw.Draw(t_img)
                    try: font = ImageFont.truetype("arial.ttf", int(calc_fontsize * zoom))
                    except: font = ImageFont.load_default()
                    
                    curr_y = 5
                    for line in ov["text"].split("\n"):
                        try:
                            l, t, r, b = draw.textbbox((0,0), line, font=font)
                            draw.text(((tw-(r-l))/2, curr_y), line, fill=ov.get("color", "red"), font=font)
                            curr_y += (b-t) + 5
                        except: pass
                    
                    if ov.get("rotation"): t_img = t_img.rotate(-ov["rotation"], expand=True)
                    if ov.get("opacity", 1.0) < 1.0:
                        alpha = t_img.split()[3].point(lambda p: p * ov["opacity"])
                        t_img.putalpha(alpha)
                    
                    photo = ImageTk.PhotoImage(t_img)
                    if not hasattr(self, 'overlay_images'): self.overlay_images = {}
                    self.overlay_images[f"ov_{i}"] = photo
                    self.canvas.create_image((x0+x1)/2, (y0+y1)/2, anchor="center", image=photo, tags=tags)
                else:
                    self.canvas.create_text((x0+x1)/2, (y0+y1)/2, text=ov["text"], 
                                            fill=ov.get("color", "red"), font=("Arial", int(calc_fontsize * zoom), "bold" if ov.get("bold") else "normal"),
                                            justify=ov.get("align_str", "center"), anchor="center", tags=tags)
            
            self.canvas.create_rectangle(x0, y0, x1, y1, fill="", outline="", tags=(tags, "hitbox"))
            if self.state.selected_overlay == ov:
                self.canvas.create_rectangle(x0, y0, x1, y1, outline=Theme.ACCENT_BLUE, width=1, dash=(4,2), tags=tags)
                self.canvas.create_rectangle(x1-6, y1-6, x1, y1, fill=Theme.ACCENT_BLUE, tags=tags)

    def on_canvas_click(self, event):
        self.canvas.focus_set()
        x = (self.canvas.canvasx(event.x) - 10) / self.state.zoom_level
        y = (self.canvas.canvasy(event.y) - 10) / self.state.zoom_level
        
        if self.mode == "digital_sign":
            self.drag_start_coords = (x, y)
            if self.selection_rect: self.canvas.delete(self.selection_rect)
            self.selection_rect = self.canvas.create_rectangle(
                event.x, event.y, event.x, event.y, 
                outline=Theme.ACCENT_BLUE, width=2, dash=(4, 4), tags="temp_selection"
            )
            return
        
        # 1. Check overlays first
        self.state.selected_overlay = None
        overlays = self.state.page_overlays.get(self.state.current_page_index, [])
        for ov in reversed(overlays):
            b = ov["bbox"]
            if b[0] <= x <= b[2] and b[1] <= y <= b[3]:
                self.state.selected_overlay = ov
                self.is_dragging = True
                self.last_mouse_x, self.last_mouse_y = x, y
                if x > b[2] - 20 and y > b[3] - 20: self.drag_mode = "resize"
                else: self.drag_mode = "move"
                self.render_overlays()
                self.show_overlay_properties(ov)
                self.state.selected_object = None # Clear object selection if overlay clicked
                self.render_content_highlights()
                return
        
        # 2. Check for PDF objects (Digital Signatures)
        self.state.selected_object = None
        objs = self.engine.get_page_objects(self.state.current_page_index)
        for obj in reversed(objs):
            if obj.get("type") == "signature":
                b = obj["bbox"]
                if b[0] <= x <= b[2] and b[1] <= y <= b[3]:
                    self.state.selected_object = obj
                    self.show_overlay_properties(obj) # Reuse prop panel for deletion
                    self.render_content_highlights()
                    self.render_overlays() # Clear overlay selection
                    return

        self.clear_properties()
        self.render_overlays()
        self.render_content_highlights()

    def on_canvas_drag(self, event):
        x = (self.canvas.canvasx(event.x) - 10) / self.state.zoom_level
        y = (self.canvas.canvasy(event.y) - 10) / self.state.zoom_level

        if self.mode == "digital_sign" and self.drag_start_coords:
            sx, sy = self.drag_start_coords
            zoom = self.state.zoom_level
            self.canvas.coords(self.selection_rect, sx*zoom+10, sy*zoom+10, event.x, event.y)
            return

        if not self.is_dragging or not self.state.selected_overlay: return
        dx, dy = x - self.last_mouse_x, y - self.last_mouse_y
        ov = self.state.selected_overlay
        if self.drag_mode == "move":
            ov["bbox"] = [ov["bbox"][0]+dx, ov["bbox"][1]+dy, ov["bbox"][2]+dx, ov["bbox"][3]+dy]
        else:
            ov["bbox"][2] = max(ov["bbox"][0]+20, x)
            ov["bbox"][3] = max(ov["bbox"][1]+20, y)
        self.last_mouse_x, self.last_mouse_y = x, y
        self.render_overlays()
        self.show_overlay_properties(ov)

    def on_canvas_release(self, event):
        if self.mode == "digital_sign" and self.drag_start_coords:
            x = (self.canvas.canvasx(event.x) - 10) / self.state.zoom_level
            y = (self.canvas.canvasy(event.y) - 10) / self.state.zoom_level
            sx, sy = self.drag_start_coords
            
            # Finalize selection bbox in PDF coords
            x0, x1 = min(sx, x), max(sx, x)
            y0, y1 = min(sy, y), max(sy, y)
            
            # Minimum area check (e.g. 5x5 pts)
            if abs(x1 - x0) > 5 and abs(y1 - y0) > 5:
                self.temp_bbox = [x0, y0, x1, y1]
                # Check if we should do DSC or standard Sign/Stamp
                # For now, "Digital Sign" mode triggers DSC
                self.open_digital_id_selection()
            else:
                self.canvas.delete("temp_selection")
                self.mode = "select"
                self.canvas.configure(cursor="")
                
            self.drag_start_coords = None
            return

        self.is_dragging = False

    def open_sign_stamp_tab(self, tab_name):
        SignStampDialog(self.master, on_apply=self.handle_sign_stamp_apply, start_tab=tab_name)

    def open_digital_id_selection(self):
        DigitalIDDialog(self.master, on_select=self.handle_digital_id_apply)

    def handle_digital_id_apply(self, cert_data):
        if not self.temp_bbox: return
        
        target_path = self.state.current_path
        if not target_path:
            self.set_status("ERROR: No PDF loaded to sign")
            return
            
        # Update status immediately
        self.set_status("ACCESSING TOKEN: Please check for Windows PIN/Password popup...")
        # self.app.show_toast("Digital Sign", "Please enter your token PIN in the Windows popup") # REMOVED: Blocks thread
        self.update() # Force UI refresh
        
        # Close the document to release file lock before signing
        bbox = self.temp_bbox
        page_idx = self.state.current_page_index
        if self.state.doc:
            self.state.doc.close()
            self.state.doc = None
        
        # Determine output path: original_signed.pdf
        base, ext = os.path.splitext(target_path)
        output_path = f"{base}_signed{ext}"
        
        try:
            success = self.sig_engine.sign_pdf(
                target_path, 
                output_path, 
                cert_data,
                bbox,
                page_idx
            )
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.app.show_toast("Error", str(e), is_error=True)
            self.set_status(f"ERROR: {str(e)}")
            success = False
        
        # Always reload the document (whether success or failure)
        # Load the SIGNED file if successful, otherwise original
        load_path = output_path if success else target_path
        self.state.load_doc(load_path)
        
        # CRITICAL FIX: Synchronize engine doc with the new document object
        self.engine.doc = self.state.doc
        self.load_pdf_data()
        
        if success:
            self.set_status(f"SUCCESS: PDF Digitally Signed with {cert_data['name']}")
            self.app.show_toast("Digital Sign", "Digital Signature applied successfully.")
        else:
            self.set_status("ERROR: Digital Signing failed. Check token connection.")

        # Cleanup
        self.mode = "select"
        self.canvas.configure(cursor="")
        self.canvas.delete("temp_selection")
        self.temp_bbox = None

    def open_sign_stamp_module(self):
        SignStampDialog(self.master, on_apply=self.handle_sign_stamp_apply)

    def handle_sign_stamp_apply(self, data):
        if not hasattr(self, '_insertion_offset'): self._insertion_offset = 0
        
        # Determine position and size
        if self.temp_bbox:
            x0, y0, x1, y1 = self.temp_bbox
            w, h = x1 - x0, y1 - y0
            # Reset mode after use
            self.mode = "select"
            self.canvas.configure(cursor="")
            self.canvas.delete("temp_selection")
            self.temp_bbox = None
        else:
            self._insertion_offset = (self._insertion_offset + 30) % 300
            x0, y0 = 100 + self._insertion_offset, 100 + self._insertion_offset
            w, h = data.get("width", 150), data.get("height", 80 if data["type"]=="text" else 50)
            x1, y1 = x0 + w, y0 + h

        ov = {
            "type": data["type"], 
            "bbox": [x0, y0, x1, y1], 
            "rotation": 0, 
            "opacity": 1.0,
            "preserve_aspect": True if data["type"] == "image" else False
        }
        
        if data["type"] == "image": 
            ov["path"] = data["path"]
            # Store original aspect ratio
            try:
                with Image.open(data["path"]) as img:
                    ov["aspect"] = img.width / img.height
            except:
                ov["aspect"] = w / h if h != 0 else 1.0
        else: 
            ov.update({
                "text": data["text"], 
                "color": data["color"], 
                "fontsize": data["fontsize"], 
                "bold": data.get("bold", False), 
                "align": data.get("align", 1),
                "align_str": data.get("align_str", "center")
            })
            
        self.state.push_command(AddOverlayCommand(self.state, self.state.current_page_index, ov))
        self.render_overlays()
        self.canvas.update()

    def show_overlay_properties(self, ov):
        try:
            for widget in self.prop_content.winfo_children(): widget.destroy()
            ctk.CTkLabel(self.prop_content, text="✒️ Object Properties", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=14, weight="bold")).pack(pady=10)
            
            # Metadata Frame
            f = ctk.CTkFrame(self.prop_content, fg_color=Theme.BG_PRIMARY, corner_radius=Theme.CORNER_RADIUS, border_width=1, border_color=Theme.BORDER_COLOR)
            f.pack(fill="x", padx=5, pady=5)
            
            if ov["type"] == "image": ov_type = "Digital Signature (Image)"
            elif ov["type"] == "signature": ov_type = "Digital Signature (Cryptographic)"
            else: ov_type = "Stamp"
            
            ctk.CTkLabel(f, text=f"Type: {ov_type}", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=11, weight="bold")).pack(anchor="w", padx=10, pady=(10, 2))
            ctk.CTkLabel(f, text=f"Page: {self.state.current_page_index + 1}", font=(Theme.FONT_FAMILY, 10), text_color=Theme.TEXT_MUTED).pack(anchor="w", padx=10)
            
            bbox = ov["bbox"]
            pos_text = f"X: {int(bbox[0])} Y: {int(bbox[1])} | W: {int(bbox[2]-bbox[0])} H: {int(bbox[3]-bbox[1])}"
            ctk.CTkLabel(f, text=pos_text, font=(Theme.FONT_FAMILY, 10), text_color=Theme.TEXT_MUTED).pack(anchor="w", padx=10, pady=(0, 10))

            # Controls
            if ov["type"] != "signature":
                ctk.CTkLabel(self.prop_content, text="Rotation", font=(Theme.FONT_FAMILY, 12)).pack(pady=(12,0))
                rot = ctk.CTkSlider(self.prop_content, from_=0, to=360, button_color=Theme.ACCENT_BLUE, command=lambda v: self.update_ov_prop(ov, "rotation", int(v)))
                rot.set(ov.get("rotation", 0)); rot.pack(pady=5)
                
                ctk.CTkLabel(self.prop_content, text="Opacity", font=(Theme.FONT_FAMILY, 12)).pack(pady=(12,0))
                opac = ctk.CTkSlider(self.prop_content, from_=0.1, to=1.0, button_color=Theme.ACCENT_BLUE, command=lambda v: self.update_ov_prop(ov, "opacity", float(v)))
                opac.set(ov.get("opacity", 1.0)); opac.pack(pady=5)

            if ov["type"] == "text":
                ctk.CTkLabel(self.prop_content, text="Edit Text", font=(Theme.FONT_FAMILY, 12)).pack(pady=(12,0))
                txt = ctk.CTkTextbox(self.prop_content, height=60, font=(Theme.FONT_FAMILY, 12), border_width=1, border_color=Theme.BORDER_COLOR)
                txt.insert("1.0", ov["text"]); txt.pack(fill="x", pady=5, padx=5)
                ctk.CTkButton(self.prop_content, text="Update Text", height=32, corner_radius=6, 
                              fg_color="transparent", border_width=1, border_color=Theme.ACCENT_BLUE, 
                              font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                              command=lambda: self.update_ov_prop(ov, "text", txt.get("1.0", "end-1c"))).pack(pady=5, fill="x")
                
                # Manual Resize
                ctk.CTkLabel(self.prop_content, text="Stamp Width", font=(Theme.FONT_FAMILY, 12)).pack(pady=(12,0))
                ws = ctk.CTkSlider(self.prop_content, from_=50, to=800, button_color=Theme.ACCENT_BLUE, command=lambda v: self.update_ov_prop(ov, "width", float(v)))
                ws.set(ov["bbox"][2] - ov["bbox"][0]); ws.pack(pady=2)
                
                ctk.CTkLabel(self.prop_content, text="Stamp Height", font=(Theme.FONT_FAMILY, 12)).pack(pady=(10,0))
                hs = ctk.CTkSlider(self.prop_content, from_=20, to=500, button_color=Theme.ACCENT_BLUE, command=lambda v: self.update_ov_prop(ov, "height", float(v)))
                hs.set(ov["bbox"][3] - ov["bbox"][1]); hs.pack(pady=2)

            # Actions
            action_f = ctk.CTkFrame(self.prop_content, fg_color="transparent")
            action_f.pack(fill="x", pady=(15, 5))
            
            if ov["type"] != "signature":
                ctk.CTkButton(action_f, text="📋 Duplicate", width=95, height=32, corner_radius=6, 
                              fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                              command=self.duplicate_overlay).pack(side="left", padx=2)
            
            ctk.CTkButton(action_f, text="🗑 Delete", width=95 if ov["type"]!="signature" else 200, height=32, corner_radius=6,
                          fg_color="#e74c3c", hover_color="#c0392b", command=self.delete_overlay).pack(side="left", padx=2)
            
            if ov["type"] != "signature":
                group_id = ov.get("group_id")
                if not group_id:
                    ctk.CTkButton(self.prop_content, text="📋 Apply to All Pages", height=32, corner_radius=6,
                                  fg_color="#27ae60", hover_color="#219150",
                                  font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                  command=lambda: self.apply_to_all_pages(ov)).pack(pady=5, fill="x", padx=5)
                else:
                    ctk.CTkButton(self.prop_content, text="🗑 Remove from All Pages", height=32, corner_radius=6,
                                  fg_color="#e67e22", hover_color="#d35400",
                                  font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                  command=lambda: self.remove_from_all_pages(group_id)).pack(pady=5, fill="x", padx=5)

                ctk.CTkButton(self.prop_content, text="Clear All on Page", height=32, corner_radius=6,
                               fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR, 
                               text_color=Theme.TEXT_MUTED, command=self.clear_all_page_overlays).pack(pady=15, fill="x", padx=5)
            
            ctk.CTkButton(self.prop_content, text="Deselect", height=32, corner_radius=6, fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR, text_color=Theme.TEXT_MUTED, command=self.clear_properties).pack(pady=5, fill="x", padx=5)

        except Exception as e:
            self.app.show_toast("Error", f"Failed to show props: {str(e)}", is_error=True)

    def update_ov_prop(self, ov, prop, val):
        if prop == "width":
            ov["bbox"][2] = ov["bbox"][0] + val
        elif prop == "height":
            ov["bbox"][3] = ov["bbox"][1] + val
        else:
            ov[prop] = val
        self.render_overlays()

    def delete_overlay(self):
        # Case 1: Session Overlay (Stamp/drawn sign)
        if self.state.selected_overlay:
            self.state.push_command(DeleteOverlayCommand(self.state, self.state.current_page_index, self.state.selected_overlay))
            self.state.selected_overlay = None
            self.render_page(); self.clear_properties()
            return
            
        # Case 2: Native PDF Object (Digital Signature)
        if self.state.selected_object and self.state.selected_object.get("type") == "signature":
            if self.engine.delete_object(self.state.current_page_index, self.state.selected_object):
                # PERSIST CHANGE TO DISK: Crucial to prevent 'reappearing' signatures
                # This call re-opens the document object in the engine on Windows.
                self.engine.save_pdf(self.state.current_path)
                
                # SYNC state with engine's new doc object
                self.state.doc = self.engine.doc
                
                self.app.show_toast("Digital Sign", "Digital Signature removed successfully.")
                self.state.selected_object = None
                self.load_pdf_data() # Reload to show changes (re-renders page)
                self.clear_properties()

    def duplicate_overlay(self):
        if self.state.selected_overlay:
            new_ov = self.state.selected_overlay.copy()
            new_ov["bbox"] = [b + 10 for b in new_ov["bbox"]]
            self.state.push_command(AddOverlayCommand(self.state, self.state.current_page_index, new_ov))
            self.render_page()

    def clear_all_page_overlays(self):
        if self.state.current_page_index in self.state.page_overlays:
            self.state.page_overlays[self.state.current_page_index] = []
            self.render_page(); self.clear_properties()

    def apply_to_all_pages(self, ov):
        if not self.state.doc: return
        target_pages = list(range(len(self.state.doc)))
        self.state.push_command(AddGroupOverlayCommand(self.state, self.state.current_page_index, ov, target_pages))
        self.render_overlays()
        self.show_overlay_properties(ov) # Refresh properties to show "Remove" button
        self.app.show_toast("Success", "Stamp applied to all pages.")

    def remove_from_all_pages(self, group_id):
        self.state.push_command(RemoveGroupOverlayCommand(self.state, group_id))
        self.render_page() 
        self.clear_properties()
        self.app.show_toast("Success", "Stamps removed from all pages.")

    def show_context_menu(self, event):
        self.on_canvas_click(event)
        if self.state.selected_overlay:
            from tkinter import Menu
            m = Menu(self, tearoff=0)
            m.add_command(label="Duplicate", command=self.duplicate_overlay)
            m.add_command(label="Delete", command=self.delete_overlay)
            m.tk_popup(event.x_root, event.y_root)

    def save_pdf(self, path=None):
        if not self.state.doc: return
        from tkinter import filedialog
        path = path or filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="signed_doc.pdf")
        if path:
            if self.engine.save_pdf(path, page_overlays=self.state.page_overlays):
                self.app.show_toast("Saved", f"File saved to {path}")

    def zoom_in(self): self.state.zoom_level += 0.2; self.render_page()
    def zoom_out(self):
        if self.state.zoom_level > 0.4: self.state.zoom_level -= 0.2; self.render_page()
    def undo_action(self):
        if self.state.undo(): self.load_pdf_data()
    def redo_action(self):
        if self.state.redo(): self.load_pdf_data()
    def rotate_current_page(self):
        self.state.push_command(RotateCommand(self.state.doc, self.state.current_page_index))
        self.render_page()
    def clear_properties(self):
        for widget in self.prop_content.winfo_children(): widget.destroy()
        ctk.CTkLabel(self.prop_content, text="No object selected", font=(Theme.FONT_FAMILY, 11), text_color=Theme.TEXT_MUTED).pack(pady=20)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")
