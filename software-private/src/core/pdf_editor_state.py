import fitz
import os

class ActionCommand:
    def execute(self):
        pass
    def undo(self):
        pass
    def redo(self):
        self.execute()

class RotateCommand(ActionCommand):
    def __init__(self, doc, page_num, angle=90):
        self.doc = doc
        self.page_num = page_num
        self.angle = angle
        self.old_rotation = doc[page_num].rotation

    def execute(self, state_ref=None):
        page = self.doc[self.page_num]
        page.set_rotation((page.rotation + self.angle) % 360)

    def redo(self, state_ref=None):
        self.execute()

    def undo(self, state_ref=None):
        self.doc[self.page_num].set_rotation(self.old_rotation)

class DeletePageCommand(ActionCommand):
    def __init__(self, doc, page_num):
        self.doc = doc
        self.page_num = page_num
        # Create a tiny doc to store the deleted page
        self.temp_doc = fitz.open()
        self.temp_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

    def execute(self, state_ref=None):
        self.doc.delete_page(self.page_num)
        # BUG FIX: Shift overlays for subsequent pages
        if state_ref and state_ref.page_overlays:
            new_overlays = {}
            for p_idx, ovs in state_ref.page_overlays.items():
                if p_idx < self.page_num:
                    new_overlays[p_idx] = ovs
                elif p_idx > self.page_num:
                    new_overlays[p_idx - 1] = ovs
            state_ref.page_overlays = new_overlays

    def undo(self, state_ref=None):
        # Insert back at same position
        self.doc.insert_pdf(self.temp_doc, from_page=0, to_page=0, start_at=self.page_num)
        # Shift overlays back
        if state_ref and state_ref.page_overlays:
            new_overlays = {}
            # Sort keys descending to avoid collisions if we were mutating, but we're creating new
            for p_idx in sorted(state_ref.page_overlays.keys(), reverse=True):
                if p_idx >= self.page_num:
                    new_overlays[p_idx + 1] = state_ref.page_overlays[p_idx]
                else:
                    new_overlays[p_idx] = state_ref.page_overlays[p_idx]
            state_ref.page_overlays = new_overlays

class DuplicatePageCommand(ActionCommand):
    def __init__(self, doc, page_num):
        self.doc = doc
        self.page_num = page_num

    def execute(self, state_ref=None):
        # Insert duplicate at the SAME position (making it the 'new' current page)
        # This shifts the original page forward by 1
        self.doc.fullcopy_page(self.page_num, self.page_num)
        # Shift overlays for subsequent pages
        if state_ref and state_ref.page_overlays:
            new_overlays = {}
            # Anything after page_num needs to shift by 1
            for p_idx in sorted(state_ref.page_overlays.keys(), reverse=True):
                if p_idx > self.page_num:
                    new_overlays[p_idx + 1] = state_ref.page_overlays[p_idx]
                else:
                    new_overlays[p_idx] = state_ref.page_overlays[p_idx]
            state_ref.page_overlays = new_overlays

    def undo(self, state_ref=None):
        # Delete the copy at the original position (0) to restore the shifted original to (0)
        self.doc.delete_page(self.page_num)
        # Shift back
        if state_ref and state_ref.page_overlays:
            new_overlays = {}
            for p_idx, ovs in state_ref.page_overlays.items():
                if p_idx > self.page_num + 1:
                    new_overlays[p_idx - 1] = ovs
                elif p_idx <= self.page_num:
                    new_overlays[p_idx] = ovs
            state_ref.page_overlays = new_overlays

class AnnotationCommand(ActionCommand):
    def __init__(self, doc, page_num, bbox, annot_type):
        self.doc = doc
        self.page_num = page_num
        self.bbox = bbox
        self.annot_type = annot_type
        self.annot_xref = None

    def execute(self):
        page = self.doc[self.page_num]
        if self.annot_type == "highlight":
            annot = page.add_highlight_annot(self.bbox)
        elif self.annot_type == "underline":
            annot = page.add_underline_annot(self.bbox)
        else:
            annot = page.add_strikeout_annot(self.bbox)
        self.annot_xref = annot.xref

    def undo(self):
        page = self.doc[self.page_num]
        page.delete_annot(page.load_annot(self.annot_xref))

class ReplaceTextCommand(ActionCommand):
    def __init__(self, engine, page_num, block, old_text, new_text, force_bold=None, force_underline=None):
        self.engine = engine
        self.page_num = page_num
        self.block = block.copy()
        self.old_text = old_text
        self.new_text = new_text
        self.force_bold = force_bold
        self.force_underline = force_underline
        self.last_warning = None
    
    def execute(self):
        success, self.last_warning = self.engine.update_text(self.page_num, self.block, self.new_text, 
                                                            force_bold=self.force_bold,
                                                            force_underline=self.force_underline)
        return success

    def undo(self):
        self.engine.update_text(self.page_num, self.block, self.old_text)

class AddOverlayCommand(ActionCommand):
    def __init__(self, state, page_num, overlay_obj):
        self.state = state
        self.page_num = page_num
        self.overlay_obj = overlay_obj

    def execute(self):
        if self.page_num not in self.state.page_overlays:
            self.state.page_overlays[self.page_num] = []
        self.state.page_overlays[self.page_num].append(self.overlay_obj)

    def undo(self):
        if self.page_num in self.state.page_overlays:
            self.state.page_overlays[self.page_num].remove(self.overlay_obj)

class UpdateOverlayCommand(ActionCommand):
    def __init__(self, overlay_obj, old_props, new_props):
        self.overlay_obj = overlay_obj
        self.old_props = old_props
        self.new_props = new_props

    def execute(self):
        for key, val in self.new_props.items():
            self.overlay_obj[key] = val

    def undo(self):
        for key, val in self.old_props.items():
            self.overlay_obj[key] = val

class DeleteOverlayCommand(ActionCommand):
    def __init__(self, state, page_num, overlay_obj):
        self.state = state
        self.page_num = page_num
        self.overlay_obj = overlay_obj

    def execute(self):
        self.state.page_overlays[self.page_num].remove(self.overlay_obj)

    def undo(self):
        self.state.page_overlays[self.page_num].append(self.overlay_obj)

class AddGroupOverlayCommand(ActionCommand):
    """Command to apply a single overlay to many pages (Apply to All)."""
    def __init__(self, state, source_page, overlay_obj, target_pages):
        self.state = state
        self.source_page = source_page
        self.overlay_obj = overlay_obj
        self.target_pages = target_pages # list of page indices
        self.group_id = f"group_{os.urandom(4).hex()}"
        self.clones = [] # list of (page_num, clone_obj)

    def execute(self):
        self.overlay_obj["group_id"] = self.group_id
        self.clones = []
        for p_idx in self.target_pages:
            if p_idx == self.source_page: continue
            clone = self.overlay_obj.copy()
            if p_idx not in self.state.page_overlays:
                self.state.page_overlays[p_idx] = []
            self.state.page_overlays[p_idx].append(clone)
            self.clones.append((p_idx, clone))

    def undo(self):
        for p_idx, clone in self.clones:
            if clone in self.state.page_overlays.get(p_idx, []):
                self.state.page_overlays[p_idx].remove(clone)
        if "group_id" in self.overlay_obj:
            del self.overlay_obj["group_id"]

class RemoveGroupOverlayCommand(ActionCommand):
    """Command to remove all overlays belonging to a specific group."""
    def __init__(self, state, group_id):
        self.state = state
        self.group_id = group_id
        self.removed_items = [] # list of (page_num, overlay_obj)

    def execute(self):
        self.removed_items = []
        for p_idx, ovs in self.state.page_overlays.items():
            to_remove = [ov for ov in ovs if ov.get("group_id") == self.group_id]
            for ov in to_remove:
                ovs.remove(ov)
                self.removed_items.append((p_idx, ov))

    def undo(self):
        for p_idx, ov in self.removed_items:
            if p_idx not in self.state.page_overlays:
                self.state.page_overlays[p_idx] = []
            self.state.page_overlays[p_idx].append(ov)

class PDFEditorState:
    def __init__(self):
        self.doc = None
        self.current_path = None
        self.current_page_index = 0
        self.zoom_level = 1.0
        self.undo_stack = []
        self.redo_stack = []
        self.selected_object = None # For PDF blocks
        self.selected_overlay = None # For Sign/Stamp
        self.page_overlays = {} # map page_num -> list of dicts
        self.status_message = "Ready"

    def load_doc(self, path):
        self.current_path = path
        self.doc = fitz.open(path)
        self.current_page_index = 0
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.page_overlays = {}
        self.status_message = f"Loaded {os.path.basename(path)}"

    def push_command(self, command):
        """Executes a command and adds it to the undo stack."""
        res = None
        # Pass self (state) to commands that need it (Page actions)
        if isinstance(command, (DeletePageCommand, DuplicatePageCommand)):
            res = command.execute(self)
        else:
            res = command.execute()
            
        self.undo_stack.append(command)
        self.redo_stack.clear()
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        return res

    def undo(self):
        if not self.undo_stack:
            return False
        command = self.undo_stack.pop()
        if isinstance(command, (DeletePageCommand, DuplicatePageCommand)):
            command.undo(self)
        else:
            command.undo()
        self.redo_stack.append(command)
        return True

    def redo(self):
        if not self.redo_stack:
            return False
        command = self.redo_stack.pop()
        if isinstance(command, (DeletePageCommand, DuplicatePageCommand)):
            command.execute(self)
        else:
            command.redo()
        self.undo_stack.append(command)
        return True

    def can_undo(self):
        return len(self.undo_stack) > 0

    def can_redo(self):
        return len(self.redo_stack) > 0
