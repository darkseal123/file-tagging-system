import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import win32com.client
import win32gui
import pythoncom
import os
import re
import threading
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any

# Third-party imports
import send2trash
from pypdf import PdfWriter

# --------------------------
# Helper Class: Tooltip
# --------------------------
class Tooltip:
    def __init__(self, widget: tk.Widget, text: str = 'Tip'):
        self.widget = widget
        self.text = text
        self.id: Optional[int] = None
        self.tw: Optional[tk.Toplevel] = None
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)

    def enter(self, event: Optional[tk.Event] = None) -> None: self.schedule()
    def leave(self, event: Optional[tk.Event] = None) -> None: self.unschedule(); self.hidetip()
    def schedule(self) -> None:
        self.unschedule()
        self.id = self.widget.after(1000, self.showtip)
    def unschedule(self) -> None:
        if self.id: self.widget.after_cancel(self.id); self.id = None
    def showtip(self, event: Optional[tk.Event] = None) -> None:
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(self.tw, text=self.text, background="#ffffe0", relief=tk.SOLID, borderwidth=1, font=("Segoe UI", 9))
        lbl.pack()
    def hidetip(self) -> None:
        if self.tw: self.tw.destroy(); self.tw = None

# --------------------------
# Main Application
# --------------------------
class SirnaomicsTagSystem:
    TAG_FILE_PATH = "taging_system.txt"
    MANUAL_GROUP_NUM = 5
    FRONT_TAG_GROUP_MAX = 5
    TAG_SEPARATOR = "_"
    BACKUP_SUFFIX = ".st_bak"  # Suffix for temporary undo backups
    
    # Office Conversion Config
    SUPPORTED_OFFICE_EXTS = {
        'word': ['.doc', '.docx'],
        'ppt': ['.ppt', '.pptx']
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Sirnaomics Tagging System Pro")
        self.root.geometry("650x820")
        
        # Handle cleanup on exit
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Data State
        self.active_front_tags: List[str] = []
        self.active_back_tags: List[str] = []
        self.tag_buttons: Dict[str, ttk.Button] = {}
        self.selected_files: List[str] = []
        # History Stack: List of Dicts {'type': str, 'data': Any}
        self.history_stack: List[Dict[str, Any]] = []
        self.group_map: Dict[int, Dict] = {}
        
        self._setup_styles()
        self._init_ui()
        self._configure_log_tags()
        self.refresh_tag_database()

    def _setup_styles(self) -> None:
        style = ttk.Style()
        base_font = ("Segoe UI", 9)
        bold_font = ("Segoe UI", 9, "bold")
        large_font = ("Segoe UI", 12, "bold")
        
        style.configure('TButton', font=base_font)
        style.configure('Toggle.TButton', font=base_font)
        style.configure('Selected.TButton', font=bold_font, foreground="#0055aa")
        style.configure('Action.TButton', font=large_font, padding=10)
        style.configure('Tool.TButton', font=base_font, padding=2)
        style.configure('Danger.TButton', font=base_font, foreground="#cc0000")

    def _init_ui(self) -> None:
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # ================= LEFT PANEL: TAGS =================
        left_frame = ttk.LabelFrame(main_pane, text=" 🏷️ Tag Groups ", padding=1)
        main_pane.add(left_frame, weight=1)
        
        self.canvas = tk.Canvas(left_frame, highlightthickness=0, width=280)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_inner = ttk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_inner, anchor="nw")
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scroll_inner.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # ================= RIGHT PANEL: OPERATIONS =================
        right_panel = ttk.Frame(main_pane)
        main_pane.add(right_panel, weight=2)

        # 1. Preview
        preview_frame = ttk.LabelFrame(right_panel, text=" 👁️ Naming Preview ", padding=15)
        preview_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        self.lbl_preview = ttk.Label(preview_frame, text="Filename", font=("Consolas", 11, "bold"), foreground="#0055aa")
        self.lbl_preview.pack()

        # 2. File Tools & Log
        file_frame = ttk.LabelFrame(right_panel, text=" 📂 File Tools & Log ", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)

        # Tools Row
        tools_grid = ttk.Frame(file_frame)
        tools_grid.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(tools_grid, text="PDF Tools:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(tools_grid, text="📄 Office ➔ PDF", style='Tool.TButton', 
                   command=self.tool_convert_to_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(tools_grid, text="📚 Combine PDF", style='Tool.TButton', 
                   command=self.tool_combine_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        ttk.Separator(file_frame, orient='horizontal').pack(fill='x', pady=8)

        # Log Controls
        log_ctrl_grid = ttk.Frame(file_frame)
        log_ctrl_grid.pack(fill=tk.X, pady=(0, 2))
        
        ttk.Button(log_ctrl_grid, text="🔄 Refresh Selection", 
                   command=self.get_selected_files).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(log_ctrl_grid, text="🗑 Clear Log", style='Danger.TButton', 
                   command=self.clear_scrollbox).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        # Text Area
        self.txt_log = scrolledtext.ScrolledText(file_frame, height=8, state='disabled', font=("Consolas", 9))
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=5)

        # 3. Tag Actions
        act_frame = ttk.LabelFrame(right_panel, text=" ✅ Tagging Actions ", padding=10)
        act_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Button(act_frame, text="APPLY TAGS TO SELECTED FILES", style='Action.TButton', 
                   command=self.apply_tags).pack(fill=tk.X, pady=(5, 10))
        
        # Undo/Delete Grid
        del_grid = ttk.Frame(act_frame)
        del_grid.pack(fill=tk.X)
        ttk.Button(del_grid, text="✂ Del Front", command=lambda: self.delete_tag_physically(True)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(del_grid, text="✂ Del Back", command=lambda: self.delete_tag_physically(False)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.btn_undo = ttk.Button(del_grid, text="↩ Undo", state='disabled', command=self.undo_last)
        self.btn_undo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(5, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # --- Resizing Logic ---
    def _on_frame_configure(self, event): self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    def _on_canvas_configure(self, event): self.canvas.itemconfig(self.canvas_window, width=event.width)

    # --- Helper for Thread-Safe UI Updates ---
    def safe_ui_update(self, func, *args):
        self.root.after(0, lambda: func(*args))

    # --- Tool: Office to PDF ---
    def tool_convert_to_pdf(self):
        self.get_selected_files()
        if not self.selected_files: return
        
        valid_files = [p for p in self.selected_files if os.path.splitext(p)[1].lower() in (self.SUPPORTED_OFFICE_EXTS['word'] + self.SUPPORTED_OFFICE_EXTS['ppt'])]
        if not valid_files:
            self.log("No valid Office files selected.", warn=True)
            return

        self.status_var.set(f"Converting {len(valid_files)} files to PDF...")
        self.log(f"Starting conversion for {len(valid_files)} files...")
        
        def convert_task():
            pythoncom.CoInitialize()
            word_app = None
            ppt_app = None
            
            undo_batch = {'created': [], 'restores': []}
            
            try:
                for path in valid_files:
                    if not os.path.exists(path): continue
                    
                    ext = os.path.splitext(path)[1].lower()
                    pdf_path = os.path.splitext(path)[0] + ".pdf"
                    
                    # Convert
                    success = False
                    try:
                        if ext in self.SUPPORTED_OFFICE_EXTS['word']:
                            if not word_app: 
                                word_app = win32com.client.DispatchEx("Word.Application")
                                word_app.Visible = False
                                word_app.DisplayAlerts = 0
                            doc = word_app.Documents.Open(path, ReadOnly=True, Visible=False)
                            doc.SaveAs(pdf_path, FileFormat=17) # wdFormatPDF
                            doc.Close(SaveChanges=False)
                        else:
                            if not ppt_app:
                                ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
                                ppt_app.DisplayAlerts = 0
                            pres = ppt_app.Presentations.Open(path, ReadOnly=True, WithWindow=False)
                            pres.SaveAs(pdf_path, FileFormat=32) # ppSaveAsPDF
                            pres.Close()
                        success = True
                    except Exception as e:
                        self.safe_ui_update(self.log, f"❌ Conversion Failed {os.path.basename(path)}: {str(e)[:50]}", True)

                    # Manage Files: Created PDF and Backup Source
                    if success and os.path.exists(pdf_path):
                        # Backup source instead of delete
                        backup_path = path + self.BACKUP_SUFFIX
                        try:
                            if os.path.exists(backup_path): os.remove(backup_path)
                            os.rename(path, backup_path)
                            
                            undo_batch['created'].append(pdf_path)
                            undo_batch['restores'].append((backup_path, path))
                            self.safe_ui_update(self.log, f"✅ Converted: {os.path.basename(path)}")
                        except Exception as e:
                            self.safe_ui_update(self.log, f"⚠️ Backup Failed: {e}", True)

            finally:
                if word_app: 
                    try: word_app.Quit()
                    except: pass
                if ppt_app: 
                    try: ppt_app.Quit()
                    except: pass
                pythoncom.CoUninitialize()
                
                # Push to history if we did anything
                if undo_batch['created']:
                    self.safe_ui_update(self._push_history, 'file_gen', undo_batch)
                    self.safe_ui_update(self.log, "Office Conversion Task Finished.")
                self.safe_ui_update(self.status_var.set, "Ready")

        threading.Thread(target=convert_task, daemon=True).start()

    # --- Tool: Combine PDF ---
    def tool_combine_pdf(self):
        self.get_selected_files()
        pdf_files = [p for p in self.selected_files if p.lower().endswith('.pdf')]
        
        if len(pdf_files) < 2:
            self.log("Please select at least 2 PDF files to combine.", warn=True)
            return

        output_dir = os.path.dirname(pdf_files[0])
        first_name = os.path.splitext(os.path.basename(pdf_files[0]))[0]
        output_name = f"Merged_{first_name}.pdf"
        output_path = os.path.join(output_dir, output_name)
        
        counter = 1
        while os.path.exists(output_path):
            output_path = os.path.join(output_dir, f"Merged_{first_name}_{counter}.pdf")
            counter += 1

        self.status_var.set("Merging PDFs...")
        try:
            merger = PdfWriter()
            for pdf in pdf_files:
                merger.append(pdf)
            merger.write(output_path)
            merger.close()
            
            self.log(f"✅ Combined PDF saved: {os.path.basename(output_path)}")
            
            # Backup sources
            restores = []
            for f in pdf_files:
                backup = f + self.BACKUP_SUFFIX
                try:
                    if os.path.exists(backup): os.remove(backup)
                    os.rename(f, backup)
                    restores.append((backup, f))
                except Exception as e:
                    self.log(f"⚠️ Failed to move source to backup: {os.path.basename(f)}", warn=True)

            self._push_history('file_gen', {'created': [output_path], 'restores': restores})
            self.log(f"ℹ️ {len(pdf_files)} source files moved to undo buffer.")

        except Exception as e:
            self.log(f"❌ PDF Merge Failed: {e}", err=True)
        finally:
            self.status_var.set("Ready")

    # --- History & Undo Logic ---
    def _push_history(self, action_type: str, data: Any):
        """Add action to history and enable Undo button."""
        self.history_stack.append({'type': action_type, 'data': data})
        self.btn_undo.config(state='normal')

    def undo_last(self):
        if not self.history_stack: return
        
        action = self.history_stack.pop()
        type_ = action['type']
        data = action['data']
        
        self.log(f"Undoing last action ({type_})...")
        
        try:
            # Type 1: Rename (Tagging)
            if type_ == 'rename':
                # data is list of (current_path, original_path)
                count = 0
                for current, original in reversed(data):
                    if os.path.exists(current):
                        try:
                            os.rename(current, original)
                            count += 1
                        except Exception as e:
                            self.log(f"Failed to rename back {os.path.basename(current)}: {e}", err=True)
                self.log(f"Undo complete. {count} files renamed back.")

            # Type 2: File Generation (PDF Tools)
            elif type_ == 'file_gen':
                # data is {'created': [paths], 'restores': [(backup, original)]}
                
                # 1. Delete generated files
                for f in data.get('created', []):
                    if os.path.exists(f):
                        try:
                            send2trash.send2trash(f) # Send result to trash
                            self.log(f"🗑️ Deleted generated: {os.path.basename(f)}")
                        except Exception as e:
                            self.log(f"Failed to delete {os.path.basename(f)}: {e}", err=True)
                
                # 2. Restore source files from backup
                for backup, original in data.get('restores', []):
                    if os.path.exists(backup):
                        try:
                            if os.path.exists(original):
                                self.log(f"⚠️ Target {os.path.basename(original)} exists, overwriting...", warn=True)
                                os.remove(original)
                            os.rename(backup, original)
                            self.log(f"♻️ Restored: {os.path.basename(original)}")
                        except Exception as e:
                            self.log(f"Failed to restore {os.path.basename(original)}: {e}", err=True)
            
            self.status_var.set("Undo successful")

        except Exception as e:
            self.log(f"Critical error during undo: {e}", err=True)

        if not self.history_stack:
            self.btn_undo.config(state='disabled')

    def on_closing(self):
        """Cleanup handler: Send left-over backups to Recycle Bin."""
        if self.history_stack:
            count = 0
            # Scan history for pending backups
            for item in self.history_stack:
                if item['type'] == 'file_gen':
                    for backup, _ in item['data'].get('restores', []):
                        if os.path.exists(backup):
                            try:
                                send2trash.send2trash(backup)
                                count += 1
                            except: pass
            if count > 0:
                print(f"Cleaned up {count} backup files to Recycle Bin.")
        
        self.root.destroy()

    # --- Core Logic ---
    def _configure_log_tags(self) -> None:
        self.txt_log.tag_config('err', foreground='red')
        self.txt_log.tag_config('info', foreground='#000000')
        self.txt_log.tag_config('warn', foreground='#FF8C00')

    def clear_scrollbox(self) -> None:
        self.txt_log.configure(state='normal'); self.txt_log.delete(1.0, tk.END); self.txt_log.configure(state='disabled')

    def log(self, msg: str, err: bool = False, warn: bool = False) -> None:
        self.txt_log.configure(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        tag, icon = 'info', "  "
        if err: tag, icon = 'err', "❌ "
        elif warn: tag, icon = 'warn', "⚠️ "
        self.txt_log.insert(tk.END, f"[{timestamp}] {icon}{msg}\n", tag)
        self.txt_log.see(tk.END); self.txt_log.configure(state='disabled')

    def get_selected_files(self) -> None:
        self.selected_files = []
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            explorer_hwnds = []
            def enum_cb(hwnd, result):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "CabinetWClass": result.append(hwnd)
            win32gui.EnumWindows(enum_cb, explorer_hwnds)
            
            if not explorer_hwnds: 
                self.status_var.set("No Explorer window found"); return
            
            target_hwnd = explorer_hwnds[0]
            found = False
            for win in shell.Windows():
                try:
                    if win.HWND == target_hwnd:
                        items = win.Document.SelectedItems()
                        self.selected_files = [item.Path for item in items]
                        found = True; break
                except: continue
            
            if found and self.selected_files:
                self.log(f"Linked to Explorer: {len(self.selected_files)} items.")
                self.status_var.set(f"Selected: {len(self.selected_files)} files")
            else:
                self.status_var.set("No files selected in Explorer")
        except Exception as e: self.log(f"Explorer Link Error: {e}", err=True)

    def refresh_tag_database(self) -> None:
        for widget in self.scroll_inner.winfo_children(): widget.destroy()
        self.group_map.clear()
        if not os.path.exists(self.TAG_FILE_PATH):
            self.log(f"Config file not found: {self.TAG_FILE_PATH}", warn=True); return

        current_group = None
        try:
            with open(self.TAG_FILE_PATH, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or (line.startswith('#') and not line.lower().startswith('#group')): continue
                    if line.lower().startswith('#group'):
                        match = re.search(r'Group\s+(\d+).*name:\s*(.*)', line, re.I)
                        if match:
                            g_num, g_name = int(match.group(1)), match.group(2).strip()
                            frame = ttk.LabelFrame(self.scroll_inner, text=f" {g_name.upper()} ", padding=1)
                            frame.pack(fill=tk.X, pady=5, padx=2)
                            grid = ttk.Frame(frame); grid.pack(fill=tk.X)
                            grid.columnconfigure(0, weight=1); grid.columnconfigure(1, weight=1)
                            current_group = {'num': g_num, 'grid': grid, 'btn_count': 0}
                            self.group_map[g_num] = current_group
                            if g_num == self.MANUAL_GROUP_NUM: self._setup_manual_group(frame)
                        continue
                    if current_group:
                        parts = line.split(',', 1)
                        name = parts[0].strip()
                        if name == 'yyyymmdd': name = datetime.now().strftime("%Y%m%d")
                        self._add_tag_to_ui(current_group, name, parts[1].strip() if len(parts) > 1 else "")
        except Exception as e: self.log(f"DB Load Error: {e}", err=True)

    def _setup_manual_group(self, frame):
        c = ttk.Frame(frame); c.pack(fill=tk.X, pady=(5,0))
        ttk.Label(c, text="Custom:").pack(side=tk.LEFT)
        ttk.Button(c, text="+", width=3, command=self.add_manual_tag).pack(side=tk.LEFT)
        self.ent_manual = ttk.Entry(c); self.ent_manual.pack(side=tk.LEFT, fill=tk.X, expand=False, padx=2)
 #       ttk.Button(c, text="+", width=3, command=self.add_manual_tag).pack(side=tk.LEFT)
        self.manual_btn_container = ttk.Frame(frame); self.manual_btn_container.pack(fill=tk.X, pady=2)

    def _add_tag_to_ui(self, group, name, comment):
        if group['num'] == self.MANUAL_GROUP_NUM: return
        count = group['btn_count']
        btn = ttk.Button(group['grid'], text=name, style='Toggle.TButton', command=lambda: self.toggle_tag(name, group['num']))
        btn.grid(row=count//2, column=count%2, sticky='ew', padx=1, pady=1)
        group['btn_count'] += 1
        self.tag_buttons[name] = btn
        if comment: Tooltip(btn, comment)

    def toggle_tag(self, name, g_num):
        target = self.active_front_tags if g_num <= self.FRONT_TAG_GROUP_MAX else self.active_back_tags
        if name in target:
            target.remove(name)
            self.tag_buttons[name].configure(style='Toggle.TButton')
        else:
            target.append(name)
            self.tag_buttons[name].configure(style='Selected.TButton')
        self.update_preview()

    def add_manual_tag(self):
        val = self.ent_manual.get().strip()
        if val:
            self.active_front_tags.append(val)
            b = ttk.Button(self.manual_btn_container, text=val, style='Selected.TButton')
            b.config(command=lambda: [self.active_front_tags.remove(val), b.destroy(), self.update_preview()])
            b.pack(side=tk.LEFT, padx=1); self.ent_manual.delete(0, tk.END); self.update_preview()

    def update_preview(self):
        f = self.TAG_SEPARATOR.join(self.active_front_tags) + self.TAG_SEPARATOR if self.active_front_tags else ""
        b = self.TAG_SEPARATOR + self.TAG_SEPARATOR.join(self.active_back_tags) if self.active_back_tags else ""
        self.lbl_preview.config(text=f"{f}FILENAME{b}")

    def apply_tags(self):
        self.get_selected_files()
        if not self.selected_files or (not self.active_front_tags and not self.active_back_tags): return
        f_str = self.TAG_SEPARATOR.join(self.active_front_tags) + self.TAG_SEPARATOR if self.active_front_tags else ""
        b_str = self.TAG_SEPARATOR + self.TAG_SEPARATOR.join(self.active_back_tags) if self.active_back_tags else ""
        
        batch = []
        for path in self.selected_files:
            folder, fname = os.path.split(path)
            name, ext = os.path.splitext(fname)
            new_path = os.path.join(folder, f"{f_str}{name}{b_str}{ext}")
            try:
                os.rename(path, new_path)
                batch.append((new_path, path)) # Store as (current, original) for undo
                self.log(f"✓ Renamed: {os.path.basename(new_path)}")
            except Exception as e: self.log(f"Error: {e}", err=True)
        
        if batch: self._push_history('rename', batch)

    def delete_tag_physically(self, is_front):
        self.get_selected_files()
        batch = []
        for path in self.selected_files:
            folder, fname = os.path.split(path)
            name, ext = os.path.splitext(fname)
            if self.TAG_SEPARATOR not in name: continue
            
            new_name = name
            if is_front:
                if self.TAG_SEPARATOR in name:
                    tag_part, rest = name.split(self.TAG_SEPARATOR, 1)
                    if tag_part in self.active_front_tags or not self.active_front_tags: new_name = rest
            else:
                if self.TAG_SEPARATOR in name:
                    rest, tag_part = name.rsplit(self.TAG_SEPARATOR, 1)
                    new_name = rest
            
            new_path = os.path.join(folder, f"{new_name}{ext}")
            if new_path != path and new_name != name:
                try:
                    os.rename(path, new_path)
                    batch.append((new_path, path))
                    self.log(f"✂ Removed Tag: {os.path.basename(new_path)}")
                except Exception as e: self.log(f"Error: {e}", err=True)
        if batch: self._push_history('rename', batch)

if __name__ == "__main__":
    root = tk.Tk()
    app = SirnaomicsTagSystem(root)
    root.mainloop()