# DocCreatorv5.py - Midmark Excel to Word document generator with template header/footer replacement
#
# 08092025 ADDED CUSTOM ICON: The script now looks for 'MidmarkTLogo.ico' in its
# directory and sets it as the application window icon.
# switched to customtkinter for improved UI elements and styling.

import os
import json
import re
import sys
import threading
import queue
import pythoncom
import pandas as pd
import customtkinter as ctk
import base64
import tempfile
import atexit
from tkinter import filedialog, messagebox
from pathlib import Path
import win32com.client as win32
import time

# REMINDER: Paste your Base64 string here
ICON_BASE64 = "AAABAAEAMC0AAAEAIABQIwAAFgAAACgAAAAwAAAAWgAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEAAAcHBAAHBwQABgYEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQUDAAcHBAAHBwQAAQEBAAAAAAAAAAAABgYEAAcHBAAHBwQAAQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAgIABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABQUDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEAAAcHBAAHBwQABgYEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgYDAAcHBAAHBwQAAQEBAAAAAAAAAAAABQUDAAcHBAAHBwQAAgIBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBAIABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABgYEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwMCAAcHBAAHBwQABwcEAAEBAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgYEAAcHBAAHBwQAAQEAAAAAAAAAAAAAAwMBAAcHBAAHBwQABgYDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBAAAHBwQABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABwcEAAEBAQAAAAAAAAAAAAAAAAAAAAAABgYEAAcHBAAHBwQABwcEAAQEAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAgEABwcEAAcHBAAGBgMAAAAAAAAAAAAAAAAAAQEAAAYGBAAHBwQABwcEAAMDAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUFAwAHBwQABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABwcEAAQEAgAAAAAAAAAAAAAAAAAFBQMABwcEAAcHBAAHBwQABwcEAAcHBAADAwIAAAAAAAAAAAAAAAAAAAAAAAICAQAHBwQABwcEAAcHBAAEBAIAAAAAAAAAAAAAAAAAAAAAAAQEAwAHBwQABwcEAAcHBAAEBAIAAQEAAAAAAAAAAAAAAAAAAAAAAAABAQEABQUDAAcHBAAHBwQABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABwcEAAcHBAAEBAMABQUDAAYGAwAHBwQABwcEAAcHBAADAwIABwcEAAcHBAAHBwQABgYDAAUFAwAEBAIABQUDAAcHBAAHBwQABwcEAAYGBAABAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFBQMABwcEAAcHBAAHBwQABgYEAAUFAwAEBAIABQUDAAUFAwAHBwQABwcEAAcHBAAFBQMABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABwcEAAcHBAAHBwQABwcEAAcHBAAHBwQABwcEAAMDAgAAAAAAAgIBAAcHBAAHBwQABwcEAAcHBAAHBwQABwcEAAcHBAAHBwQABgYEAAICAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQUDAAcHBAAHBwQABwcEAAcHBAAHBwQABwcEAAcHBAAHBwQABwcEAAQEAgACAgEABwcEAAcHBAAFBQMAAAAAAAAAAAAAAAAAAQEBAAcHBAAHBwQABQUDAAUFAwAHBwQABwcEAAcHBAAGBgMAAgIBAAAAAAAAAAAAAAAAAAEBAQAFBQMABwcEAAcHBAAHBwQABwcEAAcHBAAEBAIAAQEAAAAAAAAAAAAAAAAAAAAAA"

# --- CONFIGURATION & DATA HANDLING ---

CONFIG_FILE = Path.home() / "doccreator_last_locations.json"

def load_last_locations():
    """Loads last used locations from a JSON config file."""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}

def save_last_locations(excel_path=None, output_dir=None, template_path=None):
    """Saves last used locations to the JSON config file."""
    try:
        data = load_last_locations()
        if excel_path:
            data["excel_path"] = str(excel_path)
        if output_dir:
            data["output_dir"] = str(output_dir)
        if template_path:
            data["template_path"] = str(template_path)
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f, indent=4)
    except IOError:
        pass

# --- WORD AUTOMATION LOGIC ---

def _replace_placeholder_in_range(doc_range, placeholder, replacement_text):
    """Helper function to perform a find-and-replace in a Word range."""
    find = doc_range.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Text = placeholder
    find.Replacement.Text = replacement_text
    find.Forward = True
    find.Wrap = win32.constants.wdFindContinue
    find.Format = False
    find.MatchCase = False
    find.MatchWholeWord = False
    find.Execute(Replace=win32.constants.wdReplaceAll)

def kill_word_processes():
    """Forcefully terminates all running Microsoft Word processes."""
    try:
        os.system("taskkill /F /IM WINWORD.EXE > nul 2>&1")
        time.sleep(1)
    except Exception:
        pass

def document_generation_worker(excel_path, output_folder, template_path, progress_queue):
    """
    Worker function to be run in a separate thread.
    Handles the core logic of reading the Excel file and generating Word documents.
    """
    pythoncom.CoInitialize()
    word = None
    try:
        # --- 1. Read Excel Data ---
        try:
            df = pd.read_excel(excel_path).fillna('')
            total_docs = len(df)
            progress_queue.put(("set_max", total_docs))
        except Exception as e:
            progress_queue.put(("error", f"Error Reading Excel File: {e}"))
            return

        # --- 2. Kill Existing Word Processes & Initialize New Instance ---
        kill_word_processes()
        try:
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
        except Exception as e:
            progress_queue.put(("error", f"Could not launch MS Word: {e}"))
            return

        # --- 3. Process Each Row (One Row = One Document) ---
        filename_col = df.columns[0]
        success_count = 0
        
        heading_pattern = re.compile(r'^H(\d+(-\d+)*)$')

        for i, row in df.iterrows():
            filename = str(row[filename_col]).strip()
            if not filename:
                progress_queue.put(("log", "[SKIPPED] Row with missing filename.\n"))
                progress_queue.put(("progress", i + 1))
                continue
            
            progress_queue.put(("log", f"Processing: {filename}\n"))

            try:
                doc = word.Documents.Add(Template=str(template_path))
                
                # --- HIERARCHICAL LOGIC WITH STATIC TEXT COLUMNS ---
                for col_header in df.columns:
                    if heading_pattern.match(col_header):
                        heading_content = str(row.get(col_header, "")).strip()
                        
                        if heading_content:
                            level = len(col_header.split('-'))
                            word_style = f"Heading {level}"
                            
                            progress_queue.put(("log", f"  -> Style: '{word_style}', Content: '{heading_content[:40]}...'\n"))
                            
                            word.Selection.EndKey(Unit=win32.constants.wdStory)
                            word.Selection.Style = doc.Styles(word_style)
                            word.Selection.TypeText(Text=heading_content)
                            heading_indent = word.Selection.ParagraphFormat.LeftIndent

                            # NEW STATIC LOOKUP: Look for 'H1-1-text', etc.
                            text_col_name = f"{col_header}-text"
                            if text_col_name in df.columns:
                                body_text = str(row.get(text_col_name, "")).strip()
                                if body_text:
                                    progress_queue.put(("log", f"    -> Adding body text from '{text_col_name}'\n"))
                                    word.Selection.TypeParagraph()
                                    word.Selection.Style = doc.Styles("Normal")
                                    word.Selection.ParagraphFormat.LeftIndent = heading_indent
                                    word.Selection.TypeText(Text=body_text)

                            word.Selection.TypeParagraph()
                
                # --- ADD STATIC REFERENCES SECTION ---
                progress_queue.put(("log", "  -> Adding static 'References' section...\n"))
                word.Selection.EndKey(Unit=win32.constants.wdStory)
                word.Selection.Style = doc.Styles("Heading 1")
                word.Selection.TypeText(Text="References")
                ref_heading_indent = word.Selection.ParagraphFormat.LeftIndent
                word.Selection.TypeParagraph()
                word.Selection.Style = doc.Styles("Normal") # Reset style for the table
                
                # --- Get page dimensions for table width calculation ---
                section = doc.Sections(1)
                page_width = section.PageSetup.PageWidth
                left_margin = section.PageSetup.LeftMargin
                right_margin = section.PageSetup.RightMargin
                usable_width = page_width - left_margin - right_margin

                # Add the table
                table_range = word.Selection.Range
                table = doc.Tables.Add(Range=table_range, NumRows=6, NumColumns=3)
                
                # Style the table
                table.Style = "Table Grid"
                table.Borders.Enable = True
                
                # Set table alignment and width
                table.Rows.LeftIndent = ref_heading_indent
                table.PreferredWidthType = win32.constants.wdPreferredWidthPoints
                # Adjust total width by the heading's indent to align correctly
                table.PreferredWidth = usable_width - ref_heading_indent

                # Populate header row and make it bold
                headers = ["Control No.", "Revision", "Description"]
                for col_idx, header_text in enumerate(headers, 1):
                    cell = table.Cell(Row=1, Column=col_idx)
                    cell.Range.Text = header_text
                    cell.Range.Font.Bold = True
                    cell.VerticalAlignment = win32.constants.wdCellAlignVerticalCenter
                
                # Move cursor after the table
                word.Selection.EndKey(Unit=win32.constants.wdStory)
                word.Selection.TypeParagraph()


                # --- Cleanup and Header/Footer ---
                if doc.Paragraphs.Count > 0 and doc.Paragraphs(1).Range.Text.strip() == "":
                    doc.Paragraphs(1).Range.Delete()
                if doc.Paragraphs.Count > 0 and doc.Paragraphs(1).Format.PageBreakBefore:
                    doc.Paragraphs(1).Format.PageBreakBefore = False

                if doc.Sections.Count > 0:
                    header = doc.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary)
                    footer = doc.Sections(1).Footers(win32.constants.wdHeaderFooterPrimary)
                    title_text = str(row.get("Title", "")).strip()
                    if title_text:
                        _replace_placeholder_in_range(header.Range, "<title>", title_text)
                        _replace_placeholder_in_range(header.Range, "<header-title>", title_text)
                    footer_right_text = str(row.get("Footer-Right", "")).strip()
                    _replace_placeholder_in_range(footer.Range, "<footer-right>", footer_right_text)
                    footer_right2_text = str(row.get("Footer-Right2", "")).strip()
                    _replace_placeholder_in_range(footer.Range, "<footer-right2>", footer_right2_text)

                # --- Save and Close ---
                safe_name = "".join(c for c in filename if c.isalnum() or c in "._- ")
                doc_path = Path(output_folder) / f"{safe_name}.docx"
                doc.SaveAs(str(doc_path))
                doc.Close(SaveChanges=False)

                success_count += 1
                progress_queue.put(("log", f"[SUCCESS] {safe_name}.docx\n"))

            except Exception as e:
                import traceback
                error_details = traceback.format_exc()
                progress_queue.put(("log", f"[FAILED] {filename}: {e}\n{error_details}\n"))
            
            progress_queue.put(("progress", i + 1))
            time.sleep(0.1)

        progress_queue.put(("done", success_count, total_docs))

    finally:
        if word:
            word.Quit()
        pythoncom.CoUninitialize()


# --- CUSTOMTKINTER GUI APPLICATION ---
class DocCreatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.output_folder = None
        self.template_path = None
        self.progress_queue = queue.Queue()
        self.total_docs = 0
        self.temp_icon_path = None

        # --- UI Configuration ---
        self.left_column_weight = 1
        self.right_column_weight = 1
        font_family = "Century Gothic"
        self.fonts = {
            "header": ctk.CTkFont(family=font_family, size=20, weight="bold"),
            "card_title": ctk.CTkFont(family=font_family, size=16, weight="bold"),
            "button": ctk.CTkFont(family=font_family, size=14, weight="bold"),
            "status": ctk.CTkFont(family=font_family, size=14),
            "footer": ctk.CTkFont(family=font_family, size=12),
            "log": (font_family, 16)
        }

        self.after(10, self._deferred_setup)

    def _deferred_setup(self):
        # --- Window Size & Centering ---
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = int(screen_width * 0.50)
        height = int(screen_height * 0.50)
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.geometry(f"{width}x{height}+{x}+{y}")

        self._setup_window_icon()
        self.setup_ui()
        self.load_initial_paths()

        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.after(100, self.poll_queue)

    def _setup_window_icon(self):
        if ICON_BASE64:
            try:
                icon_data = base64.b64decode(ICON_BASE64)
                with tempfile.NamedTemporaryFile(delete=False, suffix='.ico') as temp_icon:
                    self.temp_icon_path = temp_icon.name
                    temp_icon.write(icon_data)
                
                self.iconbitmap(self.temp_icon_path)
                atexit.register(self._cleanup_temp_icon)
            except Exception as e:
                print(f"Icon Error: {e}")

    def _cleanup_temp_icon(self):
        if self.temp_icon_path and os.path.exists(self.temp_icon_path):
            os.remove(self.temp_icon_path)

    def _on_closing(self):
        self._cleanup_temp_icon()
        self.destroy()

    def _find_normal_template(self):
        try:
            appdata = os.getenv('APPDATA')
            template_path = Path(appdata) / "Microsoft" / "Templates" / "Normal.dotm"
            return template_path if template_path.exists() else None
        except Exception:
            return None

    def _show_com_error(self, error_exception):
        messagebox.showerror("Word Error", f"Could not open Word to edit template:\n{error_exception}")

    def _edit_template_worker(self, template_path):
        pythoncom.CoInitialize()
        try:
            word = win32.Dispatch("Word.Application")
            word.Visible = True
            word.Documents.Open(str(template_path))
        except Exception as e:
            self.after(0, self._show_com_error, e)
        finally:
            pythoncom.CoUninitialize()

    def setup_ui(self):
        self.colors = {
            "gray-900": "#111827", "gray-800": "#1F2937", "gray-700": "#374151",
            "gray-600": "#4B5563", "gray-500": "#6B7280", "gray-400": "#9CA3AF",
            "gray-300": "#D1D5DB", "gray-200": "#E5E7EB", "blue-600": "#2563EB",
            "blue-700": "#1D4ED8", "green-400": "#4ADE80",
        }

        self.title("Excel to Word Document Generator")
        self.minsize(800, 600)
        self.configure(fg_color=self.colors["gray-900"])

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        header = ctk.CTkFrame(self, fg_color=self.colors["gray-800"], corner_radius=0, height=60)
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(header, text="Excel to Word Document Generator", font=self.fonts["header"], text_color=self.colors["gray-200"]).pack(side="left", padx=20, pady=10)

        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=6)
        main_container.grid_columnconfigure(0, weight=self.left_column_weight)
        main_container.grid_columnconfigure(1, weight=self.right_column_weight)
        main_container.grid_rowconfigure(0, weight=1)

        left_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=12)
        left_frame.grid_columnconfigure(0, weight=1)

        right_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=12)
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(1, weight=1)

        # --- Left Frame Content ---
        step1_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8)
        step1_card.grid(row=0, column=0, sticky="ew", pady=(10, 12))
        step1_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(step1_card, text="Step 1: Select Word Template (.dotm)", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, columnspan=3, padx=24, pady=(20, 16), sticky="ew")
        step1_inner = ctk.CTkFrame(step1_card, fg_color="transparent")
        step1_inner.grid(row=1, column=0, columnspan=3, padx=24, pady=(0, 24), sticky="ew")
        step1_inner.grid_columnconfigure(0, weight=1)
        self.template_path_label = ctk.CTkLabel(step1_inner, text="No template selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.template_path_label.grid(row=0, column=0, sticky="ew")
        self.edit_template_btn = ctk.CTkButton(step1_inner, text="Edit", command=self.open_template_file, width=70, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"])
        self.edit_template_btn.grid(row=0, column=1, padx=(16, 8))
        ctk.CTkButton(step1_inner, text="Browse...", command=self.select_template_file, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=2, padx=(0, 0))

        step2_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8)
        step2_card.grid(row=1, column=0, sticky="ew", pady=12)
        step2_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(step2_card, text="Step 2: Select Excel File", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, columnspan=3, padx=24, pady=(20, 16), sticky="ew")
        step2_inner = ctk.CTkFrame(step2_card, fg_color="transparent")
        step2_inner.grid(row=1, column=0, columnspan=3, padx=24, pady=(0, 24), sticky="ew")
        step2_inner.grid_columnconfigure(0, weight=1)
        self.excel_path_label = ctk.CTkLabel(step2_inner, text="No file selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.excel_path_label.grid(row=0, column=0, sticky="ew")
        self.edit_excel_btn = ctk.CTkButton(step2_inner, text="Edit", command=self.open_excel_file, width=70, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"])
        self.edit_excel_btn.grid(row=0, column=1, padx=(16, 8))
        ctk.CTkButton(step2_inner, text="Browse...", command=self.select_excel_file, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=2, padx=(0, 0))

        step3_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8)
        step3_card.grid(row=2, column=0, sticky="ew", pady=12)
        step3_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(step3_card, text="Step 3: Select Output Folder", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, columnspan=2, padx=24, pady=(20, 16), sticky="ew")
        step3_inner = ctk.CTkFrame(step3_card, fg_color="transparent")
        step3_inner.grid(row=1, column=0, columnspan=2, padx=24, pady=(0, 24), sticky="ew")
        step3_inner.grid_columnconfigure(0, weight=1)
        self.output_folder_label = ctk.CTkLabel(step3_inner, text="No folder selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.output_folder_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(step3_inner, text="Browse...", command=self.select_output_folder, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=1, padx=(16, 0))

        self.generate_btn = ctk.CTkButton(left_frame, text="Generate Documents", state="disabled", command=self.start_generation, height=52, fg_color=self.colors["blue-600"], hover_color=self.colors["blue-700"], font=self.fonts["card_title"])
        self.generate_btn.grid(row=3, column=0, sticky="ew", pady=(24, 10))

        status_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        status_frame.grid(row=4, column=0, pady=8)
        self.status_label = ctk.CTkLabel(status_frame, text="Status: Ready", font=self.fonts["status"], text_color=self.colors["green-400"])
        self.status_label.grid(row=0, column=0, padx=(0, 8))
        self.progress_bar = ctk.CTkProgressBar(left_frame, mode="determinate", height=8, fg_color=self.colors["gray-700"], progress_color=self.colors["green-400"])
        self.progress_bar.set(0)

        # --- Right Frame Content ---
        status_log_label = ctk.CTkLabel(right_frame, text="Status Log", font=self.fonts["card_title"], text_color=self.colors["gray-300"])
        status_log_label.grid(row=0, column=0, sticky="w", pady=(10, 8))
        self.status_box = ctk.CTkTextbox(right_frame, wrap="word", font=self.fonts["log"], state="disabled", fg_color=self.colors["gray-800"], border_color=self.colors["gray-600"], border_width=1, text_color=self.colors["gray-300"], corner_radius=8)
        self.status_box.grid(row=1, column=0, sticky="nsew")
        self.open_folder_btn = ctk.CTkButton(right_frame, text="Open Output Folder", state="disabled", command=self.open_output_folder, height=52, fg_color=self.colors["gray-700"], hover_color=self.colors["gray-600"], border_color=self.colors["gray-600"], border_width=1, font=self.fonts["button"])
        self.open_folder_btn.grid(row=2, column=0, sticky="ew", pady=(16, 10))

        # --- Footer ---
        footer = ctk.CTkFrame(self, fg_color=self.colors["gray-800"], corner_radius=0)
        footer.grid(row=2, column=0, sticky="ew")
        ctk.CTkLabel(footer, text="© 2025 Midmark Corporation. All rights reserved.", text_color=self.colors["gray-500"], font=self.fonts["footer"]).pack(pady=10)

    def load_initial_paths(self):
        last_locations = load_last_locations()
        last_template = last_locations.get("template_path")
        if last_template and Path(last_template).is_file():
            self.template_path = Path(last_template)
        else:
            self.template_path = self._find_normal_template()
        if self.template_path:
            self.template_path_label.configure(text=str(self.template_path), text_color=self.colors["gray-200"])
        
        last_excel = last_locations.get("excel_path")
        if last_excel and Path(last_excel).is_file():
            self.excel_path = Path(last_excel)
            self.excel_path_label.configure(text=self.excel_path.name, text_color=self.colors["gray-200"])
        
        last_output = last_locations.get("output_dir")
        if last_output and Path(last_output).is_dir():
            self.output_folder = Path(last_output)
            self.output_folder_label.configure(text=str(self.output_folder), text_color=self.colors["gray-200"])
        
        self.update_generate_button_state()

    def select_template_file(self):
        initial_dir = self.template_path.parent if self.template_path and self.template_path.exists() else Path.home()
        file_path_str = filedialog.askopenfilename(title="Select Word Template File", filetypes=[("Word Templates", "*.dotm")], initialdir=initial_dir)
        if file_path_str:
            self.template_path = Path(file_path_str)
            self.template_path_label.configure(text=str(self.template_path), text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(template_path=self.template_path)

    def open_template_file(self):
        if self.template_path and self.template_path.is_file():
            thread = threading.Thread(target=self._edit_template_worker, args=(self.template_path,))
            thread.daemon = True
            thread.start()

    def open_excel_file(self):
        if self.excel_path and self.excel_path.is_file():
            os.startfile(self.excel_path)

    def select_excel_file(self):
        last_dir = Path(self.excel_path).parent if self.excel_path else Path.home()
        file_path_str = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")], initialdir=last_dir)
        if not file_path_str: return
        try:
            pd.read_excel(file_path_str, nrows=1)
        except Exception as e:
            messagebox.showerror("Excel Error", f"Could not read the Excel file:\n{e}")
            return
        self.excel_path = Path(file_path_str)
        self.excel_path_label.configure(text=self.excel_path.name, text_color=self.colors["gray-200"])
        self.update_generate_button_state()
        save_last_locations(excel_path=self.excel_path)

    def select_output_folder(self):
        initial_dir = self.output_folder if self.output_folder else Path.home()
        folder_path_str = filedialog.askdirectory(title="Select Output Folder", initialdir=initial_dir)
        if folder_path_str:
            self.output_folder = Path(folder_path_str)
            self.output_folder_label.configure(text=str(self.output_folder), text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(output_dir=self.output_folder)

    def update_generate_button_state(self):
        state = "normal" if self.excel_path and self.output_folder and self.template_path else "disabled"
        self.generate_btn.configure(state=state)
        self.edit_template_btn.configure(state="normal" if self.template_path else "disabled")
        self.edit_excel_btn.configure(state="normal" if self.excel_path else "disabled")

    def open_output_folder(self):
        if self.output_folder and self.output_folder.is_dir():
            os.startfile(self.output_folder)

    def start_generation(self):
        self.generate_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.progress_bar.grid(row=5, column=0, sticky="ew", pady=(8, 10))
        self.progress_bar.set(0)
        self.status_label.configure(text="Status: Initializing...", text_color=self.colors["gray-400"])
        self.status_box.configure(state="normal")
        self.status_box.delete("1.0", "end")
        self.worker_thread = threading.Thread(target=document_generation_worker, args=(self.excel_path, self.output_folder, self.template_path, self.progress_queue))
        self.worker_thread.start()

    def poll_queue(self):
        try:
            message = self.progress_queue.get_nowait()
            msg_type, *payload = message
            if msg_type == "error":
                messagebox.showerror("Error", payload[0])
                self.reset_ui()
            elif msg_type == "set_max":
                self.total_docs = payload[0]
                self.status_label.configure(text="Status: Generating documents...", text_color=self.colors["gray-400"])
            elif msg_type == "progress":
                if self.total_docs > 0:
                    progress_float = payload[0] / self.total_docs
                    self.progress_bar.set(progress_float)
                    percent = int(progress_float * 100)
                    self.status_label.configure(text=f"Status: In progress... {percent}%", text_color=self.colors["gray-400"])
            elif msg_type == "log":
                self.status_box.insert("end", payload[0])
                self.status_box.see("end")
            elif msg_type == "done":
                success_count, total_docs = payload
                self.status_label.configure(text=f"Status: Finished. {success_count}/{total_docs} documents created.", text_color=self.colors["green-400"])
                self.reset_ui(finished=True)
        except queue.Empty:
            pass
        finally:
            self.after(100, self.poll_queue)

    def reset_ui(self, finished=False):
        self.progress_bar.grid_remove()
        self.generate_btn.configure(state="normal")
        self.status_box.configure(state="disabled")
        if finished:
            self.open_folder_btn.configure(state="normal")

if __name__ == "__main__":
    app = DocCreatorApp()
    app.mainloop()
