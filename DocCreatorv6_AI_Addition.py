# DocCreatorv7_AI_Controls.py 
# Midmark Excel to Word Doc Generator
# Author: Trinidad Hernandez
# Description: Automates Word Template documents from Excel data,
#              with support for custom templates and dynamic content replacement. 

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
import google.generativeai as genai

ICON_BASE64 = ""

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

# --- MODIFICATION START ---
# --- AI CONTENT GENERATION ---
def generate_ai_sectional_content(intended_use, document_title, section_header, additional_prompts):
    """
    Calls the Gemini API to generate new content for a specific section header.
    
    Args:
        intended_use (str): The Intended Use statement from the user's input box.
        document_title (str): The title of the specific document being generated.
        section_header (str): The specific header of the section needing content.
        additional_prompts (list): A list of tuples with extra instructions (e.g., [("Target Audience", "Engineers")]).
        
    Returns:
        str: The AI-generated sectional content, or an error message if the call fails.
    """
    try:
        api_key = "AIzaSyDf3cgcfhAdktmHWLkWOIqhi8OlxzzTFR8"
        if not api_key:
            return "Error: No Gemini API key found"
        
        genai.configure(api_key=api_key)
        
        model = genai.GenerativeModel('gemini-2.5-flash-lite')
        
        # --- Build the prompt dynamically ---
        prompt_parts = [
            "You are a technical writer creating a specific section for a document.",
            f"DOCUMENT TITLE: \"{document_title}\"",
            f"DOCUMENT'S INTENDED USE: \"{intended_use}\"",
            f"CURRENT SECTION HEADER: \"{section_header}\"\n"
        ]

        # Add additional instructions if they exist
        if additional_prompts:
            instruction_parts = ["Please adhere to the following additional guidelines:"]
            for label, text in additional_prompts:
                instruction_parts.append(f"- {label}: {text}")
            prompt_parts.append("\n".join(instruction_parts))

        prompt_parts.append(
            f"\nWrite the body text for \"{section_header}\"."
            f"Rules: Output only the paragraph text, do not repeat the header. Use no special characters. Use brief sentences."
        
        )
        
        full_prompt = "\n".join(prompt_parts)
        
        response = model.generate_content(full_prompt)
        return response.text
    except Exception as e:
        print(f"Gemini API Error: {e}")
        return f"[Error generating AI content for this section: {e}]"
# --- MODIFICATION END ---


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

# --- MODIFICATION START ---
def document_generation_worker(excel_path, output_folder, template_path, progress_queue, ai_enabled, intended_use_statement, additional_prompts):
# --- MODIFICATION END ---
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

        # --- 3. Process Each Row (One Document per Row) ---
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
                
                document_title = str(row.get("Title", "")).strip()

                for col_header in df.columns:
                    if heading_pattern.match(col_header):
                        heading_content = str(row.get(col_header, "")).strip()
                        
                        if heading_content:
                            level = len(col_header.split('-'))
                            word_style = f"Heading {level}"
                            progress_queue.put(("log", f"  -> Adding Header: '{heading_content[:40]}...'\n"))
                            word.Selection.EndKey(Unit=win32.constants.wdStory)
                            word.Selection.Style = doc.Styles(word_style)
                            word.Selection.TypeText(Text=heading_content)
                            heading_indent = word.Selection.ParagraphFormat.LeftIndent

                            if ai_enabled:
                                if document_title and intended_use_statement.strip():
                                    progress_queue.put(("log", f"    -> Calling Gemini for '{heading_content}' section...\n"))
                                    # --- MODIFICATION START ---
                                    ai_text = generate_ai_sectional_content(intended_use_statement, document_title, heading_content, additional_prompts)
                                    # --- MODIFICATION END ---
                                    
                                    word.Selection.TypeParagraph()
                                    word.Selection.Style = doc.Styles("Normal")
                                    word.Selection.ParagraphFormat.LeftIndent = heading_indent
                                    word.Selection.TypeText(Text=ai_text)
                                else:
                                    progress_queue.put(("log", "    -> [SKIPPED AI] Missing 'Title' in Excel or Intended Use statement.\n"))
                            else:
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
                
                # --- ADD GENERIC STATIC REFERENCES SECTION ---
                progress_queue.put(("log", "  -> Adding static 'References' section...\n"))
                word.Selection.EndKey(Unit=win32.constants.wdStory)
                word.Selection.Style = doc.Styles("Heading 1")
                word.Selection.TypeText(Text="References")
                ref_heading_indent = word.Selection.ParagraphFormat.LeftIndent
                word.Selection.TypeParagraph()
                word.Selection.Style = doc.Styles("Normal")
                
                section = doc.Sections(1)
                page_width = section.PageSetup.PageWidth
                left_margin = section.PageSetup.LeftMargin
                right_margin = section.PageSetup.RightMargin
                usable_width = page_width - left_margin - right_margin

                table_range = word.Selection.Range
                table = doc.Tables.Add(Range=table_range, NumRows=6, NumColumns=3)
                table.Style = "Table Grid"
                table.Borders.Enable = True
                table.Rows.LeftIndent = ref_heading_indent
                table.PreferredWidthType = win32.constants.wdPreferredWidthPoints
                table.PreferredWidth = usable_width - ref_heading_indent

                headers = ["Control No.", "Revision", "Description"]
                for col_idx, header_text in enumerate(headers, 1):
                    cell = table.Cell(Row=1, Column=col_idx)
                    cell.Range.Text = header_text
                    cell.Range.Font.Bold = True
                    cell.VerticalAlignment = win32.constants.wdCellAlignVerticalCenter
                
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
                    if document_title:
                        _replace_placeholder_in_range(header.Range, "<title>", document_title)
                        _replace_placeholder_in_range(header.Range, "<header-title>", document_title)
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

        self.left_column_weight = 0
        self.right_column_weight = 1
        font_family = "Century Gothic"
        self.fonts = {
            "header": ctk.CTkFont(family=font_family, size=18, weight="normal"),
            "card_title": ctk.CTkFont(family=font_family, size=16, weight="normal"),
            "button": ctk.CTkFont(family=font_family, size=14, weight="normal"),
            "status": ctk.CTkFont(family=font_family, size=14),
            "footer": ctk.CTkFont(family=font_family, size=16),
            "log": (font_family, 12)
        }
        self.after(10, self._deferred_setup)

    def _deferred_setup(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = int(screen_width * 0.55)
        # --- MODIFICATION START ---
        height = int(screen_height * 0.8) # Increased height for new controls
        # --- MODIFICATION END ---
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
        self.title("Midmark Batch Document Generator")
        # --- MODIFICATION START ---
        self.minsize(800, 850) # Increased min height
        # --- MODIFICATION END ---
        self.configure(fg_color=self.colors["gray-900"])

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color=self.colors["gray-800"], corner_radius=0, height=60)
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(header, text="Batch Document Generator", font=self.fonts["header"], text_color=self.colors["gray-200"]).pack(side="left", padx=20, pady=10)

        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=6)
        main_container.grid_columnconfigure(0, weight=self.left_column_weight)
        main_container.grid_columnconfigure(1, weight=self.right_column_weight)
        main_container.grid_rowconfigure(0, weight=1)

        left_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=60, pady=0)
        left_frame.grid_columnconfigure(0, weight=1)

        right_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=1)
        right_frame.grid_columnconfigure(0, weight=1,minsize=200)
        right_frame.grid_rowconfigure(1, weight=1)

        # --- File Selection Cards ---
        step1_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8,border_color=self.colors["gray-600"], border_width=1)
        step1_card.grid(row=0, column=0, sticky="ew", pady=(10, 6))
        # ... (rest of step 1 card is unchanged) ...
        ctk.CTkLabel(step1_card, text="Step 1: Select Word Template (.dotm)", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, padx=24, pady=(20, 16), sticky="ew")
        step1_inner = ctk.CTkFrame(step1_card, fg_color="transparent")
        step1_inner.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="ew")
        step1_inner.grid_columnconfigure(0, weight=1)
        self.template_path_label = ctk.CTkLabel(step1_inner, text="No template selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.template_path_label.grid(row=0, column=0, sticky="ew")
        self.edit_template_btn = ctk.CTkButton(step1_inner, text="Edit", command=self.open_template_file, width=70, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"])
        self.edit_template_btn.grid(row=0, column=1, padx=(16, 8))
        ctk.CTkButton(step1_inner, text="Browse...", command=self.select_template_file, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=2)

        step2_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8, border_color=self.colors["gray-600"], border_width=1)
        step2_card.grid(row=1, column=0, sticky="ew", pady=6)
        # ... (rest of step 2 card is unchanged) ...
        ctk.CTkLabel(step2_card, text="Step 2: Select Excel File", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, padx=24, pady=(20, 16), sticky="ew")
        step2_inner = ctk.CTkFrame(step2_card, fg_color="transparent")
        step2_inner.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="ew")
        step2_inner.grid_columnconfigure(0, weight=1)
        self.excel_path_label = ctk.CTkLabel(step2_inner, text="No file selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.excel_path_label.grid(row=0, column=0, sticky="ew")
        self.edit_excel_btn = ctk.CTkButton(step2_inner, text="Edit", command=self.open_excel_file, width=70, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"])
        self.edit_excel_btn.grid(row=0, column=1, padx=(16, 8))
        ctk.CTkButton(step2_inner, text="Browse...", command=self.select_excel_file, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=2)
        
        # --- AI Controls Card ---
        ai_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8, border_color=self.colors["gray-600"], border_width=1)
        ai_card.grid(row=2, column=0, sticky="ew", pady=6)
        ai_card.grid_columnconfigure(0, weight=1)
        
        ai_header_frame = ctk.CTkFrame(ai_card, fg_color="transparent")
        ai_header_frame.grid(row=0, column=0, padx=24, pady=(20, 16), sticky="ew")
        ai_header_frame.grid_columnconfigure(1, weight=1)

        self.ai_switch_var = ctk.BooleanVar(value=False)
        ai_switch = ctk.CTkSwitch(ai_header_frame, text="", variable=self.ai_switch_var, switch_height=20, switch_width=40)
        ai_switch.grid(row=0, column=0, sticky="w")
        
        ctk.CTkLabel(ai_header_frame, text="Step 3: AI Content Generation", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=1, padx=(12,0), sticky="w")

        # Intended Use Textbox
        ctk.CTkLabel(ai_card, text="Intended Use Statement", font=self.fonts["button"], text_color=self.colors["gray-400"]).grid(row=1, column=0, padx=24, pady=(0,4), sticky="w")
        self.intended_use_textbox = ctk.CTkTextbox(ai_card, height=80, wrap="word", font=self.fonts["log"], fg_color=self.colors["gray-700"], text_color=self.colors["gray-200"], border_width=0, corner_radius=6)
        self.intended_use_textbox.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 12))
        self.intended_use_textbox.insert("1.0", "Example: This document is for trained technicians to perform annual maintenance on the Midmark XYZ model.")
        
        # --- MODIFICATION START ---
        # --- Additional Prompt Controls ---
        additional_prompts_frame = ctk.CTkFrame(ai_card, fg_color="transparent")
        additional_prompts_frame.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 24))
        additional_prompts_frame.grid_columnconfigure(1, weight=1)

        # -- Prompt Row 1: Audience --
        self.prompt_cb_var_1 = ctk.BooleanVar(value=False)
        self.prompt_checkbox_1 = ctk.CTkCheckBox(additional_prompts_frame, text="Target Audience", variable=self.prompt_cb_var_1, font=self.fonts["button"])
        self.prompt_checkbox_1.grid(row=0, column=0, sticky="w")
        self.prompt_entry_1 = ctk.CTkEntry(additional_prompts_frame, placeholder_text="e.g., Clinical Staff, Engineers", font=self.fonts["log"])
        self.prompt_entry_1.grid(row=0, column=1, sticky="ew", padx=(10,0))
        
        # -- Prompt Row 2: Tone --
        self.prompt_cb_var_2 = ctk.BooleanVar(value=False)
        self.prompt_checkbox_2 = ctk.CTkCheckBox(additional_prompts_frame, text="Tone of Voice", variable=self.prompt_cb_var_2, font=self.fonts["button"])
        self.prompt_checkbox_2.grid(row=1, column=0, sticky="w", pady=(8,0))
        self.prompt_entry_2 = ctk.CTkEntry(additional_prompts_frame, placeholder_text="e.g., Formal, Professional, Casual", font=self.fonts["log"])
        self.prompt_entry_2.grid(row=1, column=1, sticky="ew", padx=(10,0), pady=(8,0))
        
        # -- Prompt Row 3: Keywords --
        self.prompt_cb_var_3 = ctk.BooleanVar(value=False)
        self.prompt_checkbox_3 = ctk.CTkCheckBox(additional_prompts_frame, text="Include Keywords", variable=self.prompt_cb_var_3, font=self.fonts["button"])
        self.prompt_checkbox_3.grid(row=2, column=0, sticky="w", pady=(8,0))
        self.prompt_entry_3 = ctk.CTkEntry(additional_prompts_frame, placeholder_text="e.g., sterilization, compliance, maintenance", font=self.fonts["log"])
        self.prompt_entry_3.grid(row=2, column=1, sticky="ew", padx=(10,0), pady=(8,0))
        # --- MODIFICATION END ---
        
        # --- Output Folder and Generate Button ---
        step4_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8, border_color=self.colors["gray-600"], border_width=1)
        step4_card.grid(row=3, column=0, sticky="ew", pady=6)
        # ... (rest of step 4 card is unchanged) ...
        ctk.CTkLabel(step4_card, text="Step 4: Select Output Folder", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, padx=24, pady=(20, 16), sticky="ew")
        step4_inner = ctk.CTkFrame(step4_card, fg_color="transparent")
        step4_inner.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="ew")
        step4_inner.grid_columnconfigure(0, weight=1)
        self.output_folder_label = ctk.CTkLabel(step4_inner, text="No folder selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.output_folder_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(step4_inner, text="Browse...", command=self.select_output_folder, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=1, padx=(16, 0))


        self.generate_btn = ctk.CTkButton(left_frame, text="Generate Documents", state="disabled", command=self.start_generation, height=52, fg_color=self.colors["blue-600"], hover_color=self.colors["blue-700"], font=self.fonts["card_title"])
        self.generate_btn.grid(row=4, column=0,  sticky="ew", pady=(12, 6))
        
        status_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        status_frame.grid(row=5, column=0, pady=4)
        self.status_label = ctk.CTkLabel(status_frame, text="Status: Ready", font=self.fonts["status"], text_color=self.colors["green-400"])
        self.status_label.grid(row=0, column=0, padx=(0, 8))
        self.progress_bar = ctk.CTkProgressBar(left_frame, mode="determinate", height=16, fg_color=self.colors["gray-700"], progress_color=self.colors["green-400"])
        self.progress_bar.set(0)

        # Right Frame Content
        status_log_label = ctk.CTkLabel(right_frame, text="Status Log", font=self.fonts["card_title"], text_color=self.colors["gray-300"])
        status_log_label.grid(row=0, column=0,  sticky="w", pady=(10, 8))
        self.status_box = ctk.CTkTextbox(right_frame, wrap="word", font=self.fonts["log"], state="disabled", fg_color=self.colors["gray-800"], border_color=self.colors["gray-600"], border_width=1, text_color=self.colors["gray-300"], corner_radius=8)
        self.status_box.grid(row=1, column=0, sticky="nsew")
        self.open_folder_btn = ctk.CTkButton(right_frame, text="Open Output Folder", state="disabled", command=self.open_output_folder, height=52, fg_color=self.colors["gray-700"], hover_color=self.colors["gray-600"], border_color=self.colors["gray-600"], border_width=1, font=self.fonts["button"])
        self.open_folder_btn.grid(row=2, column=0, sticky="ew", pady=(16, 10))
        
        # Footer
        footer = ctk.CTkFrame(self, fg_color=self.colors["gray-800"], corner_radius=0)
        footer.grid(row=2, column=0, sticky="ew")
        ctk.CTkLabel(footer, text="© 2025 Midmark Corporation. All rights reserved.", text_color=self.colors["gray-500"], font=self.fonts["footer"]).pack(pady=10)

    def load_initial_paths(self):
        # ... (this function is unchanged) ...
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
        # ... (this function is unchanged) ...
        initial_dir = self.template_path.parent if self.template_path and self.template_path.exists() else Path.home()
        file_path_str = filedialog.askopenfilename(title="Select Word Template File", filetypes=[("Word Templates", "*.dotm")], initialdir=initial_dir)
        if file_path_str:
            self.template_path = Path(file_path_str)
            self.template_path_label.configure(text=str(self.template_path), text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(template_path=self.template_path)

    def open_template_file(self):
        if self.template_path and self.template_path.is_file():
            threading.Thread(target=self._edit_template_worker, args=(self.template_path,), daemon=True).start()

    def open_excel_file(self):
        if self.excel_path and self.excel_path.is_file():
            os.startfile(self.excel_path)

    def select_excel_file(self):
        # ... (this function is unchanged) ...
        last_dir = self.excel_path.parent if self.excel_path and self.excel_path.exists() else Path.home()
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
        # ... (this function is unchanged) ...
        initial_dir = self.output_folder if self.output_folder and self.output_folder.exists() else Path.home()
        folder_path_str = filedialog.askdirectory(title="Select Output Folder", initialdir=initial_dir)
        if folder_path_str:
            self.output_folder = Path(folder_path_str)
            self.output_folder_label.configure(text=str(self.output_folder), text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(output_dir=self.output_folder)

    def update_generate_button_state(self):
        # ... (this function is unchanged) ...
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
        self.progress_bar.grid(row=6, column=0, sticky="ew", pady=(4, 6))
        self.progress_bar.set(0)
        self.status_label.configure(text="Status: Initializing...", text_color=self.colors["gray-400"])
        self.status_box.configure(state="normal")
        self.status_box.delete("1.0", "end")

        ai_enabled = self.ai_switch_var.get()
        intended_use_statement = self.intended_use_textbox.get("1.0", "end-1c")
        
        # --- MODIFICATION START ---
        # --- Gather additional prompts from the UI ---
        additional_prompts = []
        if self.prompt_cb_var_1.get() and self.prompt_entry_1.get():
            additional_prompts.append(("Target Audience", self.prompt_entry_1.get()))
        if self.prompt_cb_var_2.get() and self.prompt_entry_2.get():
            additional_prompts.append(("Tone of Voice", self.prompt_entry_2.get()))
        if self.prompt_cb_var_3.get() and self.prompt_entry_3.get():
            additional_prompts.append(("Keywords to Include", self.prompt_entry_3.get()))

        # Pass the new list to the worker thread
        self.worker_thread = threading.Thread(
            target=document_generation_worker, 
            args=(self.excel_path, self.output_folder, self.template_path, self.progress_queue, ai_enabled, intended_use_statement, additional_prompts),
            daemon=True
        )
        # --- MODIFICATION END ---
        self.worker_thread.start()

    def poll_queue(self):
        # ... (this function is unchanged) ...
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
        # ... (this function is unchanged) ...
        self.progress_bar.grid_remove()
        self.generate_btn.configure(state="normal")
        self.status_box.configure(state="disabled")
        if finished:
            self.open_folder_btn.configure(state="normal")


if __name__ == "__main__":
    if sys.platform == "win32":
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    
    app = DocCreatorApp()
    app.mainloop()