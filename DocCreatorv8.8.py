# Filename: DocCreatorv9.5.py 
# Function: Word Document Generator
# Author: Trinidad Hernandez
# Description: Finalized bulk doc generator with two AI modes: a simple all-or-nothing generation
#              and an advanced selective mode for per-section AI content generation with filtering.

import os
import json
import re
import sys
import threading
import queue
import pythoncom
import pandas as pd
import customtkinter as ctk
from tkinter import ttk 
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

# --- AI CONTENT GENERATION ---
def generate_ai_sectional_content(intended_use, document_title, section_header, additional_prompts, model_name):
    """
    Calls the Gemini API to generate new content for a specific section header.
    """
    try:
        api_key = "AIzaSyDf3cgcfhAdktmHWLkWOIqhi8OlxzzTFR8" 
        if not api_key:
            return "Error: No Gemini API key found"
        
        genai.configure(api_key=api_key)
        
        model = genai.GenerativeModel(model_name) 
        
        prompt_parts = [
            "You are a technical writer creating a specific section for a document.",
            f"DOCUMENT TITLE: \"{document_title}\"",
            f"DOCUMENT'S INTENDED USE: \"{intended_use}\"",
            f"CURRENT SECTION HEADER: \"{section_header}\"\n"
        ]

        if additional_prompts:
            instruction_parts = ["Please adhere to the following additional guidelines:"]
            for label, text in additional_prompts:
                instruction_parts.append(f"- {label}: {text}")
            prompt_parts.append("\n".join(instruction_parts))

        prompt_parts.append(
            f"\nWrite the body text for the section: \"{section_header}\"."
            f"Rules: Output only the paragraph text. Do not repeat the header in your response. Ensure the content is unique and does not repeat sentences from other sections. The device model name should primarily be used in the opening sections. Use concise sentences and avoid special characters."
        )
        
        full_prompt = "\n".join(prompt_parts)
        response = model.generate_content(full_prompt)
        return response.text
    except Exception as e:
        print(f"Gemini API Error: {e}")
        return f"[Error generating AI content for this section: {e}]"


# --- WORD AUTOMATION LOGIC ---

def _replace_placeholder_in_range(doc_range, placeholder, replacement_text):
    """Helper function to perform a find-and-replace in a Word range."""
    find = doc_range.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Text = placeholder
    find.Replacement.Text = str(replacement_text)
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

def document_generation_worker(excel_path, output_folder, template_path, progress_queue, ai_enabled, intended_use_statement, additional_prompts, selected_model, selective_mode, selection_data):
    """
    Worker function to be run in a separate thread.
    Handles the core logic of reading the Excel file and generating Word documents.
    """
    pythoncom.CoInitialize()
    word = None
    try:
        df = pd.read_excel(excel_path).fillna('')
        total_docs = len(df)
        progress_queue.put(("set_max", total_docs))

        kill_word_processes()
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False

        filename_col = df.columns[0]
        success_count = 0
        heading_pattern = re.compile(r'^H(\d+.*)$')

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
                        if not heading_content: continue

                        level = col_header.count('-') + 1
                        word_style = f"Heading {level}"
                        progress_queue.put(("log", f"  -> Adding Header: '{heading_content[:40]}...'\n"))
                        
                        word.Selection.EndKey(Unit=win32.constants.wdStory)
                        word.Selection.Style = doc.Styles(word_style)
                        word.Selection.TypeText(Text=heading_content)
                        heading_indent = word.Selection.ParagraphFormat.LeftIndent
                        word.Selection.TypeParagraph()

                        use_ai_for_section = False
                        if selective_mode:
                            use_ai_for_section = selection_data.get(filename, {}).get(heading_content, False)
                        elif ai_enabled:
                            use_ai_for_section = True
                        
                        manual_text_content = str(row.get(f"{col_header}-text", "")).strip()

                        word.Selection.Style = doc.Styles("Normal")
                        word.Selection.ParagraphFormat.LeftIndent = heading_indent
                        
                        if use_ai_for_section:
                            if document_title and intended_use_statement.strip():
                                progress_queue.put(("log", f"    -> Calling Gemini for '{heading_content}' section...\n"))
                                ai_text = generate_ai_sectional_content(intended_use_statement, document_title, heading_content, additional_prompts, selected_model)
                                word.Selection.TypeText(Text=ai_text)
                            else:
                                progress_queue.put(("log", "    -> [SKIPPED AI] Missing 'Title' in Excel or Intended Use statement.\n"))
                        elif manual_text_content:
                            progress_queue.put(("log", f"    -> Adding body text from '{col_header}-text'\n"))
                            word.Selection.TypeText(Text=manual_text_content)
                        
                        word.Selection.TypeParagraph()
                
                if doc.Sections.Count > 0:
                    for section in doc.Sections:
                        for header in section.Headers:
                            for col_name, cell_value in row.items():
                                placeholder = f"<{col_name}>"
                                _replace_placeholder_in_range(header.Range, placeholder, cell_value)
                        for footer in section.Footers:
                             for col_name, cell_value in row.items():
                                placeholder = f"<{col_name}>"
                                _replace_placeholder_in_range(footer.Range, placeholder, cell_value)
                
                for col_name, cell_value in row.items():
                    placeholder = f"<{col_name}>"
                    _replace_placeholder_in_range(doc.Content, placeholder, cell_value)

                if doc.Paragraphs.Count > 0 and doc.Paragraphs(1).Range.Text.strip() == "":
                    doc.Paragraphs(1).Range.Delete()

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
        if word: word.Quit()
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
        self.selection_data = {}
        self.tree_item_map = {}
        
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
        width = int(screen_width * 0.6)
        height = int(screen_height * 0.85)
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.geometry(f"{width}x{height}+{x}+{y}")

        self.setup_ui()
        self.load_initial_paths()
        self._toggle_ui_mode()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(100, self.poll_queue)

    def setup_ui(self):
        self.colors = { "gray-900": "#111827", "gray-800": "#1F2937", "gray-700": "#374151", "gray-600": "#4B5563", "gray-500": "#6B7280", "gray-400": "#9CA3AF", "gray-300": "#D1D5DB", "gray-200": "#E5E7EB", "blue-600": "#2563EB", "blue-700": "#1D4ED8", "green-400": "#4ADE80", }
        self.title("Midmark Batch Document Generator")
        self.minsize(900, 900)
        self.configure(fg_color=self.colors["gray-900"])

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color=self.colors["gray-800"], corner_radius=0, height=60)
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(header, text="Batch Document Generator", font=self.fonts["header"], text_color=self.colors["gray-200"]).pack(side="left", padx=20, pady=10)

        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=6)
        main_container.grid_columnconfigure(0, weight=4)
        main_container.grid_columnconfigure(1, weight=3)
        main_container.grid_rowconfigure(0, weight=1)

        left_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(10, 5))
        left_frame.grid_columnconfigure(0, weight=1)

        right_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 10))
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(2, weight=1)
        
        step1_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8,border_color=self.colors["gray-600"], border_width=1)
        step1_card.grid(row=0, column=0, sticky="ew", pady=(10, 6))
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
        ctk.CTkLabel(step2_card, text="Step 2: Select Excel File", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, padx=24, pady=(20, 16), sticky="ew")
        step2_inner = ctk.CTkFrame(step2_card, fg_color="transparent")
        step2_inner.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="ew")
        step2_inner.grid_columnconfigure(0, weight=1)
        self.excel_path_label = ctk.CTkLabel(step2_inner, text="No file selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.excel_path_label.grid(row=0, column=0, sticky="ew")
        self.edit_excel_btn = ctk.CTkButton(step2_inner, text="Edit", command=self.open_excel_file, width=70, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"])
        self.edit_excel_btn.grid(row=0, column=1, padx=(16, 8))
        ctk.CTkButton(step2_inner, text="Browse...", command=self.select_excel_file, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=2)

        ai_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8, border_color=self.colors["gray-600"], border_width=1)
        ai_card.grid(row=2, column=0, sticky="ew", pady=6)
        ai_card.grid_columnconfigure(0, weight=1)
        
        mode_header_frame = ctk.CTkFrame(ai_card, fg_color="transparent")
        mode_header_frame.grid(row=0, column=0, padx=24, pady=(20,10), sticky="ew")
        mode_header_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(mode_header_frame, text="Step 3: AI Content Generation", font=self.fonts["card_title"], text_color=self.colors["gray-300"]).grid(row=0, column=0, rowspan=2, sticky="w")
        
        self.ai_enable_label = ctk.CTkLabel(mode_header_frame, text="Enable AI", font=self.fonts["button"])
        self.ai_enable_label.grid(row=0, column=1, sticky="e", padx=(10, 5))
        self.ai_switch_var = ctk.BooleanVar(value=False)
        self.ai_switch = ctk.CTkSwitch(mode_header_frame, text="", variable=self.ai_switch_var)
        self.ai_switch.grid(row=0, column=2, sticky="w")

        self.selective_mode_label = ctk.CTkLabel(mode_header_frame, text="Selective Mode", font=self.fonts["button"])
        self.selective_mode_label.grid(row=1, column=1, sticky="e", padx=(10, 5))
        self.selective_mode_var = ctk.BooleanVar(value=False)
        self.selective_mode_switch = ctk.CTkSwitch(mode_header_frame, text="", variable=self.selective_mode_var, command=self._toggle_ui_mode)
        self.selective_mode_switch.grid(row=1, column=2, sticky="w")
        
        model_options = [ 'gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-live-2.5-flash-preview', 'gemini-2.5-flash-preview-native-audio-dialog', 'gemini-2.5-flash-preview-tts', 'gemini-2.5-pro-preview-tts', 'gemini-2.0-flash', 'gemini-2.0-flash-lite', 'gemini-2.0-flash-live-001', 'gemini-2.0-flash-preview-image-generation', 'gemini-1.5-pro', 'gemini-1.5-flash', 'gemini-1.5-flash-8b' ]
        model_frame = ctk.CTkFrame(ai_card, fg_color="transparent")
        model_frame.grid(row=1, column=0, padx=24, pady=(0, 10), sticky="ew")
        model_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(model_frame, text="AI Model", font=self.fonts["button"], text_color=self.colors["gray-400"]).grid(row=0, column=0, sticky="w")
        self.model_selection_combo = ctk.CTkComboBox(model_frame, values=model_options, font=self.fonts["log"])
        self.model_selection_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        self.model_selection_combo.set('gemini-2.5-flash-lite')
        
        ctk.CTkLabel(ai_card, text="Intended Use Statement", font=self.fonts["button"], text_color=self.colors["gray-400"]).grid(row=2, column=0, padx=24, pady=(0,4), sticky="w")
        self.intended_use_textbox = ctk.CTkTextbox(ai_card, height=60, wrap="word", font=self.fonts["log"], fg_color=self.colors["gray-700"], text_color=self.colors["gray-200"], border_width=0, corner_radius=6)
        self.intended_use_textbox.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 12))
        self.intended_use_textbox.insert("1.0", "Example: This document is for trained technicians...")
        
        additional_prompts_frame = ctk.CTkFrame(ai_card, fg_color="transparent")
        additional_prompts_frame.grid(row=4, column=0, sticky="ew", padx=24, pady=(0, 24))
        additional_prompts_frame.grid_columnconfigure(1, weight=1)
        audience_options = [ "Field Service Technicians", "Biomedical Engineers (Biomeds)", "Clinical Staff (Nurses, Assistants)", "Physicians / Doctors", "Hospital Administrators", "Regulatory Auditors", "Internal Engineers (R&D, Manufacturing)", "Sales & Marketing Teams" ]
        tone_options = [ "Professional", "Technical", "Instructional", "Informative", "Authoritative", "Formal (Regulatory)" ]
        self.prompt_cb_var_1 = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(additional_prompts_frame, text="Target Audience", variable=self.prompt_cb_var_1, font=self.fonts["button"]).grid(row=0, column=0, sticky="w")
        self.prompt_entry_1 = ctk.CTkComboBox(additional_prompts_frame, values=audience_options, font=self.fonts["log"])
        self.prompt_entry_1.grid(row=0, column=1, sticky="ew", padx=(10,0))
        self.prompt_entry_1.set(audience_options[0])
        self.prompt_cb_var_2 = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(additional_prompts_frame, text="Tone of Voice", variable=self.prompt_cb_var_2, font=self.fonts["button"]).grid(row=1, column=0, sticky="w", pady=(8,0))
        self.prompt_entry_2 = ctk.CTkComboBox(additional_prompts_frame, values=tone_options, font=self.fonts["log"])
        self.prompt_entry_2.grid(row=1, column=1, sticky="ew", padx=(10,0), pady=(8,0))
        self.prompt_entry_2.set(tone_options[0])
        self.prompt_cb_var_3 = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(additional_prompts_frame, text="Include Keywords", variable=self.prompt_cb_var_3, font=self.fonts["button"]).grid(row=2, column=0, sticky="w", pady=(8,0))
        self.prompt_entry_3 = ctk.CTkEntry(additional_prompts_frame, placeholder_text="e.g., sterilization, compliance, maintenance", font=self.fonts["log"])
        self.prompt_entry_3.grid(row=2, column=1, sticky="ew", padx=(10,0), pady=(8,0))
        
        step4_card = ctk.CTkFrame(left_frame, fg_color=self.colors["gray-800"], corner_radius=8, border_color=self.colors["gray-600"], border_width=1)
        step4_card.grid(row=3, column=0, sticky="ew", pady=6)
        ctk.CTkLabel(step4_card, text="Step 4: Select Output Folder", font=self.fonts["card_title"], text_color=self.colors["gray-300"], anchor="w").grid(row=0, column=0, padx=24, pady=(20, 16), sticky="ew")
        step4_inner = ctk.CTkFrame(step4_card, fg_color="transparent")
        step4_inner.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="ew")
        step4_inner.grid_columnconfigure(0, weight=1)
        self.output_folder_label = ctk.CTkLabel(step4_inner, text="No folder selected", fg_color=self.colors["gray-700"], text_color=self.colors["gray-400"], corner_radius=6, anchor="w", padx=12, height=40)
        self.output_folder_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(step4_inner, text="Browse...", command=self.select_output_folder, width=120, height=40, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"]).grid(row=0, column=1, padx=(16, 0))

        self.generate_btn = ctk.CTkButton(left_frame, text="Generate Documents", state="disabled", command=self.start_generation, height=52, fg_color=self.colors["blue-600"], hover_color=self.colors["blue-700"], font=self.fonts["card_title"])
        self.generate_btn.grid(row=4, column=0,  sticky="ew", pady=(12, 6))
        self.progress_bar = ctk.CTkProgressBar(left_frame, mode="determinate", height=16, fg_color=self.colors["gray-700"], progress_color=self.colors["green-400"])
        self.status_label = ctk.CTkLabel(left_frame, text="Status: Ready", font=self.fonts["status"], text_color=self.colors["green-400"])
        self.status_label.grid(row=5, column=0, pady=4, sticky="w")
        
        right_header_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        right_header_frame.grid(row=0, column=0, sticky="ew", pady=(10, 8))
        right_header_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(right_header_frame, text="Status Log / Section Selector", font=self.fonts["card_title"], text_color=self.colors["gray-300"]).grid(row=0, column=0, sticky="w")
        self.scan_btn = ctk.CTkButton(right_header_frame, text="Refresh Sections", width=120, height=30, state="disabled", command=self.scan_excel_file, fg_color=self.colors["gray-600"], hover_color=self.colors["gray-700"], font=self.fonts["button"])
        self.scan_btn.grid(row=0, column=1, sticky="e")
        
        self.filter_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        self.filter_frame.grid(row=1, column=0, sticky="ew", pady=(0, 5))
        self.filter_var1 = ctk.BooleanVar(value=True)
        self.filter_var2 = ctk.BooleanVar(value=True)
        self.filter_var3 = ctk.BooleanVar(value=True)
        ctk.CTkLabel(self.filter_frame, text="Filter:", font=self.fonts["button"]).pack(side="left", padx=(0, 5))
        ctk.CTkCheckBox(self.filter_frame, text="Level 1 (Hn)", variable=self.filter_var1, font=self.fonts["button"], command=self._apply_tree_filter).pack(side="left", padx=5)
        ctk.CTkCheckBox(self.filter_frame, text="Level 2 (Hn-n)", variable=self.filter_var2, font=self.fonts["button"], command=self._apply_tree_filter).pack(side="left", padx=5)
        ctk.CTkCheckBox(self.filter_frame, text="Level 3+ (Hn-..)", variable=self.filter_var3, font=self.fonts["button"], command=self._apply_tree_filter).pack(side="left", padx=5)
        self.status_box = ctk.CTkTextbox(right_frame, wrap="word", font=self.fonts["log"], state="disabled", fg_color=self.colors["gray-800"], border_color=self.colors["gray-600"], border_width=1)
        self.status_box.grid(row=2, column=0, sticky="nsew")
        self.tree_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        self.tree_frame.grid(row=2, column=0, sticky="nsew")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background=self.colors["gray-800"], foreground=self.colors["gray-200"], fieldbackground=self.colors["gray-800"], borderwidth=0, font=self.fonts["log"])
        style.map('Treeview', background=[('selected', self.colors["blue-700"])])
        style.configure("Treeview.Heading", background=self.colors["gray-700"], foreground=self.colors["gray-200"], font=self.fonts["button"])
        self.selection_tree = ttk.Treeview(self.tree_frame, show="tree")
        self.selection_tree.heading("#0", text="Documents and Sections")
        self.selection_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll = ctk.CTkScrollbar(self.tree_frame, command=self.selection_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.selection_tree.configure(yscrollcommand=tree_scroll.set)
        self.selection_tree.bind("<ButtonRelease-1>", self._toggle_selection)
        self.open_folder_btn = ctk.CTkButton(right_frame, text="Open Output Folder", state="disabled", command=self.open_output_folder, height=52, fg_color=self.colors["gray-700"], hover_color=self.colors["gray-600"], font=self.fonts["button"])
        self.open_folder_btn.grid(row=3, column=0, sticky="ew", pady=(16, 10))
    
    def _toggle_ui_mode(self):
        is_selective = self.selective_mode_var.get()
        if is_selective:
            self.ai_switch_var.set(True)
            self.ai_switch.configure(state="disabled")
            self.ai_enable_label.configure(state="disabled", text_color=self.colors["gray-600"])
            self.scan_btn.grid()
            self.status_box.grid_remove()
            self.tree_frame.grid()
            self.filter_frame.grid()
            if self.excel_path:
                self.scan_excel_file()
        else:
            self.ai_switch.configure(state="normal")
            self.ai_enable_label.configure(state="normal", text_color=self.colors["gray-400"])
            self.scan_btn.grid_remove()
            self.tree_frame.grid_remove()
            self.filter_frame.grid_remove()
            self.status_box.grid()
        self.update_generate_button_state()
    
    def scan_excel_file(self):
        if not self.excel_path: return
        self.selection_tree.delete(*self.selection_tree.get_children())
        self.selection_data.clear()
        self.tree_item_map.clear()
        try:
            df = pd.read_excel(self.excel_path).fillna('')
            filename_col = df.columns[0]
            heading_pattern = re.compile(r'^H(\d+.*)$')
            for i, row in df.iterrows():
                filename = str(row[filename_col]).strip()
                if not filename: continue
                parent_id = self.selection_tree.insert("", "end", text=f"{filename}", open=True)
                self.selection_data[filename] = {}
                for col_header in df.columns:
                    if heading_pattern.match(col_header):
                        heading_content = str(row.get(col_header, "")).strip()
                        if not heading_content: continue
                        level = col_header.count('-') + 1
                        indent_prefix = "    " * (level - 1)
                        item_id = self.selection_tree.insert(parent_id, "end", text=f"{indent_prefix}[  ] {heading_content}")
                        self.tree_item_map[item_id] = (filename, heading_content, level, parent_id)
                        self.selection_data[filename][heading_content] = False
            self.generate_btn.configure(state="normal")
            self._apply_tree_filter()
        except Exception as e:
            messagebox.showerror("Excel Scan Error", f"Could not read the Excel file:\n{e}")
            self.generate_btn.configure(state="disabled")

    def _toggle_selection(self, event):
        item_id = self.selection_tree.identify_row(event.y)
        if not item_id or item_id not in self.tree_item_map: return
        filename, section_name, level, parent_id = self.tree_item_map[item_id]
        current_state = self.selection_data[filename][section_name]
        new_state = not current_state
        self.selection_data[filename][section_name] = new_state
        indent_prefix = "    " * (level - 1)
        new_text = f"{indent_prefix}[{'✓' if new_state else ' '}] {section_name}"
        self.selection_tree.item(item_id, text=new_text)

    def _apply_tree_filter(self):
        show_level1 = self.filter_var1.get()
        show_level2 = self.filter_var2.get()
        show_level3 = self.filter_var3.get()
        for item_id, data in self.tree_item_map.items():
            self.selection_tree.detach(item_id)
            filename, section_name, level, parent_id = data
            is_selected = self.selection_data[filename][section_name]
            indent_prefix = "    " * (level - 1)
            current_text = f"{indent_prefix}[{'✓' if is_selected else ' '}] {section_name}"
            self.selection_tree.item(item_id, text=current_text)
            if (level == 1 and show_level1) or (level == 2 and show_level2) or (level >= 3 and show_level3):
                self.selection_tree.move(item_id, parent_id, "end")

    def start_generation(self):
        self.progress_bar.grid(row=5, column=0, sticky="ew", pady=(4, 6))
        self.progress_bar.set(0)
        if self.selective_mode_var.get():
            self.tree_frame.grid_remove()
            self.filter_frame.grid_remove()
            self.status_box.grid()
        self.status_label.configure(text="Status: Initializing...", text_color=self.colors["gray-400"])
        self.status_box.configure(state="normal")
        self.status_box.delete("1.0", "end")
        intended_use = self.intended_use_textbox.get("1.0", "end-1c")
        selected_model = self.model_selection_combo.get()
        additional_prompts = []
        if self.prompt_cb_var_1.get() and self.prompt_entry_1.get():
            additional_prompts.append(("Target Audience", self.prompt_entry_1.get()))
        if self.prompt_cb_var_2.get() and self.prompt_entry_2.get():
            additional_prompts.append(("Tone of Voice", self.prompt_entry_2.get()))
        if self.prompt_cb_var_3.get() and self.prompt_entry_3.get():
            additional_prompts.append(("Keywords to Include", self.prompt_entry_3.get()))
        self.worker_thread = threading.Thread(
            target=document_generation_worker, 
            args=(self.excel_path, self.output_folder, self.template_path, 
                  self.progress_queue, self.ai_switch_var.get(), intended_use, 
                  additional_prompts, selected_model, self.selective_mode_var.get(),
                  self.selection_data),
            daemon=True
        )
        self.worker_thread.start()

    def poll_queue(self):
        try:
            message = self.progress_queue.get_nowait()
            msg_type, *payload = message
            if msg_type == "error":
                messagebox.showerror("Error", payload[0])
                self.reset_ui()
            elif msg_type == "set_max": self.total_docs = payload[0]
            elif msg_type == "progress":
                if self.total_docs > 0:
                    progress_float = payload[0] / self.total_docs
                    self.progress_bar.set(progress_float)
                    self.status_label.configure(text=f"Status: In progress... {int(progress_float*100)}%")
            elif msg_type == "log":
                self.status_box.insert("end", payload[0])
                self.status_box.see("end")
            elif msg_type == "done":
                success, total = payload
                self.status_label.configure(text=f"Status: Finished. {success}/{total} created.", text_color=self.colors["green-400"])
                self.reset_ui(finished=True)
        except queue.Empty: pass
        finally: self.after(100, self.poll_queue)

    def reset_ui(self, finished=False):
        self.progress_bar.grid_remove()
        self.update_generate_button_state()
        self.status_box.configure(state="disabled")
        if finished: self.open_folder_btn.configure(state="normal")
    
    def load_initial_paths(self):
        last_locations = load_last_locations()
        last_template = last_locations.get("template_path")
        if last_template and Path(last_template).is_file(): self.template_path = Path(last_template)
        else: self.template_path = self._find_normal_template()
        if self.template_path: self.template_path_label.configure(text=str(self.template_path), text_color=self.colors["gray-200"])
        last_excel = last_locations.get("excel_path")
        if last_excel and Path(last_excel).is_file():
            self.excel_path = Path(last_excel)
            self.excel_path_label.configure(text=self.excel_path.name, text_color=self.colors["gray-200"])
        last_output = last_locations.get("output_dir")
        if last_output and Path(last_output).is_dir():
            self.output_folder = Path(last_output)
            self.output_folder_label.configure(text=str(self.output_folder), text_color=self.colors["gray-200"])
        self.update_generate_button_state()
    
    def open_template_file(self):
        if self.template_path and self.template_path.is_file():
            threading.Thread(target=self._edit_template_worker, args=(self.template_path,), daemon=True).start()

    def open_excel_file(self):
        if self.excel_path and self.excel_path.is_file():
            os.startfile(self.excel_path)

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
            
    def _show_com_error(self, error_exception):
        messagebox.showerror("Word Error", f"Could not open Word to edit template:\n{error_exception}")
    
    def select_template_file(self):
        file_path_str = filedialog.askopenfilename(title="Select Word Template", filetypes=[("Word Templates", "*.dotm")])
        if file_path_str:
            self.template_path = Path(file_path_str)
            self.template_path_label.configure(text=str(self.template_path), text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(template_path=self.template_path)

    def select_excel_file(self):
        file_path_str = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path_str:
            self.excel_path = Path(file_path_str)
            self.excel_path_label.configure(text=self.excel_path.name, text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(excel_path=self.excel_path)
    
    def select_output_folder(self):
        folder_path_str = filedialog.askdirectory(title="Select Output Folder")
        if folder_path_str:
            self.output_folder = Path(folder_path_str)
            self.output_folder_label.configure(text=str(self.output_folder), text_color=self.colors["gray-200"])
            self.update_generate_button_state()
            save_last_locations(output_dir=self.output_folder)
    
    def update_generate_button_state(self):
        files_selected = self.excel_path and self.output_folder and self.template_path
        self.edit_template_btn.configure(state="normal" if self.template_path else "disabled")
        self.edit_excel_btn.configure(state="normal" if self.excel_path else "disabled")
        self.scan_btn.configure(state="normal" if self.excel_path else "disabled")
        if self.selective_mode_var.get():
            is_scanned = len(self.selection_data) > 0
            self.generate_btn.configure(state="normal" if files_selected and is_scanned else "disabled")
        else:
            self.generate_btn.configure(state="normal" if files_selected else "disabled")
            
    def open_output_folder(self):
        if self.output_folder and self.output_folder.is_dir(): os.startfile(self.output_folder)
        
if __name__ == "__main__":
    app = DocCreatorApp()
    app.mainloop()