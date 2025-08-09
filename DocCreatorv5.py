# DocCreatorv5.py - Midmark Excel to Word document generator with template header/footer replacement
#
# V11 - ADDED CUSTOM ICON: The script now looks for 'MidmarkTLogo.ico' in its
# directory and sets it as the application window icon.
#
# uploaded to github.com
# testing push
import os
import json
import re
import sys
import threading
import queue
import pythoncom
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import win32com.client as win32
import time


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

def save_last_locations(excel_path=None, output_dir=None):
    """Saves last used locations to the JSON config file."""
    try:
        data = load_last_locations()
        if excel_path:
            data["excel_path"] = str(excel_path)
        if output_dir:
            data["output_dir"] = str(output_dir)
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

def document_generation_worker(excel_path, output_folder, progress_queue):
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
                doc = word.Documents.Add(Template=word.NormalTemplate.FullName)
                
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
        self.progress_queue = queue.Queue()
        self.total_docs = 0
        
        # --- Theme and Appearance ---
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.setup_ui()
        self.load_initial_paths()

    def setup_ui(self):
        self.title("Midmark Excel to Word Document Generator")
        self.geometry("550x600")
        self.minsize(550, 600)

        # --- Set Window Icon ---
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, "MidmarkTLogo.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        # --- Main Frame ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # --- Input Frame ---
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="Step 1: Select Excel File", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 5))
        self.excel_path_label = ctk.CTkLabel(input_frame, text="No file selected", text_color="gray70", wraplength=350)
        self.excel_path_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5)
        ctk.CTkButton(input_frame, text="Browse...", command=self.select_excel_file, width=100).grid(row=1, column=2, padx=5)
        
        ctk.CTkLabel(input_frame, text="Step 2: Select Output Folder", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, columnspan=3, sticky="w", pady=(15, 5))
        self.output_folder_label = ctk.CTkLabel(input_frame, text="No folder selected", text_color="gray70", wraplength=350)
        self.output_folder_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=5)
        ctk.CTkButton(input_frame, text="Browse...", command=self.select_output_folder, width=100).grid(row=3, column=2, padx=5)

        # --- Action Button ---
        self.generate_btn = ctk.CTkButton(self, text="Generate Documents", state="disabled", command=self.start_generation, font=ctk.CTkFont(size=14))
        self.generate_btn.grid(row=1, column=0, padx=15, pady=10, sticky="ew")

        # --- Progress Bar & Status ---
        self.progress_bar = ctk.CTkProgressBar(self, mode="determinate")
        self.progress_bar.set(0)
        self.progress_bar.grid(row=2, column=0, padx=15, pady=5, sticky="ew")
        self.progress_bar.grid_remove()
        
        self.status_label = ctk.CTkLabel(self, text="Status: Ready", text_color="gray70")
        self.status_label.grid(row=3, column=0, padx=15, pady=(0,5), sticky="ew")

        # --- Status Log ---
        self.status_box = ctk.CTkTextbox(self, height=200, wrap="word", font=("Courier New", 12), state="disabled")
        self.status_box.grid(row=4, column=0, padx=15, pady=(0,10), sticky="nsew")
        self.grid_rowconfigure(4, weight=1)

        # --- Bottom Frame ---
        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.grid(row=5, column=0, padx=15, pady=(0,10), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)

        self.open_folder_btn = ctk.CTkButton(bottom_frame, text="Open Output Folder", state="disabled", command=self.open_output_folder)
        self.open_folder_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(bottom_frame, text="©2025 Midmark Corporation. All rights reserved.", text_color="gray50", font=ctk.CTkFont(size=10)).grid(row=1, column=0, columnspan=2, sticky="s", pady=(5, 0))

    def load_initial_paths(self):
        last_locations = load_last_locations()
        last_excel = last_locations.get("excel_path")
        if last_excel and Path(last_excel).is_file():
            self.excel_path = Path(last_excel)
            self.excel_path_label.configure(text=self.excel_path.name)
        last_output = last_locations.get("output_dir")
        if last_output and Path(last_output).is_dir():
            self.output_folder = Path(last_output)
            self.output_folder_label.configure(text=str(self.output_folder))
        self.update_generate_button_state()

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
        self.excel_path_label.configure(text=self.excel_path.name)
        self.update_generate_button_state()
        save_last_locations(excel_path=self.excel_path)

    def select_output_folder(self):
        initial_dir = self.output_folder if self.output_folder else Path.home()
        folder_path_str = filedialog.askdirectory(title="Select Output Folder", initialdir=initial_dir)
        if folder_path_str:
            self.output_folder = Path(folder_path_str)
            self.output_folder_label.configure(text=str(self.output_folder))
            self.update_generate_button_state()
            save_last_locations(output_dir=self.output_folder)

    def update_generate_button_state(self):
        state = "normal" if self.excel_path and self.output_folder else "disabled"
        self.generate_btn.configure(state=state)

    def open_output_folder(self):
        if self.output_folder and self.output_folder.is_dir():
            os.startfile(self.output_folder)

    def start_generation(self):
        self.generate_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.progress_bar.grid()
        self.progress_bar.set(0)
        self.status_label.configure(text="🔄 Initializing...")
        self.status_box.configure(state="normal")
        self.status_box.delete("1.0", "end")
        self.worker_thread = threading.Thread(target=document_generation_worker, args=(self.excel_path, self.output_folder, self.progress_queue))
        self.worker_thread.start()
        self.after(100, self.poll_queue)

    def poll_queue(self):
        try:
            message = self.progress_queue.get_nowait()
            msg_type, *payload = message
            if msg_type == "error":
                messagebox.showerror("Error", payload[0])
                self.reset_ui()
            elif msg_type == "set_max":
                self.total_docs = payload[0]
                self.status_label.configure(text="🔄 Generating documents...")
            elif msg_type == "progress":
                if self.total_docs > 0:
                    progress_float = payload[0] / self.total_docs
                    self.progress_bar.set(progress_float)
                    percent = int(progress_float * 100)
                    self.status_label.configure(text=f"🔄 In progress... {percent}%")
            elif msg_type == "log":
                self.status_box.insert("end", payload[0])
                self.status_box.see("end")
            elif msg_type == "done":
                success_count, total_docs = payload
                self.status_label.configure(text=f"✅ Finished. {success_count}/{total_docs} documents created.")
                self.reset_ui(finished=True)
            self.after(100, self.poll_queue) 
        except queue.Empty:
            if self.worker_thread.is_alive():
                self.after(100, self.poll_queue)
            else: 
                self.reset_ui()

    def reset_ui(self, finished=False):
        self.progress_bar.after(1000, self.progress_bar.grid_remove)
        self.generate_btn.configure(state="normal")
        self.status_box.configure(state="disabled")
        if finished:
            self.open_folder_btn.configure(state="normal")

if __name__ == "__main__":
    app = DocCreatorApp()
    app.mainloop()