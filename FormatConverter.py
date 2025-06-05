import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import shutil
from pathlib import Path
from tkinterdnd2 import DND_FILES, TkinterDnD
from pdf2docx import Converter
import subprocess
import win32con
import win32process
from docx2pdf import convert
import win32com.client
import pythoncom
import sys

class FormatConverter:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("ShugeZhuanZhuan v1.0 - Developer: Fair Paws")
        
        try:
            if getattr(sys, 'frozen', False):
                # If it's a packaged exe
                base_path = sys._MEIPASS
            else:
                # If it's a python script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, "zhuan.ico")
            self.root.iconbitmap(icon_path)
        except:
            pass  # Don't raise an error if icon is not found
            
        self.root.geometry("500x450")

        # ======== Define Color Scheme ========
        self.BG_COLOR       = "#2B2B2B"  
        self.FG_COLOR       = "#DDDDDD"
        self.ACCENT_COLOR   = "#3C3F41"
        self.HIGHLIGHT_BAR  = "#4C5052"
        self.PROG_BAR_COLOR = "#569CD6"
        self.HINT_COLOR     = "#666666"

        # Dark background
        self.root.configure(bg=self.BG_COLOR)

        # ======== Configure ttk.Style ========
        style = ttk.Style(self.root)
        style.theme_use("default")
        style.configure('.', background=self.BG_COLOR, foreground=self.FG_COLOR)
        style.configure('TFrame', background=self.BG_COLOR)
        style.configure('TLabelFrame', background=self.BG_COLOR, foreground=self.FG_COLOR)
        style.configure('TLabel', background=self.BG_COLOR, foreground=self.FG_COLOR)
        style.configure('TButton', background=self.ACCENT_COLOR, foreground=self.FG_COLOR, relief='flat')
        style.map('TButton', background=[('active', self.HIGHLIGHT_BAR), ('pressed', self.ACCENT_COLOR)])
        style.configure('TCheckbutton', background=self.BG_COLOR, foreground=self.FG_COLOR)
        style.map('TCheckbutton', background=[('active', self.BG_COLOR), ('selected', self.BG_COLOR)])
        style.configure('TEntry',
                        fieldbackground=self.ACCENT_COLOR,
                        foreground="#666666",
                        insertcolor=self.FG_COLOR,
                        readonlybackground="#444444",
                        readonlyforeground="#666666",
                        disabledforeground="#777777")
        style.configure('TProgressbar', troughcolor=self.ACCENT_COLOR, background=self.PROG_BAR_COLOR)

        # Default save location: Desktop
        self.save_path = os.path.join(os.path.expanduser("~"), "Desktop")
        
        # Target formats
        self.formats = ['PDF', 'DOCX', 'EPUB', 'AZW3', 'MOBI']
        self.selected_formats = {fmt: tk.BooleanVar() for fmt in self.formats}
        
        # Create GUI
        self.create_main_interface()
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_frame = ttk.Frame(self.root)
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100, style='TProgressbar')
        self.progress_label = ttk.Label(self.progress_frame, text="0%")

    def show_styled_messagebox(self, parent, title, message):
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.geometry("300x150")
        dialog.configure(bg=self.BG_COLOR)
        dialog.transient(parent)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog)
        frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        msg_label = ttk.Label(
            frame, 
            text=message,
            wraplength=250,
            justify='center'
        )
        msg_label.pack(expand=True, pady=10)
        
        btn = ttk.Button(frame, text="OK", command=dialog.destroy)
        btn.pack(pady=10)
        
        # Center the window
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        dialog.wait_window()

    def create_main_interface(self):
        list_frame = ttk.LabelFrame(self.root, text="File List")
        list_frame.pack(fill='x', expand=False, padx=5, pady=5)
        
        self.file_listbox = tk.Listbox(
            list_frame, 
            height=4, 
            bg=self.ACCENT_COLOR, 
            fg=self.FG_COLOR,
            highlightthickness=0,
            selectbackground=self.HIGHLIGHT_BAR,
            selectforeground=self.FG_COLOR
        )
        self.file_listbox.pack(fill='x', expand=False, padx=5, pady=5)
        self.file_listbox.insert(tk.END, "Drag files here or click 'Add Files'")
        
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.handle_drop)
        
        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill='x', padx=5, pady=2)
        ttk.Button(btn_frame, text="Add Files", command=self.add_files).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Delete Selected", command=self.remove_selected).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Clear List", command=self.clear_list).pack(side='left', padx=2)

        # Save location
        save_frame = ttk.LabelFrame(self.root, text="Save Location")
        save_frame.pack(fill='x', padx=5, pady=5)
        
        self.save_path_var = tk.StringVar(value=self.save_path)
        save_entry = ttk.Entry(save_frame, textvariable=self.save_path_var, state='readonly')
        save_entry.pack(side='left', fill='x', expand=True, padx=(5,2), pady=5)
        ttk.Button(save_frame, text="Browse", command=self.browse_save_path).pack(side='right', padx=5, pady=5)
        
        # Target formats
        format_frame = ttk.LabelFrame(self.root, text="Target Format")
        format_frame.pack(fill='x', padx=5, pady=5)
        
        format_grid = ttk.Frame(format_frame)
        format_grid.pack(fill='x', padx=5, pady=5)
        for fmt in self.formats:
            ttk.Checkbutton(format_grid, text=fmt, variable=self.selected_formats[fmt]).pack(side='left', padx=5)
        
        # Friendly Reminder
        hint_frame = ttk.LabelFrame(self.root, text="Friendly Reminder")
        hint_frame.pack(fill='x', padx=5, pady=5)
        
        hint_text = ("Oh dear! Complex layouts or special fonts might lead to conversion failures.\n"
                     "It's recommended to try with simpler documents first before handling special formats!")
        
        hint_label = ttk.Label(hint_frame, text=hint_text, foreground=self.HINT_COLOR, wraplength=450)
        hint_label.pack(padx=5, pady=5)
        
        ttk.Button(self.root, text="Start Conversion", command=self.start_conversion).pack(fill='x', padx=5, pady=5)

    def browse_save_path(self):
        path = filedialog.askdirectory(
            initialdir=self.save_path,
            title="Choose a folder to save your files!"
        )
        if path:
            self.save_path = path
            self.save_path_var.set(path)

    def handle_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if os.path.isfile(file):
                if self.file_listbox.get(0) == "Drag files here or click 'Add Files'":
                    self.file_listbox.delete(0)
                self.file_listbox.insert(tk.END, file)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Select files to convert!",
            filetypes=[
                ("Supported Files", "*.pdf;*.docx;*.epub;*.mobi;*.azw3"),
                ("All Files", "*.*")
            ]
        )
        if files:
            if self.file_listbox.get(0) == "Drag files here or click 'Add Files'":
                self.file_listbox.delete(0)
            for file in files:
                self.file_listbox.insert(tk.END, file)

    def remove_selected(self):
        selection = self.file_listbox.curselection()
        for i in selection[::-1]:
            self.file_listbox.delete(i)

    def clear_list(self):
        self.file_listbox.delete(0, tk.END)
        self.file_listbox.insert(tk.END, "Drag files here or click 'Add Files'")

    def generate_unique_filename(self, base_path):
        if not os.path.exists(base_path):
            return base_path
        base, ext = os.path.splitext(base_path)
        counter = 1
        while True:
            new_file = f"{base}_copy{counter}{ext}"
            if not os.path.exists(new_file):
                return new_file
            counter += 1

    def is_tool_available(self, tool_name):
        return shutil.which(tool_name) is not None

    def convert_file(self, input_path, output_path):
        input_ext = Path(input_path).suffix.lower()
        output_ext = Path(output_path).suffix.lower()
        
        if input_ext == output_ext:
            return False

        try:
            pythoncom.CoInitialize()

            # PDF -> DOCX
            if input_ext == ".pdf" and output_ext == ".docx":
                try:
                    cv = Converter(input_path)
                    cv.convert(output_path)
                    cv.close()
                    return True
                except Exception as e:
                    raise Exception("PDF to Word conversion failed: " + str(e))

            # DOCX -> PDF
            if input_ext == ".docx" and output_ext == ".pdf":
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    try:
                        doc = word.Documents.Open(os.path.abspath(input_path))
                        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
                        doc.Close()
                        return True
                    finally:
                        word.Quit()
                except Exception as e:
                    raise Exception(f"Word to PDF conversion failed: {str(e)}")

            # Other e-book format conversions
            if input_ext in [".pdf", ".docx", ".epub", ".mobi", ".azw3"] and \
               output_ext in [".pdf", ".docx", ".epub", ".mobi", ".azw3"]:
                if not self.is_tool_available("ebook-convert"):
                    raise Exception("Calibre's ebook-convert not detected. Please install Calibre and configure PATH.")
                try:
                    cmd = ["ebook-convert", input_path, output_path]
                    # Use startupinfo to hide the window
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    subprocess.run(cmd, 
                                   check=True, 
                                   capture_output=True, 
                                   text=True,
                                   startupinfo=startupinfo)
                    return True
                except subprocess.CalledProcessError as e:
                    raise Exception(f"Calibre conversion failed: {e.stderr}")
            
            return False

        except Exception as e:
            raise Exception(f"Conversion failed: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def start_conversion(self):
        files = list(self.file_listbox.get(0, tk.END))
        if files and files[0] == "Drag files here or click 'Add Files'":
            files = files[1:]
        if not files:
            self.show_styled_messagebox(self.root, "Oops!", "No files selected for conversion yet!")
            return

        selected = [f for f,v in self.selected_formats.items() if v.get()]
        if not selected:
            self.show_styled_messagebox(self.root, "Oops!", "Please select at least one desired format!")
            return

        valid_files = []
        source_exts = set()
        for f in files:
            if os.path.isfile(f):
                ext = Path(f).suffix.lower()
                source_exts.add(ext)
                valid_files.append(f)

        if not valid_files:
            self.show_styled_messagebox(self.root, "Oops!", "No convertible files found!")
            return

        if len(source_exts) == 1:
            single_ext = list(source_exts)[0]
            for sf in selected:
                if f".{sf.lower()}" == single_ext:
                    self.show_styled_messagebox(self.root, "Oops!", f"The target format [{sf}] is the same as the original format. No conversion needed!")
                    return
        else:
            if len(selected) > 1:
                self.show_styled_messagebox(self.root, "Oops!", "For multiple files with different formats, please select only one target format!")
                return
            tgt = selected[0].lower()
            for se in source_exts:
                if f".{tgt}" == se:
                    self.show_styled_messagebox(self.root, "Oops!", f"The target format [{tgt}] is the same as some original file formats. No conversion needed!")
                    return

        total_tasks = len(valid_files) * len(selected)
        self.progress_var.set(0)
        self.progress_frame.pack(fill='x', padx=5, pady=5)
        self.progress_bar.pack(side='left', fill='x', expand=True)
        self.progress_label.pack(side='left', padx=5)

        def do_convert():
            success_count = 0
            fail_count = 0
            done_count = 0

            for f in valid_files:
                for fmt in selected:
                    out_path = self.generate_unique_filename(
                        os.path.join(self.save_path, Path(f).stem + '.' + fmt.lower())
                    )
                    try:
                        ret = self.convert_file(f, out_path)
                        if ret:
                            success_count += 1
                        else:
                            fail_count += 1
                    except Exception as e:
                        print(f"Conversion error: {e}")
                        fail_count += 1
                    done_count += 1
                    pct = (done_count / total_tasks) * 100
                    self.progress_var.set(pct)
                    self.progress_label.config(text=f"{int(pct)}%")

            self.progress_var.set(100)
            self.progress_label.config(text="100%")
            self.show_styled_messagebox(self.root, "Done!", f"Conversion results are in!\nSuccessful: {success_count} files\nFailed: {fail_count} files")
            self.progress_frame.pack_forget()

        threading.Thread(target=do_convert, daemon=True).start()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = FormatConverter()
    app.run()