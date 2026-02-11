#I added comments in case something goes wrong or you want to edit it :)
import os
import sys
import threading
import subprocess
from pathlib import Path
from tkinter import filedialog, messagebox

def bootstrap():
    libs = ['customtkinter', 'pywin32', 'Pillow', 'PyMuPDF', 'tkinterdnd2']
    for lib in libs:
        try:
            if lib == 'pywin32': __import__('win32print')
            elif lib == 'PyMuPDF': __import__('fitz')
            else: __import__(lib.replace('2', ''))
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

bootstrap()

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32print
import win32ui
import win32con
from PIL import Image, ImageWin
import fitz

# --- Localization System ---
class i18n:
    DATA = {
        "EN": {
            "title": "Universal Print Station Pro",
            "queue": "PRINT QUEUE",
            "drop_hint": "Drop files here or click Add",
            "preview": "Preview",
            "params": "PARAMETERS",
            "printer": "Target Printer:",
            "paper": "Media Size:",
            "copies": "Copy Count:",
            "fit": "Scale to fit page",
            "print_btn": "EXECUTE PRINT",
            "lang": "Language",
            "ready": "Ready"
        },
        "RU": {
            "title": "Универсальная Станция Печати",
            "queue": "ОЧЕРЕДЬ ПЕЧАТИ",
            "drop_hint": "Перетащите файлы сюда или Добавьте",
            "preview": "Предпросмотр",
            "params": "ПАРАМЕТРЫ",
            "printer": "Принтер:",
            "paper": "Размер бумаги:",
            "copies": "Количество копий:",
            "fit": "Масштабировать в лист",
            "print_btn": "ЗАПУСТИТЬ ПЕЧАТЬ",
            "lang": "Язык",
            "ready": "Готов"
        }
    }

# --- Core Printing Logic ---
class PrinterService:
    @staticmethod
    def get_available_printers():
        return [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]

    @staticmethod
    def process_print(file_path, config):
        printer_name = config['printer']
        copies = int(config.get('copies', 1))
        
        # Low-level Windows Device Context setup
        h_printer = win32print.OpenPrinter(printer_name)
        try:
            pd = win32print.GetPrinter(h_printer, 2)
            pd["pDevMode"].Copies = copies
        finally:
            win32print.ClosePrinter(h_printer)

        ext = Path(file_path).suffix.lower()
        if ext == '.pdf':
            PrinterService._print_pdf(file_path, printer_name)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
            PrinterService._print_image(file_path, printer_name)

    @staticmethod
    def _print_image(path, printer):
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer)
        hdc.StartDoc(path)
        hdc.StartPage()
        
        img = Image.open(path)
        printable_w = hdc.GetDeviceCaps(win32con.HORZRES)
        printable_h = hdc.GetDeviceCaps(win32con.VERTRES)
        img.thumbnail((printable_w, printable_h), Image.Resampling.LANCZOS)
        
        dib = ImageWin.Dib(img)
        dib.draw(hdc.GetHandleOutput(), (0, 0, img.size[0], img.size[1]))
        
        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

    @staticmethod
    def _print_pdf(path, printer):
        pdf_doc = fitz.open(path)
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer)
        hdc.StartDoc(path)
        
        for page in pdf_doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            hdc.StartPage()
            ImageWin.Dib(img).draw(hdc.GetHandleOutput(), (0, 0, pix.width, pix.height))
            hdc.EndPage()
            
        hdc.EndDoc()
        hdc.DeleteDC()

# --- UI Components ---
class FileListItem(ctk.CTkFrame):
    def __init__(self, master, file_path, on_remove, on_preview):
        super().__init__(master, fg_color="transparent")
        self.file_path = file_path
        self.is_active = ctk.BooleanVar(value=True)
        
        self.cb = ctk.CTkCheckBox(self, text="", variable=self.is_active, width=20)
        self.cb.pack(side="left", padx=5)
        
        name = os.path.basename(file_path)
        self.btn = ctk.CTkButton(self, text=name, anchor="w", fg_color="transparent", 
                                 hover_color="#2b2b2b", command=lambda: on_preview(file_path))
        self.btn.pack(side="left", fill="x", expand=True)
        
        self.rm = ctk.CTkButton(self, text="×", width=30, fg_color="transparent", 
                                hover_color="#880000", command=lambda: on_remove(self))
        self.rm.pack(side="right", padx=5)

class Application(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.current_lang = "EN"
        self.file_entries = []
        self._initialize_window()
        self._build_ui()

    def _initialize_window(self):
        self.title(i18n.DATA[self.current_lang]["title"])
        self.geometry("1200x800")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Enable Drag and Drop
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_file_drop)

    def _build_ui(self):
        lang_pack = i18n.DATA[self.current_lang]
        
        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text=lang_pack["queue"], font=("Segoe UI", 20, "bold")).pack(pady=20)
        
        # Language Selector
        self.lang_switch = ctk.CTkSegmentedButton(self.sidebar, values=["EN", "RU"], command=self._switch_language)
        self.lang_switch.set(self.current_lang)
        self.lang_switch.pack(pady=10)

        self.file_scroll = ctk.CTkScrollableFrame(self.sidebar, label_text=lang_pack["drop_hint"])
        self.file_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.add_btn = ctk.CTkButton(self.sidebar, text="+", command=self._manual_add)
        self.add_btn.pack(pady=10)

        # Main View (Preview)
        self.main_view = ctk.CTkFrame(self, fg_color="#0f0f0f", corner_radius=15)
        self.main_view.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        self.pv_title = ctk.CTkLabel(self.main_view, text=lang_pack["preview"], font=("Segoe UI", 14))
        self.pv_title.pack(pady=10)
        
        self.preview_display = ctk.CTkLabel(self.main_view, text="")
        self.preview_display.pack(fill="both", expand=True, padx=20, pady=20)

        # Controls
        self.ctrl_panel = ctk.CTkFrame(self, width=280)
        self.ctrl_panel.grid(row=0, column=2, sticky="nsew", padx=10, pady=20)
        
        ctk.CTkLabel(self.ctrl_panel, text=lang_pack["params"], font=("Segoe UI", 16, "bold")).pack(pady=15)
        
        self.printer_cb = self._add_control(lang_pack["printer"], PrinterService.get_available_printers())
        self.paper_cb = self._add_control(lang_pack["paper"], ["A4", "Letter", "A3", "10x15"])
        
        ctk.CTkLabel(self.ctrl_panel, text=lang_pack["copies"]).pack(anchor="w", padx=20)
        self.copy_input = ctk.CTkEntry(self.ctrl_panel)
        self.copy_input.insert(0, "1")
        self.copy_input.pack(pady=5, padx=20, fill="x")

        self.fit_check = ctk.CTkCheckBox(self.ctrl_panel, text=lang_pack["fit"])
        self.fit_check.select()
        self.fit_check.pack(pady=20, padx=20)

        self.execute_btn = ctk.CTkButton(self.ctrl_panel, text=lang_pack["print_btn"], height=65,
                                        fg_color="#1a5fb4", hover_color="#3584e4",
                                        font=("Segoe UI", 18, "bold"), command=self._start_printing)
        self.execute_btn.pack(side="bottom", fill="x", padx=20, pady=20)

    def _add_control(self, label, values):
        ctk.CTkLabel(self.ctrl_panel, text=label).pack(anchor="w", padx=20, pady=(10,0))
        cb = ctk.CTkComboBox(self.ctrl_panel, values=values)
        cb.pack(pady=5, padx=20, fill="x")
        return cb

    def _switch_language(self, new_lang):
        self.current_lang = new_lang
        for widget in self.winfo_children():
            widget.destroy()
        self._initialize_window()
        self._build_ui()

    def _on_file_drop(self, event):
        files = self.tk.splitlist(event.data)
        for f in files:
            self._add_file_to_list(f)

    def _manual_add(self):
        files = filedialog.askopenfilenames()
        for f in files:
            self._add_file_to_list(f)

    def _add_file_to_list(self, path):
        item = FileListItem(self.file_scroll, path, self._remove_file, self._update_preview)
        item.pack(fill="x", pady=2)
        self.file_entries.append(item)
        if len(self.file_entries) == 1:
            self._update_preview(path)

    def _remove_file(self, item_widget):
        item_widget.destroy()
        self.file_entries.remove(item_widget)

    def _update_preview(self, path):
        try:
            if path.lower().endswith('.pdf'):
                doc = fitz.open(path)
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            else:
                img = Image.open(path)
            
            img.thumbnail((500, 600), Image.Resampling.LANCZOS)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.preview_display.configure(image=ctk_img, text="")
            self.pv_title.configure(text=f"{i18n.DATA[self.current_lang]['preview']}: {os.path.basename(path)}")
        except Exception as e:
            self.preview_display.configure(image=None, text=f"Preview Error: {e}")

    def _start_printing(self):
        targets = [e.file_path for e in self.file_entries if e.is_active.get()]
        if not targets:
            return

        cfg = {
            'printer': self.printer_cb.get(),
            'copies': self.copy_input.get()
        }
        
        self.execute_btn.configure(state="disabled", text="...")
        
        def run():
            for f in targets:
                try: PrinterService.process_print(f, cfg)
                except: pass
            self.after(0, lambda: self.execute_btn.configure(state="normal", text=i18n.DATA[self.current_lang]["print_btn"]))

        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    app = Application()
    app.mainloop()