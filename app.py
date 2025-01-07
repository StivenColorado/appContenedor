import tkinter as tk
from tkinter import ttk
import os
import logging
import time
from tkinter import messagebox
import webbrowser
from procesador_excel import ExcelProcessorApp
from procesador_pdf import PDFProcessorApp

class ModernButton(ttk.Button):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        
        style = ttk.Style()
        style.configure(
            'Modern.TButton',
            padding=15,
            font=('Helvetica', 11),
            background='#2196F3'
        )
        # Create hover style
        style.map('Modern.TButton',
            background=[('active', '#1976D2')],
            relief=[('pressed', 'groove'), ('!pressed', 'flat')]
        )
        self.configure(style='Modern.TButton')
    
    def on_enter(self, e):
        self['style'] = 'Modern.TButton'
    
    def on_leave(self, e):
        self['style'] = 'Modern.TButton'

class LoadingScreen:
    def __init__(self):
        self.window = tk.Toplevel()
        self.window.title("")
        self.window.overrideredirect(True)
        
        bg_color = '#FFFFFF'
        accent_color = '#2196F3'
        
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        width = 400
        height = 200
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.configure(bg=bg_color)
        
        self.window.attributes('-alpha', 0.95)
        
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        self.loading_label = ttk.Label(
            main_frame,
            text="Iniciando aplicación...",
            font=('Helvetica', 14),
            foreground=accent_color
        )
        self.loading_label.pack(pady=15)
        
        style = ttk.Style()
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor='#E0E0E0',
            background=accent_color,
            thickness=10
        )
        
        self.progress = ttk.Progressbar(
            main_frame,
            length=300,
            mode='determinate',
            style="Modern.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=15)
        
        self.dev_label = ttk.Label(
            main_frame,
            text="Desarrollado con ♥ por Stiven Colorado",
            font=('Helvetica', 11, 'italic'),
            foreground='#757575',
            cursor="hand2"
        )
        self.dev_label.pack(pady=15)
        self.dev_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.linkedin.com/in/stiven-colorado-370028220/"))
    
    def update_progress(self, value):
        self.progress['value'] = value
        self.window.update()
    
    def close(self):
        self.window.destroy()

class MainApp:
    def __init__(self, root):
        self.root = root
        self.setup_main_window()
        
        loading = LoadingScreen()
        for i in range(0, 101, 2):
            loading.update_progress(i)
            time.sleep(0.02)
        loading.close()
        
        self.root.deiconify()
        self.create_widgets()
    
    def setup_main_window(self):
        self.root.title("Procesador de Archivos | Desarrollado por Stiven Colorado")
        self.root.geometry("800x600")
        self.center_window()
        self.setup_app_icon()
        
        style = ttk.Style()
        style.configure('Modern.TFrame', background='#FFFFFF')
        
        self.root.configure(bg='#FFFFFF')
    
    def setup_app_icon(self):
        try:
            if os.name == 'nt':
                icon_path = self.resource_path('icon.ico')
                self.root.iconbitmap(icon_path)
            self.root.overrideredirect(False)
            self.root.iconwindow()
        except Exception as e:
            logging.error(f"Error setting up app icon: {str(e)}")
    
    def open_linkedin(self, event):
        webbrowser.open("https://www.linkedin.com/in/stiven-colorado-370028220/")
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="30", style='Modern.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        try:
            zafiro_path = self.resource_path('zafiro.png')
            zafiro_image = tk.PhotoImage(file=zafiro_path)
            zafiro_image = zafiro_image.subsample(2, 2)
            zafiro_label = ttk.Label(main_frame, image=zafiro_image)
            zafiro_label.image = zafiro_image
            zafiro_label.pack(pady=(0, 30))
        except Exception as e:
            logging.error(f"Error loading Zafiro logo: {str(e)}")
        
        title_label = ttk.Label(
            main_frame,
            text="¿Qué deseas realizar el día de hoy?",
            font=('Helvetica', 24, 'bold'),
            foreground='#1565C0'
        )
        title_label.pack(pady=30)
        
        button_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        button_frame.pack(pady=40)
        
        excel_button = ModernButton(
            button_frame,
            text="📊  Procesar Excel",
            command=self.open_excel_processor,
            width=25
        )
        excel_button.pack(pady=15)
        
        pdf_button = ModernButton(
            button_frame,
            text="📄  Procesar PDF",
            command=self.open_pdf_processor,
            width=25
        )
        pdf_button.pack(pady=15)
        
        # Nuevo label interactivo para el desarrollador
        dev_frame = ttk.Frame(main_frame)
        dev_frame.pack(side=tk.BOTTOM, pady=20)
        
        dev_label = ttk.Label(
            dev_frame,
            text="Diseñado y desarrollado por ",
            font=('Helvetica', 10),
            foreground='#757575'
        )
        dev_label.pack(side=tk.LEFT)
        
        dev_name_label = ttk.Label(
            dev_frame,
            text="Stiven Colorado",
            font=('Helvetica', 10, 'bold'),
            foreground='#2196F3',
            cursor="hand2"  # Cambia el cursor a una mano
        )
        dev_name_label.pack(side=tk.LEFT)
        
        # Vincula el evento de clic para abrir LinkedIn
        dev_name_label.bind("<Button-1>", self.open_linkedin)
        
        # Efectos de hover
        def on_enter(e):
            dev_name_label.configure(foreground='#1976D2')
            
        def on_leave(e):
            dev_name_label.configure(foreground='#2196F3')
        
        dev_name_label.bind("<Enter>", on_enter)
        dev_name_label.bind("<Leave>", on_leave)
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    def open_excel_processor(self):
        self.root.withdraw()
        excel_window = tk.Toplevel(self.root)
        ExcelProcessorApp(excel_window)
        def on_close():
            self.root.deiconify()
            excel_window.destroy()
        excel_window.protocol("WM_DELETE_WINDOW", on_close)
    
    def open_pdf_processor(self):
        self.root.withdraw()
        pdf_window = tk.Toplevel(self.root)
        PDFProcessorApp(pdf_window)
        def on_close():
            self.root.deiconify()
            pdf_window.destroy()
        pdf_window.protocol("WM_DELETE_WINDOW", on_close)

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()