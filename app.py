# app.py
import tkinter as tk
from tkinter import ttk
from procesador_excel import ExcelProcessorApp
from procesador_pdf import PDFProcessorApp
import os
import logging
import time
from tkinter import messagebox

class LoadingScreen:
    def __init__(self):
        self.window = tk.Toplevel()
        self.window.title("")
        self.window.overrideredirect(True)
        
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        width = 300
        height = 150
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.configure(bg='#f0f0f0')
        
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.loading_label = ttk.Label(
            main_frame,
            text="Cargando...",
            font=('Helvetica', 12)
        )
        self.loading_label.pack(pady=10)
        
        self.progress = ttk.Progressbar(
            main_frame,
            length=200,
            mode='determinate'
        )
        self.progress.pack(pady=10)
        
        self.dev_label = ttk.Label(
            main_frame,
            text="Desarrollado por Stiven Colorado",
            font=('Helvetica', 10, 'italic')
        )
        self.dev_label.pack(pady=10)
    
    def update_progress(self, value):
        self.progress['value'] = value
        self.window.update()
    
    def close(self):
        self.window.destroy()

class MainApp:
    def __init__(self, root):
        self.root = root
        self.setup_main_window()
        
        # Mostrar pantalla de carga
        loading = LoadingScreen()
        for i in range(0, 101, 2):
            loading.update_progress(i)
            time.sleep(0.02)
        loading.close()
        
        self.root.deiconify()
        self.create_widgets()
    
    def setup_main_window(self):
        self.root.title("Procesador de Archivos")
        self.root.geometry("700x500")
        self.center_window()
        self.setup_app_icon()
        
        style = ttk.Style()
        style.configure('Custom.TFrame', background='#f0f0f0')
    
    def setup_app_icon(self):
        try:
            if os.name == 'nt':
                icon_path = self.resource_path('icon.ico')
                self.root.iconbitmap(icon_path)
            self.root.overrideredirect(False)
            self.root.iconwindow()
        except Exception as e:
            logging.error(f"Error setting up app icon: {str(e)}")
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20", style='Custom.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Logo Zafiro
        try:
            zafiro_path = self.resource_path('zafiro.png')
            zafiro_image = tk.PhotoImage(file=zafiro_path)
            zafiro_image = zafiro_image.subsample(2, 2)
            zafiro_label = ttk.Label(main_frame, image=zafiro_image)
            zafiro_label.image = zafiro_image
            zafiro_label.pack(pady=(0, 20))
        except Exception as e:
            logging.error(f"Error loading Zafiro logo: {str(e)}")
        
        # Título
        title_label = ttk.Label(
            main_frame,
            text="¿Qué deseas realizar el día de hoy?",
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=20)
        
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=30)
        
        excel_button = ttk.Button(
            button_frame,
            text="Procesar Excel",
            command=self.open_excel_processor
        )
        excel_button.pack(pady=10)
        
        pdf_button = ttk.Button(
            button_frame,
            text="Procesar PDF",
            command=self.open_pdf_processor
        )
        pdf_button.pack(pady=10)
    
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
        excel_window = tk.Toplevel(self.root)
        ExcelProcessorApp(excel_window)
    
    def open_pdf_processor(self):
        self.root.withdraw()  # Ocultar ventana principal
        pdf_window = tk.Toplevel(self.root)
        PDFProcessorApp(pdf_window)

        # Definir el comportamiento al cerrar la ventana secundaria
        def on_close():
            self.root.deiconify()  # Mostrar ventana principal
            pdf_window.destroy()  # Cerrar la ventana secundaria

        pdf_window.protocol("WM_DELETE_WINDOW", on_close)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()