# procesador_pdf.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from tkinter import StringVar
import logging

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.create_widgets()
    
    def setup_window(self):
        self.root.title("Procesador de PDF")
        self.root.geometry("700x500")
        self.center_window()
        
        style = ttk.Style()
        style.configure('Custom.TFrame', background='#f0f0f0')
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20", style='Custom.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(
            main_frame,
            text="Separador de PDF",
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=20)
        
        # Frame para los archivos
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=20)
        
        # Input file
        self.input_path = StringVar()
        input_label = ttk.Label(
            file_frame,
            text="Archivo PDF:",
            font=('Helvetica', 10)
        )
        input_label.pack(anchor=tk.W)
        
        input_frame = ttk.Frame(file_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        self.input_entry = ttk.Entry(
            input_frame,
            textvariable=self.input_path,
            state='readonly'
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        input_button = ttk.Button(
            input_frame,
            text="Seleccionar archivo",
            command=self.select_input_file
        )
        input_button.pack(side=tk.RIGHT)
        
        # Output directory
        self.output_path = StringVar()
        output_label = ttk.Label(
            file_frame,
            text="Carpeta de salida:",
            font=('Helvetica', 10)
        )
        output_label.pack(anchor=tk.W, pady=(20, 0))
        
        output_frame = ttk.Frame(file_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_entry = ttk.Entry(
            output_frame,
            textvariable=self.output_path,
            state='readonly'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        output_button = ttk.Button(
            output_frame,
            text="Seleccionar carpeta",
            command=self.select_output_directory
        )
        output_button.pack(side=tk.RIGHT)
        
        # Botón de procesar
        self.process_button = ttk.Button(
            main_frame,
            text="Procesar PDF",
            command=self.process_pdf
        )
        self.process_button.pack(pady=30)
        
        # Status label
        self.status_label = ttk.Label(
            main_frame,
            text="",
            font=('Helvetica', 10)
        )
        self.status_label.pack(pady=10)
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def select_input_file(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if filename:
            self.input_path.set(filename)
    
    def select_output_directory(self):
        directory = filedialog.askdirectory(
            title="Seleccionar carpeta de salida"
        )
        if directory:
            self.output_path.set(directory)
    
    def process_pdf(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Por favor seleccione el archivo PDF y la carpeta de salida")
            return
        
        try:
            self.status_label.config(text="Procesando PDF...")
            self.process_button.state(['disabled'])
            self.root.update()
            
            input_pdf = PyPDF2.PdfReader(self.input_path.get())
            base_name = os.path.splitext(os.path.basename(self.input_path.get()))[0]
            
            for page_num in range(len(input_pdf.pages)):
                output = PyPDF2.PdfWriter()
                output.add_page(input_pdf.pages[page_num])
                
                output_filename = os.path.join(
                    self.output_path.get(),
                    f"{base_name}_pagina_{page_num + 1}.pdf"
                )
                
                with open(output_filename, "wb") as output_file:
                    output.write(output_file)
            
            self.status_label.config(text="¡PDF procesado correctamente!")
            messagebox.showinfo(
                "Éxito",
                f"Se han generado {len(input_pdf.pages)} archivos PDF"
            )
        except Exception as e:
            self.status_label.config(text="Error al procesar el PDF")
            messagebox.showerror("Error", str(e))
        finally:
            self.process_button.state(['!disabled'])