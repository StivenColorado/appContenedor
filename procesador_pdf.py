import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import PyPDF2
from tkinter import StringVar, Canvas
import fitz
import re
from PIL import Image, ImageTk
from datetime import datetime
import pytesseract
import platform
import cv2
import numpy as np

# Set TESSDATA_PREFIX environment variable
os.environ['TESSDATA_PREFIX'] = '/usr/local/share/tessdata/'

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.create_widgets()
        self.ocr_lang = 'spa'
    def setup_window(self):
        self.root.title("Selección de Archivo PDF")
        self.root.geometry("500x300")
        self.center_window()
        
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(
            main_frame,
            text="Procesador de PDF - Subpartidas",
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=20)
        
        # Input file
        self.input_path = StringVar()
        input_label = ttk.Label(main_frame, text="Archivo PDF:", font=('Helvetica', 10))
        input_label.pack(anchor=tk.W)
        
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_path, state='readonly')
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        input_button = ttk.Button(
            input_frame,
            text="Seleccionar archivo",
            command=self.select_input_file
        )
        input_button.pack(side=tk.RIGHT)
        
        # Output directory
        self.output_path = StringVar()
        output_label = ttk.Label(main_frame, text="Carpeta de salida:", font=('Helvetica', 10))
        output_label.pack(anchor=tk.W, pady=(10, 0))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path, state='readonly')
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        output_button = ttk.Button(
            output_frame,
            text="Seleccionar carpeta",
            command=self.select_output_directory
        )
        output_button.pack(side=tk.RIGHT)
        
        # Process button
        self.process_button = ttk.Button(
            main_frame,
            text="Procesar PDF",
            command=self.start_processing
        )
        self.process_button.pack(pady=20)
        
    def select_input_file(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if filename:
            self.input_path.set(filename)
            
    def select_output_directory(self):
        directory = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if directory:
            self.output_path.set(directory)
            
    def start_processing(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Por favor seleccione el archivo PDF y la carpeta de salida")
            return
            
        self.root.withdraw()  # Ocultar ventana actual
        PreviewWindow(self.input_path.get(), self.output_path.get(), self)
        
    def show(self):
        self.root.deiconify()  # Mostrar ventana nuevamente
        
    def run(self):
        self.root.mainloop()

class PreviewWindow:
    def __init__(self, input_path, output_path, parent_window):
        self.root = tk.Toplevel()
        self.input_path = input_path
        self.output_path = output_path
        self.parent_window = parent_window
        self.current_page = 0
        self.pdf_document = None
        self.preview_image = None
        self.subpartida_numbers = []
        self.backups = []
        self.page_types = ['p']
        self.zoom_factor = 1.5
        self.ocr_lang = 'spa'
        
        if platform.system() == 'Windows':
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        
        self.setup_window()
        self.create_widgets()
        self.load_pdf()

    def preprocess_image(self, image):
        # Convert PIL Image to OpenCV format
        opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        # Convert to grayscale
        gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
        
        # Apply adaptive thresholding
        thresh = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
        )
        
        return thresh

    def extract_text_from_roi(self, image, roi_coords, config):
        # Extract region of interest (ROI)
        x1, y1, x2, y2 = roi_coords
        roi = image[y1:y2, x1:x2]
        
        # Extract text with Tesseract
        text = pytesseract.image_to_string(roi, config=config)
        
        return text

    def check_tesseract_installation(self):
        try:
            pytesseract.get_tesseract_version()
            return True
        except Exception:
            return False

    def extract_numbers_from_text(self, text):
        match = re.search(r'\d+', text)
        if match:
            return match.group()
        return None

    def setup_window(self):
        self.root.title("Previsualización y Edición")
        # Configurar pantalla completa real
        self.root.attributes('-fullscreen', True)
        # Agregar botón de escape para salir de pantalla completa
        self.root.bind('<Escape>', lambda e: self.toggle_fullscreen())
        
    def toggle_fullscreen(self):
        if self.root.attributes('-fullscreen'):
            self.root.attributes('-fullscreen', False)
        else:
            self.root.attributes('-fullscreen', True)
            
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame superior para controles
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Navegación
        nav_frame = ttk.Frame(control_frame)
        nav_frame.pack(side=tk.LEFT)
        
        self.prev_button = ttk.Button(nav_frame, text="← Anterior", command=self.prev_page)
        self.prev_button.pack(side=tk.LEFT, padx=5)
        
        self.page_label = ttk.Label(nav_frame, text="Página: 0/0")
        self.page_label.pack(side=tk.LEFT, padx=20)
        
        self.next_button = ttk.Button(nav_frame, text="Siguiente →", command=self.next_page)
        self.next_button.pack(side=tk.LEFT, padx=5)
        
        # Frame para número de subpartida
        subpartida_frame = ttk.Frame(control_frame)
        subpartida_frame.pack(side=tk.LEFT, padx=20)
        
        self.subpartida_var = StringVar()
        ttk.Label(subpartida_frame, text="Número de Subpartida:", font=('Helvetica', 10)).pack(side=tk.LEFT, padx=5)
        self.subpartida_entry = ttk.Entry(subpartida_frame, textvariable=self.subpartida_var, width=20)
        self.subpartida_entry.pack(side=tk.LEFT, padx=5)
        
        # Frame para tipo de página
        tipo_frame = ttk.Frame(control_frame)
        tipo_frame.pack(side=tk.LEFT, padx=20)
        
        self.tipo_var = StringVar()
        ttk.Label(tipo_frame, text="Tipo de Página (p/e):", font=('Helvetica', 10)).pack(side=tk.LEFT, padx=5)
        self.tipo_entry = ttk.Entry(tipo_frame, textvariable=self.tipo_var, width=5)
        self.tipo_entry.pack(side=tk.LEFT, padx=5)
        
        # Botones de acción
        action_frame = ttk.Frame(control_frame)
        action_frame.pack(side=tk.RIGHT)
        
        self.select_button = ttk.Button(action_frame, text="Seleccionar Texto", command=self.select_text)
        self.select_button.pack(side=tk.RIGHT, padx=5)
        
        self.save_button = ttk.Button(action_frame, text="Guardar", command=self.save_pdfs)
        self.save_button.pack(side=tk.RIGHT, padx=5)
        
        # Frame para preview con scrollbars
        preview_frame = ttk.Frame(main_frame)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        self.v_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical")
        self.h_scrollbar = ttk.Scrollbar(preview_frame, orient="horizontal")
        
        # Canvas
        self.canvas = Canvas(
            preview_frame,
            bg='white',
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set
        )
        
        # Configurar scrollbars
        self.v_scrollbar.config(command=self.canvas.yview)
        self.h_scrollbar.config(command=self.canvas.xview)
        
        # Posicionar elementos
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
    def load_pdf(self):
        try:
            self.pdf_document = fitz.open(self.input_path)
            self.update_page_display()
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el PDF: {str(e)}")

    def get_page_text_ocr(self, page):
        try:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Use default language if Spanish not available
            try:
                text = pytesseract.image_to_string(img, lang='spa')
            except pytesseract.TesseractError:
                text = pytesseract.image_to_string(img)
            
            return text
        except Exception as e:
            print(f"OCR Error: {str(e)}")
            return ""

    def detect_subpartidas_and_backups(self):
        if not self.pdf_document:
            return

        patterns = [
            r'[Ss]ubpartida\s*[Aa]rancelaria\s*[:#]?\s*(\d{4,}(?:\.\d+)?)',
            r'[Ss]ubpartida\s*[:#]?\s*(\d{4,}(?:\.\d+)?)',
            r'SUBPARTIDA\s+ARANCELARIA\s*[:#]?\s*(\d{4,}(?:\.\d+)?)',
            r'(?:^|\s)(\d{4}\.\d{2}\.\d{2})(?:\s|$)',
            r'(?<=subpartida\s)(\d{4,}(?:\.\d+)?)',
        ]

        for page_num in range(self.pdf_document.page_count):
            page = self.pdf_document[page_num]
            # Use OCR instead of direct text extraction
            text = self.get_page_text_ocr(page)
            print(f"\nTexto extraído por OCR de la página {page_num + 1}:")
            print(text)

            found_subpartida = False
            for pattern in patterns:
                matches = re.finditer(pattern, text, re.IGNORECASE)
                for match in matches:
                    potential_subpartida = match.group(1).strip()
                    if self.validate_subpartida_format(potential_subpartida):
                        print(f"Subpartida encontrada: {potential_subpartida}")
                        self.subpartida_numbers.append(potential_subpartida)
                        self.page_types.append('p')
                        found_subpartida = True
                        break
                if found_subpartida:
                    break

            if not found_subpartida:
                self.backups.append([page_num])
                self.page_types.append('e')
                 
    def validate_subpartida_format(self, subpartida):
        """
        Valida el formato de una subpartida
        """
        # Eliminar espacios y puntos extras
        subpartida = subpartida.strip().replace(' ', '')
        
        # Verificar formato básico (debe tener números y opcionalmente puntos)
        if not re.match(r'^\d+(\.\d+)*$', subpartida):
            return False

        # Verificar longitud mínima y máxima
        if len(subpartida.replace('.', '')) < 4 or len(subpartida.replace('.', '')) > 10:
            return False

        # Verificar que cada sección entre puntos sea válida
        parts = subpartida.split('.')
        for part in parts:
            if not (1 <= len(part) <= 4):  # Cada parte debe tener entre 1 y 4 dígitos
                return False

        return True
    
    def update_page_display(self):
        if not self.pdf_document:
            return
            
        try:
            page = self.pdf_document[self.current_page]
            # Usar una matriz de zoom más alta para mejor calidad
            pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_factor, self.zoom_factor))
            
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.preview_image = ImageTk.PhotoImage(image=img)
            
            # Actualizar canvas
            self.canvas.delete("all")
            self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.preview_image)
            
            # Actualizar etiqueta de página
            self.page_label.config(text=f"Página: {self.current_page + 1}/{self.pdf_document.page_count}")
            
            # Obtener el texto completo de la página
            text = page.get_text("text")
            print(f"Texto completo de la página {self.current_page + 1}:")
            print(text)  # Debug: mostrar todo el texto de la página
            
            # Intentar detectar subpartida en la página actual usando múltiples patrones
            patterns = [
                r'Subpartida\s+[Aa]rancelaria\s*[:#]?\s*(\d+)',
                r'SUBPARTIDA\s+ARANCELARIA\s*[:#]?\s*(\d+)',
                r'Subpartida\s*[:#]?\s*(\d+)',
            ]
            
            found_subpartida = False
            for pattern in patterns:
                match = re.search(pattern, text)
                if match:
                    detected_number = match.group(1).strip()
                    print(f"Subpartida detectada: {detected_number}")  # Debug
                    self.subpartida_var.set(detected_number)
                    found_subpartida = True
                    break
            
            if not found_subpartida:
                print("No se detectó subpartida en esta página")  # Debug
            
            # Actualizar el tipo de página
            if self.current_page < len(self.page_types):
                current_type = self.page_types[self.current_page]
                self.tipo_var.set(current_type)
                print(f"Tipo de página establecido: {current_type}")  # Debug
            
            # Actualizar el número de subpartida si existe
            subpartida_index = self.get_subpartida_index(self.current_page)
            if subpartida_index >= 0 and subpartida_index < len(self.subpartida_numbers):
                current_subpartida = self.subpartida_numbers[subpartida_index]
                if not found_subpartida:  # Solo actualizar si no se encontró una nueva
                    self.subpartida_var.set(current_subpartida)
                print(f"Subpartida del índice: {current_subpartida}")  # Debug
            
            # Asegurar que los widgets estén visibles
            self.subpartida_entry.pack(side=tk.LEFT, padx=5)
            self.tipo_entry.pack(side=tk.LEFT, padx=5)
            
            # Configurar el botón de guardar
            if self.current_page == self.pdf_document.page_count - 1:
                self.save_button.pack(side=tk.RIGHT, padx=5)
            else:
                self.save_button.pack_forget()
                
        except Exception as e:
            error_msg = f"Error en update_page_display: {str(e)}"
            print(error_msg)  # Debug
            messagebox.showerror("Error", f"Error al actualizar la visualización: {str(e)}")
          
    def get_subpartida_index(self, page_num):
        # Encontrar a qué subpartida pertenece esta página
        for i, backups in enumerate(self.backups):
            if page_num == i or page_num in backups:
                return i
        return -1
        
    def next_page(self):
        if self.pdf_document and self.current_page < self.pdf_document.page_count - 1:
            self.current_page += 1
            self.update_page_display()
            
    def prev_page(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()
            
    def select_text(self):
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        
    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, 
            self.start_x, self.start_y, 
            outline='red', 
            width=2
        )
        
    def on_mouse_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)
        
    def on_button_release(self, event):
        try:
            coords = self.canvas.coords(self.rect)
            if not coords:
                return
                
            x0, y0, x1, y1 = coords
            page = self.pdf_document[self.current_page]
            
            pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_factor, self.zoom_factor))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            selection_box = (
                min(x0, x1),
                min(y0, y1),
                max(x0, x1),
                max(y0, y1)
            )
            
            selection = img.crop(selection_box)
            processed_image = self.preprocess_image(selection)
            
            tesseract_config = r'--oem 3 --psm 6'
            height, width = processed_image.shape[:2]
            text = self.extract_text_from_roi(processed_image, (0, 0, width, height), tesseract_config)
            detected_number = self.extract_numbers_from_text(text)
            
            # Print coordinates and detected number
            print(f"Coordinates: {coords}")
            print(f"Detected number: {detected_number}")
            
            if detected_number:
                self.subpartida_var.set(detected_number)
                print(f"Número detectado: {detected_number}")
            
            # Display the selected area in a temporary window
            temp_window = tk.Toplevel(self.root)
            temp_window.title("Selected Area")
            temp_window.geometry("+0+0")  # Position at the bottom left
            temp_image = ImageTk.PhotoImage(Image.fromarray(processed_image))
            temp_label = ttk.Label(temp_window, image=temp_image)
            temp_label.image = temp_image
            temp_label.pack()
            
            # Automatically close the temporary window after a short delay
            self.root.after(3000, temp_window.destroy)
            
            self.canvas.delete(self.rect)
            
        except Exception as e:
            print(f"Error en OCR: {str(e)}")
            messagebox.showerror("Error", f"Error al procesar el texto: {str(e)}")      
    def save_pdfs(self):
        try:
            input_pdf = PyPDF2.PdfReader(self.input_path)
            current_year = datetime.now().year
            
            # Asegurarse de que tenemos la misma cantidad de subpartidas y backups
            while len(self.backups) < len(self.subpartida_numbers):
                self.backups.append([])
            
            for i, subpartida in enumerate(self.subpartida_numbers):
                output = PyPDF2.PdfWriter()
                
                # Añadir página principal
                if i < len(input_pdf.pages):
                    output.add_page(input_pdf.pages[i])
                
                # Añadir páginas de respaldo
                if i < len(self.backups):
                    for backup_page in self.backups[i]:
                        if backup_page < len(input_pdf.pages):
                            output.add_page(input_pdf.pages[backup_page])
                
                # Crear nombre de archivo con subpartida y año
                output_filename = os.path.join(
                    self.output_path,
                    f"subpartida_{subpartida}_{current_year}.pdf"
                )
                
                with open(output_filename, "wb") as output_file:
                    output.write(output_file)
            
            messagebox.showinfo("Éxito", "PDFs generados correctamente")
            self.root.destroy()
            self.parent_window.show()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar los PDFs: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()