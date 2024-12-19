import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from datetime import datetime
import io
from PIL import Image as PILImage

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Excel")
        self.root.geometry("600x400")
        
        # Configurar el estilo
        style = ttk.Style()
        style.configure('Custom.TFrame', background='#f0f0f0')
        
        # Frame principal
        self.main_frame = ttk.Frame(root, padding="20", style='Custom.TFrame')
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Etiqueta de título
        self.title_label = ttk.Label(
            self.main_frame,
            text="Procesador de Excel",
            font=('Helvetica', 16, 'bold')
        )
        self.title_label.pack(pady=20)
        
        # Frame para los archivos
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=tk.X, pady=20)
        
        # Input file
        self.input_path = tk.StringVar()
        self.input_label = ttk.Label(
            self.file_frame,
            text="Archivo de entrada:",
            font=('Helvetica', 10)
        )
        self.input_label.pack(anchor=tk.W)
        
        self.input_frame = ttk.Frame(self.file_frame)
        self.input_frame.pack(fill=tk.X, pady=5)
        
        self.input_entry = ttk.Entry(
            self.input_frame,
            textvariable=self.input_path,
            state='readonly'
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.input_button = ttk.Button(
            self.input_frame,
            text="Seleccionar archivo",
            command=self.select_input_file
        )
        self.input_button.pack(side=tk.RIGHT)
        
        # Output directory
        self.output_path = tk.StringVar()
        self.output_label = ttk.Label(
            self.file_frame,
            text="Carpeta de salida:",
            font=('Helvetica', 10)
        )
        self.output_label.pack(anchor=tk.W, pady=(20, 0))
        
        self.output_frame = ttk.Frame(self.file_frame)
        self.output_frame.pack(fill=tk.X, pady=5)
        
        self.output_entry = ttk.Entry(
            self.output_frame,
            textvariable=self.output_path,
            state='readonly'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.output_button = ttk.Button(
            self.output_frame,
            text="Seleccionar carpeta",
            command=self.select_output_directory
        )
        self.output_button.pack(side=tk.RIGHT)
        
        # Botón de procesar
        self.process_button = ttk.Button(
            self.main_frame,
            text="Procesar Excel",
            command=self.process_file
        )
        self.process_button.pack(pady=30)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(
            self.main_frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='indeterminate'
        )
        self.progress.pack(pady=10)
        
        # Status label
        self.status_label = ttk.Label(
            self.main_frame,
            text="",
            font=('Helvetica', 10)
        )
        self.status_label.pack(pady=10)

    def select_input_file(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.input_path.set(filename)

    def select_output_directory(self):
        directory = filedialog.askdirectory(
            title="Seleccionar carpeta de salida"
        )
        if directory:
            self.output_path.set(directory)

    def process_file(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror(
                "Error",
                "Por favor seleccione el archivo de entrada y la carpeta de salida"
            )
            return
        
        self.progress.start()
        self.status_label.config(text="Procesando archivo...")
        self.process_button.state(['disabled'])
        
        self.root.after(100, self.run_processing)

    def run_processing(self):
        try:
            process_excel(self.input_path.get(), self.output_path.get())
            self.progress.stop()
            self.status_label.config(text="¡Archivos procesados correctamente!")
            messagebox.showinfo(
                "Éxito",
                "Los archivos han sido procesados correctamente"
            )
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Error al procesar los archivos")
            messagebox.showerror("Error", str(e))
        finally:
            self.process_button.state(['!disabled'])
            self.progress.stop()

def clean_numeric_value(value):
    if pd.isna(value):
        return 0
    value = str(value).replace('¥', '').replace(' ', '')
    value = value.replace('.', '').replace(',', '.')
    try:
        return float(value)
    except:
        return 0

def find_header_row(df):
    keywords = ['CTNS', 'MARCA', 'CBM', 'WEIGHT', 'PRODUCTO', 'PRODUCT PICTURE']
    for idx, row in df.iterrows():
        row_str = ' '.join(str(val).upper().strip() for val in row)
        matches = sum(keyword in row_str for keyword in keywords)
        if matches >= 3:  # Aumentamos el número mínimo de coincidencias
            return idx
    return None

def find_column(columns, keyword):
    matches = [col for col in columns if keyword in col.upper()]
    if matches:
        return matches[0]
    raise ValueError(f"Columna requerida no encontrada: {keyword}")

def get_color_by_brand(brand):
    predefined_colors = {
        'SONY': 'ADD8E6',
        'XIAOMI': '90EE90',
        'SAMSUNG': 'FFD700',
        'DEFAULT': 'FFFFFF'
    }
    return predefined_colors.get(brand.upper(), predefined_colors['DEFAULT'])

def process_brand_excel(brand_df, output_path, marca, year):
    filename = f"MARCA {marca} CONSO 25 - {year}.xlsx"
    filepath = os.path.join(output_path, filename)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # Crear una copia del DataFrame sin la columna de imágenes para exportar
        export_df = brand_df.copy()
        if '产品图片' in export_df.columns:
            export_df['产品图片'] = export_df['产品图片'].fillna('N/A')
        if 'PRODUCT PICTURE' in export_df.columns:
            export_df['PRODUCT PICTURE'] = export_df['PRODUCT PICTURE'].fillna('N/A')
            
        export_df.to_excel(writer, index=False, sheet_name='Datos')
        
        workbook = writer.book
        worksheet = writer.sheets['Datos']
        
        # Procesar imágenes si existen
        img_column = None
        if 'PRODUCT PICTURE' in brand_df.columns:
            img_column = 'PRODUCT PICTURE'
        elif '产品图片' in brand_df.columns:
            img_column = '产品图片'
            
        if img_column:
            for idx, row in brand_df.iterrows():
                if pd.notna(row[img_column]) and not isinstance(row[img_column], (int, float)):
                    try:
                        img_path = str(row[img_column])
                        if os.path.exists(img_path):
                            img = Image(img_path)
                            img_col = brand_df.columns.get_loc(img_column) + 1
                            img_row = idx + 2
                            worksheet.add_image(img, f'{get_column_letter(img_col)}{img_row}')
                    except Exception as e:
                        print(f"Error al procesar imagen en fila {idx + 2}: {str(e)}")
        
        # Procesar sumas y estilos
        suma_columns = {}
        for col_name, keyword in [('CTNS', 'CTNS'), ('T/CBM', 'CBM'), ('T/WEIGHT (KG)', 'WEIGHT')]:
            try:
                column = find_column(brand_df.columns, keyword)
                suma_columns[col_name] = column
            except ValueError:
                print(f"Columna {keyword} no encontrada")
        
        for column_name, full_column_name in suma_columns.items():
            col_index = list(brand_df.columns).index(full_column_name) + 1
            column_letter = get_column_letter(col_index)
            num_rows = len(brand_df) + 1
            formula = f'=SUM({column_letter}2:{column_letter}{num_rows})'
            sum_cell = f'{column_letter}{num_rows + 2}'
            worksheet[sum_cell] = formula
            worksheet[sum_cell].font = Font(bold=True)
            worksheet[sum_cell].alignment = Alignment(horizontal='right')
        
        # Aplicar colores por marca
        color_hex = get_color_by_brand(marca)
        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        for row in range(2, len(brand_df) + 2):
            for col in range(1, len(brand_df.columns) + 1):
                worksheet.cell(row=row, column=col).fill = fill

def process_excel(input_path, output_path):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(input_path, header=None)
        header_row = find_header_row(df)

        if header_row is None:
            raise ValueError("No se encontró la fila de encabezados adecuada.")

        df = pd.read_excel(input_path, header=header_row)
        
        # Identificar la columna de marca
        marca_col = None
        for col in df.columns:
            if '品牌' in str(col) or 'MARCA' in str(col).upper():
                marca_col = col
                break
                
        if marca_col is None:
            raise ValueError("No se encontró la columna de marca")
        
        # Obtener el año actual
        current_year = str(datetime.now().year)[-2:]
        
        # Procesar cada marca por separado
        unique_brands = df[marca_col].unique()
        for marca in unique_brands:
            if pd.notna(marca):  # Asegurarse de que la marca no sea NA/None
                brand_df = df[df[marca_col] == marca].copy()
                process_brand_excel(brand_df, output_path, str(marca), current_year)

        # Crear archivo de resumen general
        summary_filename = f"RESUMEN_GENERAL_CONSO_25-{current_year}.xlsx"
        summary_filepath = os.path.join(output_path, summary_filename)
        
        with pd.ExcelWriter(summary_filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Datos', index=False)
            
            # Crear hoja de resumen por marca
            try:
                numeric_columns = df.select_dtypes(include=[np.number]).columns
                summary_columns = {}
                for keyword in ['CTNS', 'CBM', 'WEIGHT']:
                    for col in df.columns:
                        if keyword in str(col).upper():
                            summary_columns[keyword] = col
                            break
                
                if summary_columns:
                    summary = df.groupby(marca_col).agg({
                        col: 'sum' for col in summary_columns.values()
                    }).reset_index()
                    
                    summary.to_excel(writer, sheet_name='Resumen por Marca', index=False)
            except Exception as e:
                print(f"Error al crear resumen: {str(e)}")

        return True

    except Exception as e:
        raise ValueError(f"Error al procesar el archivo: {str(e)}")

def main():
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()