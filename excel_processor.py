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
import logging

# Configure logging

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
            consolidado = "CONSOLIDADO"  # Define the consolidado value
            process_excel(self.input_path.get(), self.output_path.get(), consolidado)
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

def process_brand_excel(brand_df, input_path, header_end_row, output_path, marca, year, consolidado):
    filename = f"MARCA {marca} {consolidado}-{year}.xlsx"
    filepath = os.path.join(output_path, filename)
    
    # Copiar estructura original incluyendo imágenes
    src_wb, images_info = copy_workbook_structure(input_path, header_end_row)
    src_ws = src_wb.active
    
    # Crear nuevo archivo
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # Escribir los datos de la marca
        brand_df.columns = brand_df.columns.str.strip()  # Trim whitespace from column names
        brand_df = brand_df.drop(columns=['UNIT PRICE (RMB)', 'AMOUNT (RMB)'], errors='ignore')
        brand_df.to_excel(writer, sheet_name='Datos', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Datos']
        
        # Copiar imágenes del encabezado
        for img_info in images_info:
            try:
                img_path = img_info['image'].path  # Ruta de la imagen original
                if os.path.exists(img_path):
                    img = Image(img_path)
                    cell = worksheet.cell(row=img_info['row'], column=img_info['col'])
                    worksheet.add_image(img, cell.coordinate)
            except Exception as e:
                print(f"Error al copiar imagen: {str(e)}")
        
        # Aplicar color de marca a las filas de datos
        color_hex = get_color_by_brand(marca)
        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        for row in range(2, len(brand_df) + 2):
            for col in range(1, len(brand_df.columns) + 1):
                worksheet.cell(row=row, column=col).fill = fill
        
        # Procesar sumas para columnas totales
        total_columns = {
            'T/CBM': 'T/CBM',
            'T/WEIGHT (KG)': 'T/WEIGHT (KG)',
            'CTNS': 'CTNS'
        }
        
        total_row = len(brand_df) + 2
        worksheet.cell(row=total_row, column=1, value='TOTAL')
        
        totals = {}
        for display_name, col_name in total_columns.items():
            if col_name in brand_df.columns:
                col_index = list(brand_df.columns).index(col_name) + 1
                column_letter = get_column_letter(col_index + 1)
                formula = f'=SUM({column_letter}2:{column_letter}{total_row - 1})'
                worksheet.cell(row=total_row, column=col_index + 1, value=formula)
                worksheet.cell(row=total_row, column=col_index + 1).font = Font(bold=True)
                worksheet.cell(row=total_row, column=col_index + 1).alignment = Alignment(horizontal='right')
                totals[col_name] = brand_df[col_name].sum()
        
        # Print totals to console
        print(f"File: {filename}")
        for col_name, total in totals.items():
            print(f"{col_name}: {total}")

        # Clear other cells in the total row
        for col in range(2, len(brand_df.columns) + 1):
            if col not in [list(brand_df.columns).index(col_name) + 1 for col_name in total_columns.values()]:
                worksheet.cell(row=total_row, column=col).value = None

def process_excel(input_path, output_path, consolidado):
    try:
        logging.debug(f"Processing Excel file: {input_path}")
        logging.debug(f"Output directory: {output_path}")
        logging.debug(f"Consolidado: {consolidado}")
        
        # Leer el archivo completo primero sin especificar encabezados
        df_raw = pd.read_excel(input_path, header=None)
        logging.debug("Excel file read successfully")
        
        # Encontrar la fila real de encabezados
        header_row = find_real_header_row(df_raw)
        if header_row is None:
            raise ValueError("No se encontró la fila de encabezados adecuada")
        logging.debug(f"Header row found at index: {header_row}")
        
        # Leer el archivo nuevamente usando la fila de encabezados correcta
        df = pd.read_excel(input_path, header=header_row)
        logging.debug("Excel file re-read with correct header row")
        
        # Trim whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Print columns to debug
        print("Columns received from Excel:", df.columns.tolist())
        
        # Eliminar columnas UNIT PRICE (RMB) y AMOUNT (RMB) si existen
        df = df.drop(columns=['UNIT PRICE (RMB)', 'AMOUNT (RMB)'], errors='ignore')
        
        # Identificar columna de marca
        marca_col = None
        for col in df.columns:
            if any(keyword in str(col).upper() for keyword in ['品牌', 'MARCA']):
                marca_col = col
                break
                
        if marca_col is None:
            raise ValueError("No se encontró la columna de marca")
        logging.debug(f"Marca column found: {marca_col}")
        
        # Obtener año actual
        current_year = str(datetime.now().year)[-2:]
        
        # Limpiar y convertir la columna de marca a string
        df[marca_col] = df[marca_col].fillna('SIN MARCA').astype(str)
        
        # Procesar cada marca por separado
        unique_brands = df[marca_col].unique()
        logging.debug(f"Unique brands found: {unique_brands}")
        for marca in unique_brands:
            if marca and marca.strip():  # Verificar que la marca no esté vacía
                brand_df = df[df[marca_col] == marca].copy()
                logging.debug(f"Processing brand: {marca}")
                process_brand_excel(brand_df, input_path, header_row, output_path, marca, current_year, consolidado)
        
        # Crear archivo de resumen general
        summary_filename = f"RESUMEN_GENERAL_CONSO_{consolidado}-{current_year}.xlsx"
        summary_filepath = os.path.join(output_path, summary_filename)
        logging.debug(f"Creating summary file: {summary_filepath}")
        
        with pd.ExcelWriter(summary_filepath, engine='openpyxl') as writer:
            # Escribir datos principales
            df.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            # Crear hoja RESULTADOS
            try:
                summary_columns = {
                    'CTNS': 'CARTONES',
                    'T/CBM': 'CUBICAJE',
                    'T/WEIGHT (KG)': 'PESO'
                }
                
                # Buscar las columnas correctas en el DataFrame
                agg_dict = {}
                for old_col in summary_columns.keys():
                    matching_cols = [col for col in df.columns if old_col in str(col)]
                    if matching_cols:
                        agg_dict[matching_cols[0]] = 'sum'
                
                if agg_dict:
                    summary = df.groupby(marca_col).agg(agg_dict).reset_index()
                    # Renombrar columnas
                    new_columns = [marca_col]
                    for col in summary.columns[1:]:
                        for old_col, new_col in summary_columns.items():
                            if old_col in str(col):
                                new_columns.append(new_col)
                                break
                        else:
                            new_columns.append(col)
                    summary.columns = new_columns
                    
                    summary.to_excel(writer, sheet_name='RESULTADOS', index=False)
                    
                    # Asegurar que la hoja RESULTADOS esté visible
                    writer.book['RESULTADOS'].sheet_state = 'visible'
                    
                    # Agregar fila de totales
                    workbook = writer.book
                    worksheet = writer.sheets['RESULTADOS']
                    row_num = len(summary) + 2
                    worksheet.cell(row=row_num, column=1, value='TOTAL')
                    
                    for col_idx, col in enumerate(summary.columns[1:], start=2):
                        column_letter = get_column_letter(col_idx)
                        formula = f'=SUM({column_letter}2:{column_letter}{row_num-1})'
                        worksheet.cell(row=row_num, column=col_idx, value=formula)
                        worksheet.cell(row=row_num, column=col_idx).font = Font(bold=True)
                        worksheet.cell(row=row_num, column=col_idx).alignment = Alignment(horizontal='right')
            except Exception as e:
                logging.error(f"Error al crear la hoja de resultados: {str(e)}")
            
            # Asegurar que la hoja RESUMEN esté visible
            writer.book['RESUMEN'].sheet_state = 'visible'
        logging.debug("Summary file created successfully")
    except Exception as e:
        logging.error(f"Error al procesar el archivo: {str(e)}")

def find_real_header_row(df):
    keywords = ['CTNS', 'MARCA', 'CBM', 'WEIGHT', 'PRODUCTO', 'PRODUCT PICTURE']
    for idx, row in df.iterrows():
        row_str = ' '.join(str(val).upper().strip() for val in row)
        matches = sum(keyword in row_str for keyword in keywords)
        if matches >= 3:  # Aumentamos el número mínimo de coincidencias
            return idx
    return None

def copy_workbook_structure(input_path, header_end_row):
    from openpyxl import load_workbook

    wb = load_workbook(input_path)
    ws = wb.active

    images_info = []
    for image in ws._images:
        images_info.append({
            'image': image,
            'row': image.anchor._from.row + 1,
            'col': image.anchor._from.col + 1
        })

    return wb, images_info

def main():
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()