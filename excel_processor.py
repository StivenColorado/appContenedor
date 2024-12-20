import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from datetime import datetime
import logging
import tempfile
import shutil
import io
from PIL import Image as PILImage
from fpdf import FPDF

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

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
            messagebox.showerror("Error", "Por favor seleccione el archivo de entrada y la carpeta de salida")
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

def extract_and_save_images(workbook, temp_dir, header_row):
    """
    Extract images from Excel and categorize them as header or product images
    """
    images_info = {
        'header': [],
        'products': []
    }
    
    for sheet in workbook.worksheets:
        for image in sheet._images:
            img_cell = f"{image.anchor._from.col}_{image.anchor._from.row}"
            img_path = os.path.join(temp_dir, f"image_{img_cell}.png")
            
            try:
                # Get image data from BytesIO
                image_data = image.ref
                # Convert to PIL Image
                pil_image = PILImage.open(io.BytesIO(image_data.getvalue()))
                # Save as PNG
                pil_image.save(img_path, 'PNG')
                
                # Categorize image based on row position
                image_info = {
                    'path': img_path,
                    'row': image.anchor._from.row,
                    'col': image.anchor._from.col,
                    'cell_reference': img_cell
                }
                
                # If image is above or at header row, it's a header image
                if image.anchor._from.row <= header_row:
                    images_info['header'].append(image_info)
                else:
                    images_info['products'].append(image_info)
                
                logging.debug(f"Successfully saved image: {img_path}")
            except Exception as e:
                logging.error(f"Failed to save image {img_cell}: {str(e)}")
    
    return images_info

def process_brand_excel(brand_df, output_path, marca, year, consolidado, images_info, start_row):
    filename = f"MARCA {marca} {consolidado}-{year}.xlsx"
    filepath = os.path.join(output_path, filename)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # Write the data starting after header images
        brand_df.to_excel(writer, sheet_name='Datos', index=False, startrow=len(images_info['header']))
        
        workbook = writer.book
        worksheet = writer.sheets['Datos']
        
        # Add header images first
        for img_info in images_info['header']:
            try:
                img = Image(img_info['path'])
                cell = worksheet.cell(row=img_info['row'] + 1, 
                                   column=img_info['col'] + 1)
                worksheet.add_image(img, cell.coordinate)
            except Exception as e:
                logging.error(f"Failed to add header image: {str(e)}")
        
        # Find the image column index (assuming it's named 'PRODUCT PICTURE' or similar)
        image_col_idx = None
        for idx, col in enumerate(brand_df.columns):
            if 'PRODUCT PICTURE' in str(col).upper():
                image_col_idx = idx
                break
        
        if image_col_idx is None:
            logging.warning("Product image column not found")
            return
        
        # Add product images in their corresponding rows
        for idx, row in brand_df.iterrows():
            row_num = idx + len(images_info['header']) + 2  # +2 for header row and 0-based index
            
            # Find corresponding product image
            product_images = [img for img in images_info['products'] 
                            if img['row'] == idx + start_row]
            
            if product_images:
                try:
                    img = Image(product_images[0]['path'])
                    cell = worksheet.cell(row=row_num, 
                                       column=image_col_idx + 1)
                    worksheet.add_image(img, cell.coordinate)
                except Exception as e:
                    logging.error(f"Failed to add product image for row {row_num}: {str(e)}")
        
        # Add autosum formulas
        last_row = len(brand_df) + len(images_info['header']) + 1
        sum_row = last_row + 1
        
        sum_columns = {
            'CTNS': 'Total Cartones',
            'T/CBM': 'Total CBM',
            'T/WEIGHT (KG)': 'Total Peso'
        }
        
        for col_name, sum_label in sum_columns.items():
            if col_name in brand_df.columns:
                col_idx = brand_df.columns.get_loc(col_name) + 1
                col_letter = get_column_letter(col_idx)
                
                # Adjust formula to account for header images
                start_data_row = len(images_info['header']) + 2
                formula = f'=SUM({col_letter}{start_data_row}:{col_letter}{last_row})'
                
                cell = worksheet.cell(row=sum_row, column=col_idx)
                cell.value = formula
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='right')
                
                label_cell = worksheet.cell(row=sum_row, column=col_idx-1)
                label_cell.value = sum_label
                label_cell.font = Font(bold=True)
                
def create_pdf_from_excel(excel_path, pdf_path):
    workbook = load_workbook(excel_path)
    sheet = workbook.active
    
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    
    for row in sheet.iter_rows(values_only=True):
        row_data = [str(cell) if cell is not None else '' for cell in row]
        pdf.cell(200, 10, txt=" | ".join(row_data).encode('latin-1', 'replace').decode('latin-1'), ln=True)
    
    pdf.output(pdf_path)

def process_excel(input_path, output_path, consolidado):
    try:
        # Create temporary directory for images
        temp_dir = tempfile.mkdtemp()
        
        # Load workbook
        workbook = load_workbook(input_path)
        
        # Find header row first
        df_temp = pd.read_excel(input_path, header=None)
        header_row = None
        for idx, row in df_temp.iterrows():
            row_values = [str(val).upper().strip() for val in row.values]
            row_text = ' '.join(row_values)
            if 'MARCA DEL PRODUCTO' in row_text or 'PRODUCT PICTURE' in row_text:
                header_row = idx
                break
        
        if header_row is None:
            raise ValueError("No se encontró la fila de encabezados")
        
        # Extract and categorize images
        images_info = extract_and_save_images(workbook, temp_dir, header_row)
        logging.debug(f"Extracted {len(images_info['header'])} header images and {len(images_info['products'])} product images")
        
        # Read data with correct header
        df = pd.read_excel(input_path, header=header_row)
        
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Remove unnecessary columns if they exist
        columns_to_drop = ['UNIT PRICE\n(RMB)', 'AMOUNT \n(RMB)']
        df = df.drop(columns=[col for col in columns_to_drop if col in df.columns])
        
        # Process each brand
        marca_col = 'MARCA DEL PRODUCTO'
        current_year = str(datetime.now().year)[-2:]
        
        for marca in df[marca_col].unique():
            if pd.notna(marca):
                brand_df = df[df[marca_col] == marca].copy()
                process_brand_excel(brand_df, output_path, marca, current_year, 
                                 consolidado, images_info, header_row + 2)
        
        # Crear archivo de resumen general
        summary_filename = f"RESUMEN_GENERAL_CONSO_{consolidado}-{current_year}.xlsx"
        summary_filepath = os.path.join(output_path, summary_filename)
        logging.debug(f"Creating summary file: {summary_filepath}")
        
        with pd.ExcelWriter(summary_filepath, engine='openpyxl') as writer:
            # Write main data sheet
            df.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            # Create RESULTADOS sheet
            try:
                # Define columns for summary
                summary_columns = {
                    'CTNS': 'CARTONES',
                    'T/CBM': 'CUBICAJE',
                    'T/WEIGHT (KG)': 'PESO'
                }
                
                # Find matching columns in DataFrame
                agg_dict = {}
                for old_col, new_col in summary_columns.items():
                    matching_cols = [col for col in df.columns if old_col in str(col)]
                    if matching_cols:
                        agg_dict[matching_cols[0]] = 'sum'
                
                if agg_dict:
                    # Create summary by brand
                    summary = df.groupby(marca_col).agg(agg_dict).reset_index()
                    
                    # Rename columns
                    new_columns = [marca_col]
                    for col in summary.columns[1:]:
                        for old_col, new_col in summary_columns.items():
                            if old_col in str(col):
                                new_columns.append(new_col)
                                break
                        else:
                            new_columns.append(col)
                    summary.columns = new_columns
                    
                    # Write summary to Excel
                    summary.to_excel(writer, sheet_name='RESULTADOS', index=False)
                    
                    # Get worksheet reference
                    worksheet = writer.sheets['RESULTADOS']
                    
                    # Add totals row
                    row_num = len(summary) + 2
                    worksheet.cell(row=row_num, column=1, value='TOTAL')
                    
                    # Add sum formulas for numeric columns
                    for col_idx, col in enumerate(summary.columns[1:], start=2):
                        column_letter = get_column_letter(col_idx)
                        formula = f'=SUM({column_letter}2:{column_letter}{row_num-1})'
                        cell = worksheet.cell(row=row_num, column=col_idx, value=formula)
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal='right')
                    
                    # Apply formatting to summary sheet
                    for row in worksheet.iter_rows(min_row=1, max_row=1):
                        for cell in row:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal='center')
                    
                    # Adjust column widths
                    for column_cells in worksheet.columns:
                        length = max(len(str(cell.value)) for cell in column_cells)
                        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 2
                
            except Exception as e:
                logging.error(f"Error creating results sheet: {str(e)}")
            
            # Format main sheet
            worksheet = writer.sheets['RESUMEN']
            for row in worksheet.iter_rows(min_row=1, max_row=1):
                for cell in row:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
            
            # Add images to main sheet
            for img_info in images_info['products']:
                try:
                    img = Image(img_info['path'])
                    cell = worksheet.cell(row=img_info['row'] + 1, 
                                       column=img_info['col'] + 1)
                    worksheet.add_image(img, cell.coordinate)
                except Exception as e:
                    logging.error(f"Failed to add image to summary: {str(e)}")
        
        # Create PDF from Excel
        pdf_filename = f"RESUMEN_GENERAL_CONSO_{consolidado}-{current_year}.pdf"
        pdf_filepath = os.path.join(output_path, pdf_filename)
        create_pdf_from_excel(summary_filepath, pdf_filepath)
        
        # Clean up temporary files
        shutil.rmtree(temp_dir)
        logging.debug("Finished processing successfully")
        
    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        if 'temp_dir' in locals():
            shutil.rmtree(temp_dir)
        raise
    
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