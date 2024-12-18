import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

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
        
        # Output file
        self.output_path = tk.StringVar()
        self.output_label = ttk.Label(
            self.file_frame,
            text="Guardar como:",
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
            text="Guardar como",
            command=self.select_output_file
        )
        self.output_button.pack(side=tk.RIGHT)
        
        # Botón de procesar
        self.process_button = ttk.Button(
            self.main_frame,
            text="Procesar Excel",
            command=self.process_file,
            style='Accent.TButton'
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
            # Sugerir nombre de archivo de salida
            suggested_output = os.path.join(
                os.path.dirname(filename),
                "procesado_" + os.path.basename(filename)
            )
            self.output_path.set(suggested_output)

    def select_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Guardar como",
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx"
        )
        if filename:
            self.output_path.set(filename)

    def process_file(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror(
                "Error",
                "Por favor seleccione los archivos de entrada y salida"
            )
            return
        
        self.progress.start()
        self.status_label.config(text="Procesando archivo...")
        self.process_button.state(['disabled'])
        
        # Ejecutar el procesamiento en un hilo separado para no bloquear la UI
        self.root.after(100, self.run_processing)

    def run_processing(self):
        try:
            process_excel(self.input_path.get(), self.output_path.get())
            self.progress.stop()
            self.status_label.config(text="¡Archivo procesado correctamente!")
            messagebox.showinfo(
                "Éxito",
                "El archivo ha sido procesado correctamente"
            )
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Error al procesar el archivo")
            messagebox.showerror("Error", str(e))
        finally:
            self.process_button.state(['!disabled'])
            self.progress.stop()

# Mantener las funciones auxiliares existentes
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
    keywords = ['CTNS', 'MARCA', 'CBM', 'WEIGHT', 'PRODUCTO']
    for idx, row in df.iterrows():
        row_str = ' '.join(str(val).upper().strip() for val in row)
        matches = sum(keyword in row_str for keyword in keywords)
        if matches >= 2:
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

def process_excel(input_path, output_path):
    # La función process_excel se mantiene igual que en tu código original
    # Solo eliminamos la parte relacionada con Django
    try:
        df = pd.read_excel(input_path, header=None)
        header_row = find_header_row(df)

        if header_row is None:
            raise ValueError("No se encontró la fila de encabezados adecuada.")

        df = pd.read_excel(input_path, header=header_row)
        df.columns = df.columns.str.strip().str.upper()

        ctns_col = find_column(df.columns, 'CTNS')
        cbm_col = find_column(df.columns, 'CBM')
        weight_col = find_column(df.columns, 'WEIGHT')
        marca_col = find_column(df.columns, 'MARCA')

        df[ctns_col] = df[ctns_col].apply(clean_numeric_value)
        df[cbm_col] = df[cbm_col].apply(clean_numeric_value)
        df[weight_col] = df[weight_col].apply(clean_numeric_value)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Datos')

            workbook = writer.book
            worksheet = writer.sheets['Datos']

            suma_columns = {
                'CTNS': ctns_col,
                'T/CBM': cbm_col,
                'T/WEIGHT (KG)': weight_col
            }
            for column_name, full_column_name in suma_columns.items():
                col_index = list(df.columns).index(full_column_name) + 1
                column_letter = get_column_letter(col_index)
                num_rows = len(df) + 1
                formula = f'=SUM({column_letter}2:{column_letter}{num_rows-1})'
                sum_cell = f'{column_letter}{num_rows + 1}'
                worksheet[sum_cell] = formula
                worksheet[sum_cell].font = Font(bold=True)
                worksheet[sum_cell].alignment = Alignment(horizontal='right')

            for row in range(2, len(df) + 2):
                marca = worksheet.cell(row=row, column=list(df.columns).index(marca_col) + 1).value
                color_hex = get_color_by_brand(marca)
                fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                for col in range(1, len(df.columns) + 1):
                    worksheet.cell(row=row, column=col).fill = fill

            summary = df.groupby(marca_col).agg({
                ctns_col: 'sum',
                cbm_col: 'sum',
                weight_col: 'sum'
            }).reset_index()

            summary.to_excel(writer, sheet_name='Resumen por Marca', index=False)

        return df

    except Exception as e:
        raise ValueError(f"Error al procesar el archivo: {str(e)}")

def main():
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()