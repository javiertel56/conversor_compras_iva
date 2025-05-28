import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Color, numbers

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Inventario Tiendas")
        self.geometry("700x500")
        self.configure(bg="#1c2c47")
        self.selected_file = None
        self.create_tabs()

    def create_tabs(self):
        style = ttk.Style(self)
        style.theme_use('clam')

        dark_bg = "#1c2c47"
        dark_fg = "#f5f6fa"
        tab_bg = "#1c2c47"
        tab_active = "#1c2c47"

        style.configure('TNotebook', background=dark_bg, borderwidth=0)
        style.configure('TNotebook.Tab', background=tab_bg, foreground=dark_fg, font=('Segoe UI', 12, 'bold'), padding=[10, 5])
        style.map('TNotebook.Tab',
                  background=[('selected', tab_active)],
                  foreground=[('selected', '#00b894')])

        style.configure('TFrame', background=dark_bg)
        style.configure('TLabel', background=dark_bg, foreground=dark_fg, font=('Segoe UI', 10))

        style.configure('Blue.TButton',
                        background='#2980b9',
                        foreground='white',
                        borderwidth=0,
                        font=('Segoe UI', 11, 'bold'),
                        padding=10,
                        relief='flat')
        style.map('Blue.TButton',
                  background=[('active', '#1c5980')],
                  foreground=[('active', 'white')])

        style.configure('Yellow.TButton',
                        background='#fdcb6e',
                        foreground='#1c2c47',
                        borderwidth=0,
                        font=('Segoe UI', 11, 'bold'),
                        padding=10,
                        relief='flat')
        style.map('Yellow.TButton',
                  background=[('active', '#e1b84b')],
                  foreground=[('active', '#1c2c47')])

        style.configure('Green.TButton',
                        background='#00b894',
                        foreground='white',
                        borderwidth=0,
                        font=('Segoe UI', 11, 'bold'),
                        padding=10,
                        relief='flat')
        style.map('Green.TButton',
                  background=[('active', '#00916e')],
                  foreground=[('active', 'white')])

        notebook = ttk.Notebook(self)
        notebook.pack(expand=1, fill='both', padx=20, pady=20)

        rm_frame = ttk.Frame(notebook)
        tco_frame = ttk.Frame(notebook)

        notebook.add(rm_frame, text='Rosa Marcela')
        notebook.add(tco_frame, text='Tcomunicamos')

        self.create_tab_content(rm_frame, "Tabulador iva - Rosa Marcela")
        self.create_tab_content(tco_frame, "Tabulador iva - Tcomunicamos")

    def create_tab_content(self, frame, titulo):
        lbl_titulo = ttk.Label(frame, text=titulo, background="#1c2c47", foreground="#f5f6fa", font=('Segoe UI', 16, 'bold'))
        lbl_titulo.pack(pady=(10, 20))

        btn_frame = tk.Frame(frame, bg="#1c2c47")
        btn_frame.pack(pady=10)

        lbl_file = ttk.Label(frame, text="Archivo seleccionado: Ninguno", background="#1c2c47", foreground="#f5f6fa", font=('Segoe UI', 10))
        lbl_file.pack(pady=(10, 20))

        def subir_archivo():
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            if file_path:
                self.selected_file = file_path
                lbl_file.config(text=f"Archivo seleccionado: {file_path}")

        def procesar_excel():
            if not self.selected_file:
                messagebox.showwarning("Advertencia", "Primero selecciona un archivo.")
                return
            try:
                destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                if not destino:
                    return
                convertir_excel(self.selected_file, destino)
                self.archivo_procesado = destino  # Guarda la ruta del archivo procesado
                messagebox.showinfo("Éxito", "Archivo procesado y guardado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error: {e}")

        def abrir_archivo():
            # Abre el archivo procesado, no el original
            if hasattr(self, 'archivo_procesado') and self.archivo_procesado:
                import os
                os.startfile(self.archivo_procesado)
            else:
                messagebox.showwarning("Advertencia", "Primero procesa y guarda un archivo.")

        btn_subir = ttk.Button(btn_frame, text="Subir archivo", style="Blue.TButton", command=subir_archivo)
        btn_subir.pack(side=tk.LEFT, padx=10)

        btn_procesar = ttk.Button(btn_frame, text="Procesar Excel", style="Yellow.TButton", command=procesar_excel)
        btn_procesar.pack(side=tk.LEFT, padx=10)

        btn_abrir = ttk.Button(btn_frame, text="Abrir archivo", style="Green.TButton", command=abrir_archivo)
        btn_abrir.pack(side=tk.LEFT, padx=10)

def limpiar_valor(valor):
    if pd.isna(valor):
        return 0.0
    try:
        return float(str(valor).replace(',', '').replace(' ', ''))
    except Exception:
        return 0.0

def ajustar_formato_excel(ruta_archivo):
    wb = openpyxl.load_workbook(ruta_archivo)
    azul = PatternFill(start_color="2980b9", end_color="2980b9", fill_type="solid")
    # blanco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Ya no se usa
    for ws in wb.worksheets:
        # Ajustar ancho de columnas
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

        # Encabezado azul, negrita y centrado
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
            cell.fill = azul

        # Congelar la primera fila
        ws.freeze_panes = "A2"

        # Formato numérico y pintar de blanco los ceros en columnas de dinero e IVA
        header = [cell.value for cell in ws[1]]
        cols_dinero = []
        for nombre in ['Cargo 16', 'Abono 16', 'Cargo 8', 'Abono 8']:
            if nombre in header:
                cols_dinero.append(header.index(nombre) + 1)
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
            for idx in cols_dinero:
                cell = row[idx-1]
                cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                if cell.value == 0 or cell.value == 0.0:
                    cell.font = Font(color="FFFFFF")  # Solo texto blanco
                    # cell.fill = blanco  # Elimina o comenta esta línea

    wb.save(ruta_archivo)

def convertir_excel(origen, destino):
    # Lee el archivo original
    df = pd.read_excel(origen, header=None)
    movimientos = []
    hoja3 = []
    agrupados = {}

    i = 0
    while i < len(df):
        row = df.iloc[i]
        # Detecta fila de movimiento principal (tiene número, referencia y cuenta)
        if pd.notna(row[0]) and pd.notna(row[2]) and isinstance(row[2], str) and '-' in row[2]:
            numero = row[0]
            referencia = row[1]
            cuenta = row[2]
            nombre = row[3] if pd.notna(row[3]) else ""
            concepto = ""
            cargo = limpiar_valor(row[6]) if len(row) > 6 else 0.0
            abono = limpiar_valor(row[7]) if len(row) > 7 else 0.0

            # Busca el concepto en la siguiente fila si la columna de cuenta está vacía
            if i + 1 < len(df):
                next_row = df.iloc[i + 1]
                if pd.isna(next_row[2]) and pd.notna(next_row[4]):
                    concepto = str(next_row[4])
                    i += 1  # Salta la fila de concepto
                else:
                    concepto = nombre
            else:
                concepto = nombre

            # --- NUEVO: Busca la fila de IVA asociada (cuenta inicia con 1104-) ---
            abono_iva_16 = 0.0
            abono_iva_8 = 0.0
            j = i + 1
            while j < len(df):
                iva_row = df.iloc[j]
                if pd.notna(iva_row[2]) and isinstance(iva_row[2], str) and iva_row[2].startswith("1104-"):
                    iva_valor = limpiar_valor(iva_row[6]) if len(iva_row) > 6 else 0.0
                    # Solo toma el primer IVA que encuentre, sea 16 o 8, y sale del ciclo
                    if iva_row[2].endswith("-01"):
                        abono_iva_16 = iva_valor
                        abono_iva_8 = 0.0
                    elif iva_row[2].endswith("-02"):
                        abono_iva_8 = iva_valor
                        abono_iva_16 = 0.0
                    break
                j += 1

            # Agrupa por referencia y concepto (solo para la cuenta principal, ignora otras cuentas)
            clave = (referencia, concepto)
            if clave not in agrupados:
                agrupados[clave] = {
                    "numero": numero,
                    "referencia": referencia,
                    "cuenta": cuenta,
                    "concepto": concepto,
                    "nombre": nombre,
                    "cargos": 0.0,
                    "abonos": 0.0,
                    "abono_iva_16": 0.0,
                    "abono_iva_8": 0.0
                }
            agrupados[clave]["cargos"] += cargo
            agrupados[clave]["abonos"] += abono
            agrupados[clave]["abono_iva_16"] += abono_iva_16
            agrupados[clave]["abono_iva_8"] += abono_iva_8
        i += 1

    # Antes de procesar movimientos, crea un set con referencias que tienen IVA 8%
    referencias_iva_8 = set()
    for v in agrupados.values():
        cuenta_str = str(v["cuenta"]).lower() if v["cuenta"] else ""
        if cuenta_str.startswith("1104-") and cuenta_str.endswith("-02"):
            referencias_iva_8.add(v["referencia"])

    # Procesa movimientos para hojas
    for v in agrupados.values():
        referencia_str = str(v["referencia"]).lower() if v["referencia"] else ""
        concepto_str = str(v["concepto"]).lower() if v["concepto"] else ""
        nombre_str = str(v["nombre"]).lower() if v["nombre"] else ""
        cuenta_str = str(v["cuenta"]).lower() if v["cuenta"] else ""
        if (
            "finiquito" in referencia_str or
            "nomina" in referencia_str or
            "sueldos y salarios" in concepto_str or
            "sueldos y salarios" in nombre_str or
            "imss" in concepto_str or
            "imss" in nombre_str or
            "comprobante sin requisitos" in concepto_str or
            "comprobante sin requisitos" in nombre_str
        ):
            continue

        cargo_16 = abono_16 = cargo_8 = abono_8 = 0.0

        if cuenta_str.startswith("1104-"):
            if cuenta_str.endswith("-01"):
                cargo_16 = abs(v["cargos"])
                abono_16 = abs(v["abonos"])
            elif cuenta_str.endswith("-02"):
                cargo_8 = abs(v["cargos"])
                abono_8 = abs(v["abonos"])
        else:
            cargo_16 = abs(v["cargos"]) if v.get("abono_iva_16", 0.0) > 0 else 0.0
            abono_16 = abs(v.get("abono_iva_16", 0.0))
            cargo_8 = abs(v["cargos"]) if v.get("abono_iva_8", 0.0) > 0 else 0.0
            abono_8 = abs(v.get("abono_iva_8", 0.0))

        # Redondear a dos decimales para mostrar en Excel
        cargo_16 = round(cargo_16, 2)
        abono_16 = round(abono_16, 2)
        cargo_8 = round(cargo_8, 2)
        abono_8 = round(abono_8, 2)

        # Fórmulas Excel (referencias de columna: F=6, G=7, H=8, I=9)
        # Ejemplo: =F2*0.16-G2
        fila = [
            v["numero"], v["referencia"], v["cuenta"], v["concepto"], v["nombre"],
            cargo_16, abono_16, cargo_8, abono_8,
            None, None  # Las fórmulas se agregan después
        ]
        if v["numero"] == 2:
            hoja3.append(fila)
        if (
            v["cargos"] > 0 and
            isinstance(v["cuenta"], str) and (
                v["cuenta"].startswith("1102-") or
                v["cuenta"].startswith("5100-") or
                v["cuenta"].startswith("5200-") or
                v["cuenta"].startswith("61000-") or
                v["cuenta"].startswith("1201-")
            )
        ):
            movimientos.append(fila)

    columnas = [
        'No.', 'Refer.', 'Cuenta', 'Concepto', 'Nombre',
        'Cargo 16', 'Abono 16', 'Cargo 8', 'Abono 8',
        'Fórmula 16', 'Fórmula 8'
    ]
    with pd.ExcelWriter(destino, engine='openpyxl') as writer:
        # Hoja 1: Excel original, pero separado por líneas de movimientos
        df.to_excel(writer, index=False, header=False, sheet_name='Hoja1')
        # Hoja 2: Lo que era Sheet1 (movimientos procesados)
        df_mov = pd.DataFrame(movimientos, columns=columnas)
        df_mov.to_excel(writer, index=False, sheet_name='Hoja2')
        # Hoja 3: Lo que era Sheet3 (hoja3)
        if hoja3:
            df_hoja3 = pd.DataFrame(hoja3, columns=columnas)
            df_hoja3.to_excel(writer, index=False, sheet_name='Hoja3')

    # Formato y fórmulas solo para Hoja2 y Hoja3
    wb = openpyxl.load_workbook(destino)
    gris_suave = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # Gris claro
    pastel = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")      # Verde pastel suave

    for sheet_name in ['Hoja2', 'Hoja3']:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_row = ws.max_row
            # Formato y fórmulas por fila
            for row in range(2, max_row + 1):
                # Fórmula 16: =F2*0.16-G2
                ws[f'J{row}'] = f"=F{row}*0.16-G{row}"
                ws[f'J{row}'].number_format = '0.00'
                ws[f'J{row}'].fill = gris_suave
                # Fórmula 8: =H2*0.08-I2
                ws[f'K{row}'] = f"=H{row}*0.08-I{row}"
                ws[f'K{row}'].number_format = '0.00'
                ws[f'K{row}'].fill = gris_suave
                # Formato de dos decimales y colores en columnas de dinero
                for col in ['F', 'G', 'H', 'I']:
                    cell = ws[f'{col}{row}']
                    cell.number_format = '0.00'
                    if cell.value == 0 or cell.value == "0.00" or cell.value == 0.0:
                        cell.font = Font(color="FFFFFF")  # blanco
                    elif cell.value is not None and cell.value > 0:
                        cell.fill = pastel

            # Agregar suma al final de cada columna de dinero
            suma_row = max_row + 1
            ws[f'E{suma_row}'] = "TOTAL"
            ws[f'E{suma_row}'].font = Font(bold=True)
            for idx, col in enumerate(['F', 'G', 'H', 'I'], start=6):
                ws[f'{chr(64+idx)}{suma_row}'] = f"=SUM({chr(64+idx)}2:{chr(64+idx)}{max_row})"
                ws[f'{chr(64+idx)}{suma_row}'].number_format = '0.00'

            # Pintar la fila de totales de fórmulas de gris suave
            ws[f'J{suma_row}'].fill = gris_suave
            ws[f'K{suma_row}'].fill = gris_suave

    wb.save(destino)
    ajustar_formato_excel(destino)

if __name__ == "__main__":
    app = App()
    app.mainloop()