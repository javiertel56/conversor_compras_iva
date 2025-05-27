import pandas as pd

def limpiar_valor(valor):
    if pd.isna(valor):
        return 0.0
    try:
        return float(str(valor).replace(',', ''))
    except:
        return 0.0

def procesar_archivo(origen, destino):
    # Leer el archivo original (puede requerir ajustar skiprows y sheet_name)
    df = pd.read_excel(origen, header=None)
    
    # Buscar filas que contienen movimientos (ejemplo: filas con número y cuenta)
    movimientos = []
    for idx, row in df.iterrows():
        # Detectar filas de movimientos por heurística (ajusta según tu archivo)
        if pd.notna(row[0]) and pd.notna(row[2]) and isinstance(row[2], str) and '-' in row[2]:
            numero = row[0]
            referencia = row[1]
            cuenta = row[2]
            nombre = row[3]
            concepto = row[4] if pd.notna(row[4]) else nombre
            cargo = limpiar_valor(row[6])
            abono = limpiar_valor(row[7])
            # Puedes agregar más columnas según el formato destino
            movimientos.append([
                numero, referencia, cuenta, concepto, cargo, abono, 0.0, 0.0
            ])
    
    # Crear DataFrame de salida
    columnas = ['No.', 'Refer.', 'Cuenta', 'Nombre', 'Cargos', 'Abonos', 'ColExtra1', 'ColExtra2']
    df_out = pd.DataFrame(movimientos, columns=columnas)
    
    # Guardar el archivo convertido
    df_out.to_excel(destino, index=False)

# Uso:
# procesar_archivo('origen.xlsx', 'destino.xlsx')