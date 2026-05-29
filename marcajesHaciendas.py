import pandas as pd
from tkinter import Tk, filedialog, simpledialog
from datetime import datetime

root = Tk()
root.withdraw()

ruta_csv = filedialog.askopenfilename(
    title="Selecciona el archivo CSV de marcajes",
    filetypes=[("CSV files", "*.csv")]
)
if not ruta_csv:
    print("❌ No se seleccionó ningún archivo.")
    exit()

fecha_input = simpledialog.askstring("Fecha de corte", "Introduce la fecha (YYYY-MM-DD):")
try:
    fecha_corte = datetime.strptime(fecha_input, "%Y-%m-%d")
except:
    print("❌ Fecha inválida.")
    exit()

# Leer CSV con encabezados
df = pd.read_csv(ruta_csv)

# Limpiar sJobNo
df['sJobNo'] = df['sJobNo'].astype(str).str.replace("'", "").str.replace(",", "")

# Convertir y filtrar por fecha
df['Date'] = pd.to_datetime(df['Date'], format="%Y-%m-%d", errors='coerce')
df_filtrado = df[df['Date'] >= fecha_corte]

# Eliminar duplicados
df_filtrado = df_filtrado.drop_duplicates(subset=['Date', 'Time'])

# Convertir fecha a texto plano
df_filtrado['Date'] = df_filtrado['Date'].dt.strftime('%Y-%m-%d')

# Mantener columnas deseadas
columnas_deseadas = ['sName', 'sJobNo', 'Date', 'Time', 'IN/OUT', 'AttendanceStatus']
df_final = df_filtrado[columnas_deseadas]

# Debug
print("Registros originales:", len(df))
print("Registros después del filtro:", len(df_filtrado))
print("Registros finales:", len(df_final))
print("Columnas disponibles:", df.columns.tolist())

# Guardar Excel
ruta_excel = filedialog.asksaveasfilename(
    title="Guardar archivo Excel",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)
if not ruta_excel:
    print("❌ No se seleccionó ubicación para guardar.")
    exit()

df_final.to_excel(ruta_excel, index=False)
print(f"✅ Archivo guardado exitosamente en: {ruta_excel}")
