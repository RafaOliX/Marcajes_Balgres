import pandas as pd
from datetime import datetime
import subprocess
import sys

def dialogo_abrir_windows(titulo="Selecciona archivo", filtro="CSV files (*.csv)|*.csv|All files (*.*)|*.*"):
    """Abre el diálogo nativo de Windows y devuelve la ruta en formato WSL."""
    ps = f'''
    Add-Type -AssemblyName System.Windows.Forms
    $f = New-Object System.Windows.Forms.OpenFileDialog
    $f.Filter = "{filtro}"
    $f.Title = "{titulo}"
    if ($f.ShowDialog() -eq "OK") {{ $f.FileName }} else {{ "" }}
    '''
    r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    ruta_win = r.stdout.strip()
    if not ruta_win:
        return None
    # Convertir C:\... a /mnt/c/...
    ruta_wsl = subprocess.run(["wslpath", "-u", ruta_win],
                              capture_output=True, text=True).stdout.strip()
    return ruta_wsl

def dialogo_guardar_windows(titulo="Guardar archivo", filtro="Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*", nombre_def="marcajes_filtrados.xlsx"):
    """Abre el diálogo Guardar Como nativo de Windows y devuelve la ruta en formato WSL."""
    ps = f'''
    Add-Type -AssemblyName System.Windows.Forms
    $f = New-Object System.Windows.Forms.SaveFileDialog
    $f.Filter = "{filtro}"
    $f.Title = "{titulo}"
    $f.FileName = "{nombre_def}"
    if ($f.ShowDialog() -eq "OK") {{ $f.FileName }} else {{ "" }}
    '''
    r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    ruta_win = r.stdout.strip()
    if not ruta_win:
        return None
    ruta_wsl = subprocess.run(["wslpath", "-u", ruta_win],
                              capture_output=True, text=True).stdout.strip()
    return ruta_wsl

def pedir_fecha_windows(titulo="Fecha de corte", mensaje="Introduce la fecha (YYYY-MM-DD):"):
    """Pide un texto con un inputbox nativo de Windows."""
    ps = f'''
    Add-Type -AssemblyName Microsoft.VisualBasic
    $r = [Microsoft.VisualBasic.Interaction]::InputBox("{mensaje}", "{titulo}", "")
    $r
    '''
    r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    return r.stdout.strip()

# ---------- FLUJO PRINCIPAL ----------

ruta_csv = dialogo_abrir_windows("Selecciona el archivo CSV de marcajes")
if not ruta_csv:
    print("❌ No se seleccionó ningún archivo.")
    sys.exit()

fecha_input = pedir_fecha_windows()
try:
    fecha_corte = datetime.strptime(fecha_input, "%Y-%m-%d")
except:
    print("❌ Fecha inválida.")
    sys.exit()

df = pd.read_csv(ruta_csv)
df['sJobNo'] = df['sJobNo'].astype(str).str.replace("'", "").str.replace(",", "")
df['Date'] = pd.to_datetime(df['Date'], format="%Y-%m-%d", errors='coerce')
df_filtrado = df[df['Date'] >= fecha_corte]
df_filtrado = df_filtrado.drop_duplicates(subset=['Date', 'Time'])
df_filtrado['Date'] = df_filtrado['Date'].dt.strftime('%Y-%m-%d')

columnas_deseadas = ['sName', 'sJobNo', 'Date', 'Time', 'IN/OUT', 'AttendanceStatus']
df_final = df_filtrado[columnas_deseadas]

print("Registros originales:", len(df))
print("Registros después del filtro:", len(df_filtrado))
print("Registros finales:", len(df_final))

ruta_excel = dialogo_guardar_windows()
if not ruta_excel:
    print("❌ No se seleccionó ubicación para guardar.")
    sys.exit()

df_final.to_excel(ruta_excel, index=False)
print(f"✅ Archivo guardado exitosamente en: {ruta_excel}")