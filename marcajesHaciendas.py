import pandas as pd
from datetime import datetime
import subprocess
import sys
import base64
import os

POWERSHELL = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
TEMP_WSL = "/mnt/c/Users/RafaelO/AppData/Local/Temp/marcajes_temp.xlsx"


def _win_a_wsl(ruta_win):
    """Convierte 'E:\\carpeta\\archivo.csv' a '/mnt/e/carpeta/archivo.csv'."""
    ruta_win = ruta_win.strip()
    if len(ruta_win) >= 2 and ruta_win[1] == ":":
        unidad = ruta_win[0].lower()
        resto = ruta_win[2:].replace("\\", "/")
        return f"/mnt/{unidad}{resto}"
    return None


def _wsl_a_win(ruta_wsl):
    """Convierte '/mnt/c/Users/...' a 'C:\\Users\\...'."""
    if ruta_wsl.startswith("/mnt/") and len(ruta_wsl) > 6:
        unidad = ruta_wsl[5].upper()
        resto = ruta_wsl[6:].replace("/", "\\")
        return f"{unidad}:{resto}"
    return None


def _run_ps(script):
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    r = subprocess.run(
        [POWERSHELL, "-NoProfile", "-STA", "-EncodedCommand", encoded],
        capture_output=True, text=True,
        encoding="utf-8", errors="replace"
    )
    return r.stdout.strip()


def dialogo_abrir_windows(titulo="Selecciona archivo",
                          filtro="CSV files (*.csv)|*.csv|All files (*.*)|*.*"):
    ps = f'''
Add-Type -AssemblyName System.Windows.Forms
$f = New-Object System.Windows.Forms.OpenFileDialog
$f.Filter = "{filtro}"
$f.Title = "{titulo}"
if ($f.ShowDialog() -eq "OK") {{ Write-Output $f.FileName }}
'''
    ruta_win = _run_ps(ps)
    print(f"DEBUG abrir -> {ruta_win!r}")
    if not ruta_win:
        return None
    if ruta_win.startswith("\\\\"):
        print("Ruta UNC para abrir: se usara tal cual via /mnt temporal")
        return None
    return _win_a_wsl(ruta_win)


def dialogo_guardar_windows(titulo="Guardar archivo",
                            filtro="Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                            nombre_def="marcajes_filtrados.xlsx"):
    ps = f'''
Add-Type -AssemblyName System.Windows.Forms
$f = New-Object System.Windows.Forms.SaveFileDialog
$f.Filter = "{filtro}"
$f.Title = "{titulo}"
$f.FileName = "{nombre_def}"
if ($f.ShowDialog() -eq "OK") {{ Write-Output $f.FileName }}
'''
    return _run_ps(ps)


def pedir_fecha_windows(titulo="Fecha de corte",
                        mensaje="Introduce la fecha (YYYY-MM-DD):"):
    ps = f'''
Add-Type -AssemblyName Microsoft.VisualBasic
Write-Output ([Microsoft.VisualBasic.Interaction]::InputBox("{mensaje}", "{titulo}", ""))
'''
    return _run_ps(ps)


def guardar_excel(df, ruta_win):
    """Guarda el Excel en la ruta Windows indicada (soporta UNC y locales)."""
    if ruta_win.startswith("\\\\"):
        # Ruta UNC: guardamos en temp y copiamos con PowerShell
        print(f"   Ruta UNC detectada, guardando en temp y copiando...")
        df.to_excel(TEMP_WSL, index=False)
        src_win = _wsl_a_win(TEMP_WSL)
        ps = f'Copy-Item -Path "{src_win}" -Destination "{ruta_win}" -Force'
        _run_ps(ps)
        try:
            os.remove(TEMP_WSL)
        except Exception:
            pass
        return True
    else:
        # Ruta local (C:, D:, E:, Z:, etc.): conversion directa
        ruta_wsl = _win_a_wsl(ruta_win)
        if not ruta_wsl:
            print(f"Ruta no reconocida: {ruta_win}")
            return False
        df.to_excel(ruta_wsl, index=False)
        return True


# ---------- FLUJO PRINCIPAL ----------

print("--> Abriendo dialogo para seleccionar CSV...")
ruta_csv = dialogo_abrir_windows("Selecciona el archivo CSV de marcajes")
print(f"DEBUG ruta CSV (WSL): {ruta_csv!r}")
if not ruta_csv:
    print("No se selecciono ningun archivo.")
    sys.exit()

fecha_input = pedir_fecha_windows()
print(f"DEBUG fecha ingresada: {fecha_input!r}")
try:
    fecha_corte = datetime.strptime(fecha_input, "%Y-%m-%d")
except Exception:
    print("Fecha invalida.")
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
print("Registros despues del filtro:", len(df_filtrado))
print("Registros finales:", len(df_final))

print("--> Abriendo dialogo para guardar Excel...")
ruta_excel_win = dialogo_guardar_windows()
print(f"DEBUG ruta Excel (Windows): {ruta_excel_win!r}")
if not ruta_excel_win:
    print("No se selecciono ubicacion para guardar.")
    sys.exit()

print("--> Guardando Excel...")
if guardar_excel(df_final, ruta_excel_win):
    print(f"OK - Archivo guardado en: {ruta_excel_win}")
else:
    print("ERROR al guardar el archivo.")
    sys.exit(1)
