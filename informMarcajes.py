import pandas as pd
import os
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Para abrir el explorador de archivos
import tkinter as tk
from tkinter import filedialog

# Para formato Excel
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

def verificar_dependencias():
    """Verifica e instala dependencias necesarias"""
    dependencias = {
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'xlrd': 'xlrd>=2.0.1'
    }
    
    print("🔍 Verificando dependencias...")
    
    for nombre, paquete in dependencias.items():
        try:
            if nombre == 'pandas':
                import pandas
                print(f"✅ {nombre}: OK")
            elif nombre == 'openpyxl':
                import openpyxl
                print(f"✅ {nombre}: OK")
            elif nombre == 'xlrd':
                import xlrd
                version = xlrd.__version__
                print(f"✅ {nombre}: v{version}")
        except ImportError:
            print(f"⚠️ {nombre}: FALTANTE")
            
            try:
                print(f"   Instalando {paquete}...")
                import subprocess
                import sys
                subprocess.check_call([sys.executable, "-m", "pip", "install", paquete])
                print(f"   ✅ {nombre} instalado correctamente")
            except:
                print(f"   ❌ No se pudo instalar {nombre}")
                print(f"   Por favor, instálelo manualmente:")
                print(f"   pip install {paquete}")
                return False
    
    print("✅ Todas las dependencias están instaladas\n")
    return True

def mostrar_banner():
    """Muestra el banner del sistema"""
    print("\n" + "="*60)
    print("      SISTEMA DE UNIFICACIÓN DE MARCAJE POR PERIODO")
    print("="*60)

def mostrar_menu():
    """Muestra el menú de opciones depurado"""
    print("\nOPCIONES DISPONIBLES:")
    print("1. Unificar periodo entre fechas (puede cruzar meses)")
    print("2. Salir")

def seleccionar_archivos_explorador(titulo="Seleccionar archivos", multiple=True):
    """Abre el explorador de Windows para seleccionar archivos - MODIFICADO para mostrar TODOS"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # Definir tipos de archivo - PRIMERO TODOS LOS ARCHIVOS
    filetypes = [
        ("Todos los archivos", "*.*"),  # ¡ESTO ES CLAVE! Va primero
        ("Archivos Excel (.xlsx)", "*.xlsx"),
        ("Archivos Excel antiguos (.xls)", "*.xls"),
        ("Archivos CSV (.csv)", "*.csv"),
        ("Archivos de texto (.txt)", "*.txt"),
        ("Archivos PDF (.pdf)", "*.pdf")
    ]
    
    if multiple:
        archivos = filedialog.askopenfilenames(
            title=titulo,
            filetypes=filetypes,
            initialdir=os.getcwd()
        )
        return list(archivos)
    else:
        archivo = filedialog.askopenfilename(
            title=titulo,
            filetypes=filetypes,
            initialdir=os.getcwd()
        )
        return [archivo] if archivo else []

def seleccionar_archivos_uno_por_uno():
    """Permite seleccionar archivos uno por uno"""
    print("\n" + "="*50)
    print("SELECCIÓN DE ARCHIVOS UNO POR UNO")
    print("="*50)
    print("\n📁 Se abrirá el explorador de archivos para cada selección.")
    print("   En el explorador, en 'Tipo de archivos' selecciona:")
    print("   • 'Todos los archivos' para ver todo")
    print("   • O el tipo específico que necesites")
    print("   Selecciona un archivo y haz clic en 'Abrir'.")
    print("   Cuando termines, cancela o cierra el explorador.")
    print("-" * 50)
    
    archivos_seleccionados = []
    contador = 1
    
    while True:
        print(f"\n📂 Selección #{contador}")
        print("   (Cierra o cancela el explorador para terminar)")
        
        archivos = seleccionar_archivos_explorador(
            titulo=f"Selecciona el archivo #{contador} - Cierra para terminar",
            multiple=False
        )
        
        if not archivos:
            if contador == 1:
                print("❌ No se seleccionó ningún archivo")
                return []
            else:
                print(f"✓ Terminada la selección. Total: {len(archivos_seleccionados)} archivos")
                break
        
        archivo = archivos[0]
        nombre_archivo = os.path.basename(archivo)
        
        if archivo in archivos_seleccionados:
            print(f"⚠️ El archivo '{nombre_archivo}' ya fue seleccionado")
            continue
        
        archivos_seleccionados.append(archivo)
        print(f"✓ Agregado: {nombre_archivo}")
        contador += 1
    
    if archivos_seleccionados:
        print("\n📋 RESUMEN DE SELECCIÓN:")
        print("-" * 60)
        for i, archivo in enumerate(archivos_seleccionados, 1):
            nombre = os.path.basename(archivo)
            extension = os.path.splitext(nombre)[1].upper()
            if not extension:
                extension = "(sin extensión)"
            tamaño = os.path.getsize(archivo) / 1024
            print(f"{i:2}. {nombre:<45} {extension:<8} ({tamaño:.1f} KB)")
        print("-" * 60)
    
    return archivos_seleccionados

def seleccionar_archivos_multiples():
    """Permite seleccionar múltiples archivos a la vez"""
    print("\n" + "="*50)
    print("SELECCIÓN MÚLTIPLE DE ARCHIVOS")
    print("="*50)
    print("\n📌 En el explorador, selecciona 'Todos los archivos' para ver todo.")
    print("   O selecciona el tipo específico que necesites.")
    
    archivos = seleccionar_archivos_explorador(
        titulo="Selecciona los archivos a unificar (Ctrl para múltiples)",
        multiple=True
    )
    
    if archivos:
        print(f"\n✓ Seleccionados {len(archivos)} archivo(s):")
        print("-" * 60)
        
        # Contar por extensiones
        extensiones = {}
        for archivo in archivos:
            nombre = os.path.basename(archivo)
            ext = os.path.splitext(nombre)[1].lower()
            if not ext:
                ext = "(sin extensión)"
            extensiones[ext] = extensiones.get(ext, 0) + 1
        
        # Mostrar archivos
        for i, archivo in enumerate(archivos, 1):
            nombre = os.path.basename(archivo)
            extension = os.path.splitext(nombre)[1].upper()
            if not extension:
                extension = "(SIN EXT)"
            tamaño = os.path.getsize(archivo) / 1024
            print(f"{i:2}. {nombre:<40} {extension:<8} ({tamaño:.1f} KB)")
        
        print("-" * 60)
        
        # Mostrar resumen de extensiones
        if extensiones:
            print("\n📊 RESUMEN DE EXTENSIONES:")
            for ext, count in sorted(extensiones.items()):
                print(f"   • {ext.upper() if ext != '(sin extensión)' else ext}: {count} archivo(s)")
        
        print("-" * 60)
    
    return archivos

def formatear_horas_vertical(texto):
    """Formatea las horas para mostrarlas en líneas verticales"""
    if pd.isna(texto) or texto == "" or texto is None:
        return ""
    
    texto = str(texto).strip()
    
    # Si ya tiene saltos de línea, mantenerlos
    if '\n' in texto:
        return texto
    
    # Dividir por espacios
    partes = texto.split()
    horas_formateadas = []
    
    for parte in partes:
        parte = parte.strip()
        if parte:
            # Verificar si tiene formato de hora (XX:XX)
            if ':' in parte and len(parte) >= 4:
                horas_formateadas.append(parte)
            else:
                # Si no es hora, mantener el texto original
                horas_formateadas.append(parte)
    
    # Unir con saltos de línea
    return '\n'.join(horas_formateadas)

def limpiar_datos_marcaje(df):
    """Limpia y formatea los datos de marcaje con horas verticales"""
    if df is None or df.empty:
        return df
    
    df_limpio = df.copy()
    
    for col in df_limpio.columns:
        if df_limpio[col].dtype == 'object':
            # Primero limpiar caracteres extraños
            df_limpio[col] = df_limpio[col].astype(str).str.replace('\n', ' ', regex=True)
            df_limpio[col] = df_limpio[col].str.replace('\r', ' ', regex=True)
            df_limpio[col] = df_limpio[col].str.strip()
            df_limpio[col] = df_limpio[col].replace('nan', '', regex=False)
            df_limpio[col] = df_limpio[col].replace('None', '', regex=False)
            
            # Aplicar formato de horas verticales
            df_limpio[col] = df_limpio[col].apply(formatear_horas_vertical)
    
    return df_limpio

def leer_archivo_completo(ruta_archivo):
    """Lee un archivo completo de marcaje"""
    nombre = os.path.basename(ruta_archivo)
    extension = os.path.splitext(nombre)[1].lower()
    print(f"  Leyendo: {nombre} ({extension if extension else 'sin extensión'})")
    
    try:
        # Determinar el engine basado en la extensión
        if extension == '.xls':
            engine = 'xlrd'
        elif extension == '.xlsx':
            engine = 'openpyxl'
        else:
            print(f"    ⚠️ Extensión no soportada: {extension}")
            print(f"    ℹ️ Solo se soportan archivos .xls y .xlsx")
            return None
        
        # Buscar la fila donde empiezan los datos reales
        for skip_rows in range(0, 10):
            try:
                df_prueba = pd.read_excel(
                    ruta_archivo, 
                    skiprows=skip_rows, 
                    header=None, 
                    nrows=5,
                    engine=engine
                )
                
                # Buscar 'Employee ID' o similar
                for i in range(len(df_prueba)):
                    primera_celda = str(df_prueba.iloc[i, 0]).lower() if i < len(df_prueba) else ""
                    
                    if any(keyword in primera_celda for keyword in ['employee', 'empleado', 'id']):
                        df = pd.read_excel(
                            ruta_archivo, 
                            skiprows=skip_rows + i,
                            engine=engine
                        )
                        
                        print(f"    ✓ {len(df)} registros encontrados")
                        df = limpiar_datos_marcaje(df)
                        return df
            except Exception as e:
                continue
        
        # Si no encuentra, leer directamente
        df = pd.read_excel(ruta_archivo, engine=engine)
        print(f"    ✓ {len(df)} registros (lectura directa)")
        df = limpiar_datos_marcaje(df)
        return df
        
    except Exception as e:
        print(f"    ✗ Error: {str(e)}")
        return None

def combinar_datos_empleados(lista_dataframes):
    """
    Combina datos de empleados de múltiples dataframes en uno solo
    """
    if not lista_dataframes:
        return pd.DataFrame()
    
    # Empezar con el primer dataframe
    df_combinado = lista_dataframes[0].copy()
    
    # Si hay más dataframes, combinarlos
    for i in range(1, len(lista_dataframes)):
        df_actual = lista_dataframes[i]
        
        # Buscar columnas comunes (ID, Name, Department)
        col_id = None
        for col in df_combinado.columns:
            if 'id' in str(col).lower() and 'employee' in str(col).lower():
                col_id = col
                break
        
        if col_id is None:
            # Si no encuentra ID, usar la primera columna
            col_id = df_combinado.columns[0]
        
        # Para cada empleado en el dataframe actual
        for idx, row in df_actual.iterrows():
            emp_id = row[col_id] if col_id in row else None
            
            if emp_id:
                # Buscar si este empleado ya está en el combinado
                mask = df_combinado[col_id] == emp_id
                
                if mask.any():
                    # Empleado existe, actualizar datos
                    idx_existente = df_combinado[mask].index[0]
                    
                    # Actualizar columnas de días
                    for col in df_actual.columns:
                        if str(col).replace('.', '').isdigit():  # Es una columna de día
                            valor = row[col]
                            if pd.notna(valor) and str(valor).strip():
                                df_combinado.at[idx_existente, col] = valor
                else:
                    # Empleado nuevo, agregar fila
                    df_combinado = pd.concat([df_combinado, pd.DataFrame([row])], ignore_index=True)
    
    return df_combinado

def ordenar_datos_por_columna_a(df):
    """Ordena el DataFrame por la columna A (primera columna)"""
    if df is None or df.empty:
        return df
    
    df_ordenado = df.copy()
    
    # Identificar la columna para ordenar (columna A)
    columna_a = df_ordenado.columns[0]
    
    print(f"  🔍 Identificada columna para ordenar: '{columna_a}'")
    
    try:
        # Intentar ordenar numéricamente si es posible
        if df_ordenado[columna_a].dtype in ['int64', 'float64']:
            df_ordenado = df_ordenado.sort_values(by=columna_a, na_position='last')
            print(f"  ✓ Ordenado numéricamente por '{columna_a}'")
        else:
            # Convertir a string y ordenar
            df_ordenado[columna_a] = df_ordenado[columna_a].astype(str)
            
            # Intentar extraer números para ordenar
            def extract_number(x):
                try:
                    # Buscar números en el string
                    import re
                    numbers = re.findall(r'\d+', x)
                    if numbers:
                        return int(numbers[0])
                    return float('inf')  # Si no tiene número, va al final
                except:
                    return float('inf')
            
            # Agregar columna temporal para ordenar
            df_ordenado['_temp_sort'] = df_ordenado[columna_a].apply(extract_number)
            df_ordenado = df_ordenado.sort_values(by='_temp_sort', na_position='last')
            df_ordenado = df_ordenado.drop('_temp_sort', axis=1)
            
            # También ordenar alfabéticamente
            df_ordenado = df_ordenado.sort_values(by=columna_a, key=lambda col: col.str.lower())
            
            print(f"  ✓ Ordenado por '{columna_a}' (alfanuméricamente)")
    
    except Exception as e:
        print(f"  ⚠️ No se pudo ordenar por '{columna_a}': {str(e)}")
        print(f"  ℹ️ Se mantendrá el orden original")
    
    return df_ordenado

def aplicar_formato_excel_con_filtros(writer, df, sheet_name='Marcaje'):
    """Aplica formato y filtros al archivo Excel con ajuste de altura para horas"""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Aplicar filtros a todas las columnas
    worksheet.auto_filter.ref = worksheet.dimensions
    
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Formatear encabezados
    for col in range(1, len(df.columns) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Formatear celdas de datos con ajuste de altura para horas
    for row in range(2, len(df) + 2):
        # Ajustar altura de fila para mostrar horas verticales
        max_lineas = 1
        
        for col in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.border = thin_border
            
            # Configurar alineación y wrap text para todas las celdas
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            
            # Contar líneas si es texto con saltos de línea
            if cell.value and isinstance(cell.value, str):
                lineas = cell.value.count('\n') + 1
                if lineas > max_lineas:
                    max_lineas = lineas
        
        # Ajustar altura de fila según el contenido
        worksheet.row_dimensions[row].height = 15 * max_lineas  # 15 puntos por línea
    
    # Ajustar ancho de columnas automáticamente
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        
        # Ajustar ancho basado en el contenido
        for cell in column:
            try:
                if cell.value:
                    # Para celdas con múltiples líneas, tomar la línea más larga
                    if isinstance(cell.value, str) and '\n' in cell.value:
                        lineas = cell.value.split('\n')
                        longest_line = max(len(line) for line in lineas)
                        cell_length = longest_line
                    else:
                        cell_length = len(str(cell.value))
                    
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        
        # Ajustar ancho
        adjusted_width = max_length + 2
        
        if column_letter in ['A', 'B', 'C']:
            # Columnas fijas más anchas
            adjusted_width = min(max(adjusted_width, 12), 25)
        elif column[0].column > 3:
            # Columnas de días más estrechas
            adjusted_width = min(max(adjusted_width, 6), 10)
        
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Congelar paneles (primera fila y primeras 3 columnas)
    worksheet.freeze_panes = 'D2'
    
    print(f"  ✓ Filtros aplicados a todas las columnas")
    print(f"  ✓ Formato de horas vertical aplicado")
    print(f"  ✓ Altura de filas ajustada automáticamente")
    print(f"  ✓ Paneles congelados (fila 1 y columnas A-C)")
    
    return worksheet

def procesar_periodo_cruzado(lista_archivos, fecha_inicio, fecha_fin):
    """Procesa archivos de un periodo que cruza meses"""
    print(f"\n🔄 Procesando periodo: {fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}")
    print("-" * 60)
    
    # Filtrar solo archivos compatibles (.xls, .xlsx)
    archivos_compatibles = []
    archivos_no_compatibles = []
    
    for archivo in lista_archivos:
        if not os.path.exists(archivo):
            print(f"✗ Archivo no encontrado: {os.path.basename(archivo)}")
            continue
        
        nombre = os.path.basename(archivo)
        extension = os.path.splitext(archivo)[1].lower()
        
        if extension in ['.xls', '.xlsx']:
            archivos_compatibles.append(archivo)
        else:
            archivos_no_compatibles.append((nombre, extension))
    
    # Mostrar archivos no compatibles
    if archivos_no_compatibles:
        print(f"\n⚠️ Archivos no compatibles (ignorados):")
        for nombre, extension in archivos_no_compatibles:
            if not extension:
                extension = "(sin extensión)"
            print(f"   • {nombre} - Extensión: {extension}")
        print(f"   ℹ️ Solo se procesan archivos .xls y .xlsx")
    
    if not archivos_compatibles:
        print("❌ No hay archivos compatibles (.xls, .xlsx) para procesar")
        return None
    
    print(f"\n📁 Archivos compatibles encontrados: {len(archivos_compatibles)}")
    
    dataframes_por_mes = {}
    
    for archivo in archivos_compatibles:
        # Determinar mes del archivo
        nombre = os.path.basename(archivo).lower()
        mes = None
        
        meses_dict = {
            'ene': 1, 'jan': 1, '01': 1, '1': 1,
            'feb': 2, '02': 2, '2': 2,
            'mar': 3, '03': 3, '3': 3,
            'abr': 4, 'apr': 4, '04': 4, '4': 4,
            'may': 5, '05': 5, '5': 5,
            'jun': 6, '06': 6, '6': 6,
            'jul': 7, '07': 7, '7': 7,
            'ago': 8, 'aug': 8, '08': 8, '8': 8,
            'sep': 9, '09': 9, '9': 9,
            'oct': 10, '10': 10,
            'nov': 11, '11': 11,
            'dic': 12, 'dec': 12, '12': 12
        }
        
        for clave, valor in meses_dict.items():
            if clave in nombre:
                mes = valor
                break
        
        if mes is None:
            print(f"  ⚠️ No se pudo determinar el mes para: {os.path.basename(archivo)}")
            continue
        
        # Leer archivo
        df = leer_archivo_completo(archivo)
        
        if df is not None and not df.empty:
            # Filtrar solo los días del rango que corresponden a este mes
            dias_en_rango = []
            current_date = fecha_inicio
            
            while current_date <= fecha_fin:
                if current_date.month == mes:
                    dias_en_rango.append(current_date.day)
                current_date += timedelta(days=1)
            
            # Seleccionar columnas
            columnas_fijas = []
            for col in df.columns:
                col_str = str(col).lower()
                if any(keyword in col_str for keyword in ['employee', 'empleado', 'id', 'name', 'nombre', 'department', 'departamento']):
                    columnas_fijas.append(col)
            
            if len(columnas_fijas) < 2:
                columnas_fijas = df.columns[:3].tolist()
            
            columnas_dias = []
            for col in df.columns:
                try:
                    dia = int(float(str(col)))
                    if dia in dias_en_rango:
                        columnas_dias.append(col)
                except:
                    pass
            
            columnas_a_mantener = columnas_fijas + columnas_dias
            columnas_existentes = [col for col in columnas_a_mantener if col in df.columns]
            
            if columnas_existentes:
                df_filtrado = df[columnas_existentes].copy()
                
                # Normalizar nombres de columnas de días
                for col in df_filtrado.columns:
                    try:
                        dia = int(float(str(col)))
                        # Renombrar columna a solo el número
                        df_filtrado = df_filtrado.rename(columns={col: str(dia)})
                    except:
                        pass
                
                if mes not in dataframes_por_mes:
                    dataframes_por_mes[mes] = []
                
                dataframes_por_mes[mes].append(df_filtrado)
                print(f"    📅 Mes {mes}: {len(columnas_dias)} días")
            else:
                print(f"    ⚠️ No hay datos para este mes en el rango")
    
    # Combinar dataframes del mismo mes
    dataframes_combinados = {}
    for mes, dfs in dataframes_por_mes.items():
        if dfs:
            # Combinar empleados del mismo mes
            df_mes_combinado = combinar_datos_empleados(dfs)
            dataframes_combinados[mes] = df_mes_combinado
            print(f"  ✅ Mes {mes}: {len(df_mes_combinado)} empleados únicos")
    
    # Ahora combinar entre meses
    if not dataframes_combinados:
        print("❌ No se pudieron procesar los datos")
        return None
    
    # Tomar el primer dataframe como base
    meses = list(dataframes_combinados.keys())
    df_final = dataframes_combinados[meses[0]].copy()
    
    # Combinar con los otros meses
    for mes in meses[1:]:
        df_mes = dataframes_combinados[mes]
        
        # Buscar columna ID
        col_id = None
        for col in df_final.columns:
            if 'id' in str(col).lower():
                col_id = col
                break
        
        if col_id:
            # Para cada empleado en el nuevo mes
            for idx, row in df_mes.iterrows():
                emp_id = row[col_id] if col_id in row else None
                
                if emp_id:
                    # Buscar empleado en df_final
                    mask = df_final[col_id] == emp_id
                    
                    if mask.any():
                        # Empleado existe, agregar columnas de días
                        idx_existente = df_final[mask].index[0]
                        
                        for col in df_mes.columns:
                            if str(col).isdigit():  # Columna de día
                                valor = row[col]
                                if pd.notna(valor) and str(valor).strip():
                                    df_final.at[idx_existente, col] = valor
                    else:
                        # Empleado nuevo, agregar fila completa
                        df_final = pd.concat([df_final, pd.DataFrame([row])], ignore_index=True)
    
    # Ordenar columnas
    columnas_fijas = []
    columnas_dias = []
    
    for col in df_final.columns:
        if str(col).isdigit() and 1 <= int(col) <= 31:
            columnas_dias.append(int(col))
        else:
            columnas_fijas.append(col)
    
    # Ordenar días según el rango
    dias_ordenados = []
    current_date = fecha_inicio
    while current_date <= fecha_fin:
        dia = current_date.day
        if dia in columnas_dias:
            dias_ordenados.append(dia)
        current_date += timedelta(days=1)
    
    # Crear lista final de columnas
    columnas_finales = columnas_fijas + [str(dia) for dia in dias_ordenados]
    
    # Filtrar columnas que existen
    columnas_existentes = [col for col in columnas_finales if col in df_final.columns]
    df_final = df_final[columnas_existentes]
    
    # ORDENAR POR COLUMNA A
    print("\n🔀 Ordenando datos por columna A...")
    df_final = ordenar_datos_por_columna_a(df_final)
    
    return df_final

def unificar_periodo():
    """Función principal para unificar periodo entre fechas"""
    print("\n" + "="*60)
    print("UNIFICAR PERIODO ENTRE FECHAS")
    print("="*60)
    
    try:
        print("\n📅 Especifica el periodo a procesar:")
        print("Formato de fecha: DD/MM/AAAA")
        print("Ejemplo: 26/11/2025")
        print("-" * 30)
        
        fecha_inicio = input("Fecha de inicio (DD/MM/AAAA): ").strip()
        fecha_fin = input("Fecha de fin (DD/MM/AAAA): ").strip()
        
        try:
            inicio = datetime.strptime(fecha_inicio, "%d/%m/%Y")
            fin = datetime.strptime(fecha_fin, "%d/%m/%Y")
            
            if fin < inicio:
                print("❌ La fecha de fin debe ser posterior a la fecha de inicio")
                return
            
            dias = (fin - inicio).days + 1
            print(f"\n📆 Periodo: {fecha_inicio} al {fecha_fin} ({dias} días)")
            
        except ValueError:
            print("❌ Formato de fecha incorrecto. Use DD/MM/AAAA")
            return
        
        print("\n¿Cómo quieres seleccionar los archivos?")
        print("1. Uno por uno (recomendado)")
        print("2. Múltiples a la vez")
        print("3. Cancelar")
        
        seleccion = input("\n👉 Opción (1-3): ").strip()
        
        if seleccion == "3":
            print("❌ Operación cancelada")
            return
        
        if seleccion == "1":
            archivos = seleccionar_archivos_uno_por_uno()
        elif seleccion == "2":
            archivos = seleccionar_archivos_multiples()
        else:
            print("❌ Opción no válida")
            return
        
        if not archivos:
            print("❌ No se seleccionaron archivos")
            return
        
        # Procesar los archivos
        df_final = procesar_periodo_cruzado(archivos, inicio, fin)
        
        if df_final is not None and not df_final.empty:
            # Generar nombre de archivo
            inicio_str = inicio.strftime("%Y%m%d")
            fin_str = fin.strftime("%Y%m%d")
            nombre_salida = f"Marcaje_{inicio_str}_{fin_str}_Unificado.xlsx"
            
            # Crear carpeta de resultados si no existe
            if not os.path.exists('Resultados'):
                os.makedirs('Resultados')
            
            ruta_completa = os.path.join('Resultados', nombre_salida)
            
            # Guardar con formato y filtros
            with pd.ExcelWriter(ruta_completa, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Marcaje', index=False)
                
                # Aplicar formato y filtros
                aplicar_formato_excel_con_filtros(writer, df_final, 'Marcaje')
            
            print("\n" + "="*60)
            print("✅ PROCESAMIENTO COMPLETADO")
            print("="*60)
            print(f"📊 Total de empleados: {len(df_final)}")
            print(f"📁 Archivos procesados: {len(archivos)}")
            print(f"📅 Periodo: {fecha_inicio} - {fecha_fin}")
            print(f"🔢 Días mostrados: {len([c for c in df_final.columns if str(c).isdigit()])}")
            print(f"💾 Guardado como: {ruta_completa}")
            print("="*60)
            
            # Mostrar características del archivo generado
            print("\n🎯 CARACTERÍSTICAS DEL ARCHIVO GENERADO:")
            print("-" * 50)
            print("✓ Datos ordenados por la primera columna (Columna A)")
            print("✓ Filtros aplicados en TODAS las columnas")
            print("✓ Horas formateadas en líneas verticales")
            print("✓ Altura de filas ajustada automáticamente")
            print("✓ Paneles congelados (fila 1 + columnas A-C)")
            print("✓ Ancho de columnas optimizado")
            print("✓ Bordes en todas las celdas")
            print("-" * 50)
            
            # Mostrar vista previa con formato
            print("\n📋 VISTA PREVIA DEL FORMATO DE HORAS:")
            print("-" * 50)
            
            if len(df_final) > 0:
                # Buscar una celda con horas para mostrar ejemplo
                ejemplo_encontrado = False
                for idx, row in df_final.iterrows():
                    for col in df_final.columns:
                        if str(col).isdigit() and pd.notna(row[col]) and row[col] != "":
                            valor = str(row[col])
                            if ':' in valor:
                                print(f"Ejemplo de formato en Columna {col}:")
                                print("-" * 30)
                                print(valor)
                                print("-" * 30)
                                ejemplo_encontrado = True
                                break
                    if ejemplo_encontrado:
                        break
            
            # Información sobre los filtros
            print("\n🔍 USO DE FILTROS EN EXCEL:")
            print("-" * 40)
            print("1. Abre el archivo en Excel")
            print("2. Haz clic en cualquier celda de la tabla")
            print("3. Ve a la pestaña 'Datos'")
            print("4. Haz clic en 'Filtro' si no está activado")
            print("5. Verás flechas ▼ en cada columna")
            print("6. Haz clic en la flecha para filtrar por esa columna")
            print("-" * 40)
            
            # Preguntar si abrir carpeta
            print("\n¿Quieres abrir la carpeta de resultados? (s/n): ", end="")
            if input().strip().lower() == 's':
                try:
                    if os.name == 'nt':
                        os.startfile('Resultados')
                    elif os.name == 'posix':
                        os.system(f'open "{os.path.abspath("Resultados")}"' if os.name == 'darwin' else f'xdg-open "{os.path.abspath("Resultados")}"')
                except:
                    print("No se pudo abrir la carpeta automáticamente")
        
        else:
            print("❌ No se pudieron procesar los datos")
            
    except Exception as e:
        print(f"❌ Error: {str(e)}")

def main():
    """Función principal del sistema depurado"""
    print("\n" + "="*60)
    print("      SISTEMA DE UNIFICACIÓN DE MARCAJE POR PERIODO")
    print("="*60)
    print("\n🎯 CARACTERÍSTICAS PRINCIPALES:")
    print("• Unificación por periodo de fechas (cruza meses)")
    print("• Ordenamiento automático por Columna A")
    print("• Formato de horas en líneas verticales")
    print("• Filtros en TODAS las columnas del Excel")
    print("• Formato profesional automático")
    print("• Soporte para archivos .xls y .xlsx")
    print("="*60)
    print("\n📌 IMPORTANTE: En el explorador de archivos, selecciona")
    print("   'Todos los archivos' para ver tus archivos .xls")
    print("="*60)
    
    if not verificar_dependencias():
        print("❌ Se requieren dependencias faltantes.")
        input("Presione Enter para salir...")
        return
    
    while True:
        mostrar_banner()
        mostrar_menu()
        
        opcion = input("\n👉 Seleccione una opción (1-2): ").strip()
        
        if opcion == "1":
            unificar_periodo()
        elif opcion == "2":
            print("\n" + "="*60)
            print("👋 ¡Gracias por usar el sistema! Hasta luego.")
            print("="*60)
            break
        else:
            print("\n⚠️ Opción no válida. Por favor, seleccione 1 o 2.")
        
        if opcion != "2":
            input("\nPresione Enter para continuar...")

if __name__ == "__main__":
    main()