import sys
import pandas as pd
import os
import re
from datetime import datetime, timedelta
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import warnings
warnings.filterwarnings('ignore')

# Para formato Excel
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

class Worker(QThread):
    """Hilo para procesamiento en segundo plano"""
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal(pd.DataFrame)
    error = pyqtSignal(str)
    
    def __init__(self, archivos, fecha_inicio, fecha_fin):
        super().__init__()
        self.archivos = archivos
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        self._is_running = True
        
    def run(self):
        try:
            df = self.procesar_periodo_cruzado()
            if df is not None and not df.empty:
                self.finished.emit(df)
            else:
                self.error.emit("No se pudieron procesar los datos")
        except Exception as e:
            self.error.emit(f"Error: {str(e)}")
    
    def stop(self):
        self._is_running = False
        
    def formatear_horas_vertical(self, texto):
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
    
    def limpiar_datos_marcaje(self, df):
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
                df_limpio[col] = df_limpio[col].apply(self.formatear_horas_vertical)
        
        return df_limpio
    
    def leer_archivo_completo(self, ruta_archivo):
        """Lee un archivo completo de marcaje y detecta fecha del contenido"""
        nombre = os.path.basename(ruta_archivo)
        extension = os.path.splitext(nombre)[1].lower()
        self.message.emit(f"Leyendo: {nombre}")
        
        try:
            # Determinar el engine basado en la extensión
            if extension == '.xls':
                engine = 'xlrd'
            elif extension == '.xlsx':
                engine = 'openpyxl'
            else:
                self.message.emit(f"Extensión no soportada: {extension}")
                return None
            
            # PRIMERO: Intentar leer la fila 4 para obtener la fecha
            mes_detectado = None
            año_detectado = None
            
            try:
                # Leer solo la fila 4 (índice 3 si empezamos desde 0)
                df_fila_4 = pd.read_excel(
                    ruta_archivo, 
                    skiprows=3,  # Saltar 3 filas para llegar a la fila 4
                    nrows=1,     # Leer solo 1 fila
                    header=None, # Sin encabezado
                    engine=engine
                )
                
                if not df_fila_4.empty:
                    # Buscar "Made Date:" en cualquier celda de esta fila
                    for col in range(min(10, len(df_fila_4.columns))):  # Buscar en primeras 10 columnas
                        celda = str(df_fila_4.iloc[0, col])
                        if 'made date:' in celda.lower():
                            # Extraer fecha del formato "Made Date:2025/11/01-2025/11/30"
                            
                            # Buscar patrones de fecha
                            patrones_fecha = [
                                r'(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})',  # YYYY/MM/DD
                                r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})',  # DD/MM/YYYY
                            ]
                            
                            for patron in patrones_fecha:
                                matches = re.findall(patron, celda)
                                if matches:
                                    for match in matches:
                                        if len(match) == 3:
                                            try:
                                                # Intentar diferentes interpretaciones
                                                # Formato YYYY/MM/DD
                                                if len(match[0]) == 4:  # Primer grupo es año
                                                    año = int(match[0])
                                                    mes = int(match[1])
                                                    dia = int(match[2])
                                                # Formato DD/MM/YYYY  
                                                elif len(match[2]) == 4:  # Tercer grupo es año
                                                    año = int(match[2])
                                                    mes = int(match[1])
                                                    dia = int(match[0])
                                                else:
                                                    continue
                                                
                                                # Validar mes
                                                if 1 <= mes <= 12:
                                                    mes_detectado = mes
                                                    año_detectado = año
                                                    self.message.emit(f"  📅 Fecha detectada en contenido: {año}/{mes:02d}/{dia:02d}")
                                                    break
                                            except:
                                                continue
                                    
                                    if mes_detectado:
                                        break
                            
                            break
            except Exception as e:
                self.message.emit(f"  ⚠️ No se pudo leer fecha de cabecera: {str(e)}")
            
            # Ahora buscar la fila donde empiezan los datos reales
            fila_inicio_datos = 0
            df_datos = None
            
            # Buscar desde la fila 5 en adelante (después de "Made Date:")
            for skip_rows in range(4, 15):  # Buscar desde fila 5 hasta fila 19
                try:
                    df_prueba = pd.read_excel(
                        ruta_archivo, 
                        skiprows=skip_rows, 
                        header=None, 
                        nrows=5,
                        engine=engine
                    )
                    
                    # Buscar 'Employee ID' o similar en la primera fila
                    primera_fila = df_prueba.iloc[0] if len(df_prueba) > 0 else pd.Series()
                    primera_fila_str = " ".join(str(x).lower() for x in primera_fila if pd.notna(x))
                    
                    if any(keyword in primera_fila_str for keyword in ['employee', 'empleado', 'id']):
                        fila_inicio_datos = skip_rows
                        df_datos = pd.read_excel(
                            ruta_archivo, 
                            skiprows=skip_rows,
                            engine=engine
                        )
                        
                        # Guardar el mes detectado como atributo del DataFrame
                        if mes_detectado:
                            df_datos.attrs['mes_detectado'] = mes_detectado
                            df_datos.attrs['año_detectado'] = año_detectado
                        
                        self.message.emit(f"    ✓ {len(df_datos)} registros encontrados")
                        df_datos = self.limpiar_datos_marcaje(df_datos)
                        return df_datos
                except Exception as e:
                    continue
            
            # Si no encontró con skip_rows, leer directamente
            if df_datos is None:
                df_datos = pd.read_excel(ruta_archivo, engine=engine)
                if mes_detectado:
                    df_datos.attrs['mes_detectado'] = mes_detectado
                    df_datos.attrs['año_detectado'] = año_detectado
            
            self.message.emit(f"    ✓ {len(df_datos)} registros (lectura directa)")
            df_datos = self.limpiar_datos_marcaje(df_datos)
            return df_datos
            
        except Exception as e:
            self.message.emit(f"    ✗ Error leyendo archivo: {str(e)}")
            return None
    
    def ordenar_datos_por_columna_a(self, df):
        """Ordena el DataFrame por la columna A (primera columna)"""
        if df is None or df.empty:
            return df
        
        df_ordenado = df.copy()
        
        # Identificar la columna para ordenar (columna A)
        columna_a = df_ordenado.columns[0] if len(df_ordenado.columns) > 0 else None
        
        if columna_a:
            try:
                # Intentar ordenar numéricamente si es posible
                if df_ordenado[columna_a].dtype in ['int64', 'float64']:
                    df_ordenado = df_ordenado.sort_values(by=columna_a, na_position='last')
                else:
                    # Convertir a string y ordenar
                    df_ordenado[columna_a] = df_ordenado[columna_a].astype(str)
                    
                    # Intentar extraer números para ordenar
                    def extract_number(x):
                        try:
                            numbers = re.findall(r'\d+', x)
                            if numbers:
                                return int(numbers[0])
                            return float('inf')
                        except:
                            return float('inf')
                    
                    # Agregar columna temporal para ordenar
                    df_ordenado['_temp_sort'] = df_ordenado[columna_a].apply(extract_number)
                    df_ordenado = df_ordenado.sort_values(by='_temp_sort', na_position='last')
                    df_ordenado = df_ordenado.drop('_temp_sort', axis=1)
                    
                    # También ordenar alfabéticamente
                    df_ordenado = df_ordenado.sort_values(by=columna_a, key=lambda col: col.str.lower())
            
            except Exception as e:
                self.message.emit(f"  ⚠️ No se pudo ordenar por '{columna_a}': {str(e)}")
        
        return df_ordenado
    
    def procesar_periodo_cruzado(self):
        """Procesa archivos de un periodo que cruza meses - VERSIÓN CORREGIDA"""
        # Filtrar solo archivos compatibles (.xls, .xlsx)
        archivos_compatibles = []
        
        for archivo in self.archivos:
            if not os.path.exists(archivo):
                continue
            
            extension = os.path.splitext(archivo)[1].lower()
            if extension in ['.xls', '.xlsx']:
                archivos_compatibles.append(archivo)
        
        if not archivos_compatibles:
            return None
        
        self.message.emit(f"📁 Procesando {len(archivos_compatibles)} archivo(s)")
        
        # Diccionario para almacenar dataframes con información de MES
        # Estructura: {mes: [(dia, dataframe), ...]}
        datos_por_mes_dia = {}
        
        for i, archivo in enumerate(archivos_compatibles):
            if not self._is_running:
                return None
                
            nombre_archivo = os.path.basename(archivo)
            self.message.emit(f"📄 Procesando: {nombre_archivo}")
            
            # Leer archivo
            df = self.leer_archivo_completo(archivo)
            
            if df is not None and not df.empty:
                # IMPORTANTE: Intentar detectar el MES del contenido
                mes_archivo = self.detectar_mes_del_archivo(archivo, df)
                
                if mes_archivo is None:
                    self.message.emit(f"  ⚠️ No se pudo detectar el mes, se intentará inferir")
                    # Intentar inferir del rango de fechas
                    mes_archivo = self.inferir_mes_del_rango(df, self.fecha_inicio, self.fecha_fin)
                
                if mes_archivo:
                    self.message.emit(f"  📅 Mes detectado: {mes_archivo}")
                    
                    # Obtener TODOS los días en el rango completo
                    todos_los_dias_en_rango = []
                    current_date = self.fecha_inicio
                    while current_date <= self.fecha_fin:
                        todos_los_dias_en_rango.append((current_date.month, current_date.day))
                        current_date += timedelta(days=1)
                    
                    # Seleccionar columnas
                    columnas_fijas = []
                    for col in df.columns:
                        col_str = str(col).lower()
                        if any(keyword in col_str for keyword in ['employee', 'empleado', 'id', 'name', 'nombre', 'department', 'departamento']):
                            columnas_fijas.append(col)
                    
                    if len(columnas_fijas) < 2:
                        columnas_fijas = df.columns[:3].tolist()
                    
                    # Para cada columna de día, verificar si está en el rango PARA ESTE MES
                    columnas_a_mantener = columnas_fijas.copy()
                    
                    for col in df.columns:
                        try:
                            dia_numero = int(float(str(col)))
                            if 1 <= dia_numero <= 31:
                                # Verificar si este día de este mes está en el rango
                                if (mes_archivo, dia_numero) in todos_los_dias_en_rango:
                                    columnas_a_mantener.append(col)
                        except:
                            pass
                    
                    columnas_existentes = [col for col in columnas_a_mantener if col in df.columns]
                    
                    if columnas_existentes:
                        df_filtrado = df[columnas_existentes].copy()
                        
                        # Renombrar columnas de días para INCLUIR EL MES
                        # Ejemplo: "5" → "nov-5" o "dic-5"
                        for col in df_filtrado.columns:
                            try:
                                dia_numero = int(float(str(col)))
                                if 1 <= dia_numero <= 31:
                                    # Crear nombre con mes y día
                                    nombre_mes = self.obtener_nombre_mes(mes_archivo)
                                    nuevo_nombre = f"{nombre_mes}-{dia_numero}"
                                    df_filtrado = df_filtrado.rename(columns={col: nuevo_nombre})
                            except:
                                pass
                        
                        # Guardar con información del mes
                        if mes_archivo not in datos_por_mes_dia:
                            datos_por_mes_dia[mes_archivo] = []
                        datos_por_mes_dia[mes_archivo].append(df_filtrado)
                        self.message.emit(f"  ✅ {len(df_filtrado)} registros para mes {mes_archivo}")
                    
            # Actualizar progreso
            progreso = int((i + 1) / len(archivos_compatibles) * 50)
            self.progress.emit(progreso)
        
        if not datos_por_mes_dia:
            self.message.emit("❌ No se encontraron datos en los archivos")
            return None
        
        self.message.emit(f"🔀 Combinando datos de {len(datos_por_mes_dia)} mes(es)...")
        
        # Combinar todos los dataframes
        df_combinado = self.combinar_dataframes_con_mes(datos_por_mes_dia, self.fecha_inicio, self.fecha_fin)
        
        if df_combinado is None or df_combinado.empty:
            return None
        
        # ORDENAR POR COLUMNA A
        self.message.emit("Ordenando datos por columna A...")
        df_combinado = self.ordenar_datos_por_columna_a(df_combinado)
        
        self.progress.emit(100)
        return df_combinado
    
    def detectar_mes_del_archivo(self, ruta_archivo, df):
        """Intenta detectar el mes al que pertenece el archivo"""
        # 1. Primero verificar si ya se detectó el mes en leer_archivo_completo
        if hasattr(df, 'attrs') and 'mes_detectado' in df.attrs:
            return df.attrs['mes_detectado']
        
        nombre = os.path.basename(ruta_archivo).lower()
        
        # 2. Si no, intentar por nombre del archivo
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
                return valor
        
        # 3. Intentar detectar del contenido (buscar fechas en cabeceras)
        try:
            # Leer las primeras filas para buscar información de fecha
            for i in range(min(5, len(df))):
                for j in range(min(5, len(df.columns))):
                    celda = str(df.iloc[i, j]).lower()
                    for mes_nombre, mes_num in meses_dict.items():
                        if mes_nombre in celda and len(mes_nombre) > 1:  # Evitar coincidencias cortas
                            return mes_num
        except:
            pass
        
        return None
    
    def inferir_mes_del_rango(self, df, fecha_inicio, fecha_fin):
        """Intenta inferir el mes basado en el rango de fechas y los días con datos"""
        # Obtener días del dataframe que tienen datos
        dias_con_datos = []
        for col in df.columns:
            try:
                dia = int(float(str(col)))
                if 1 <= dia <= 31:
                    # Verificar si esta columna tiene algún dato no vacío
                    columna_no_vacia = False
                    for valor in df[col]:
                        if pd.notna(valor) and str(valor).strip() and str(valor).strip() not in ['', 'nan']:
                            columna_no_vacia = True
                            break
                    
                    if columna_no_vacia:
                        dias_con_datos.append(dia)
            except:
                pass
        
        if not dias_con_datos:
            return None
        
        # Verificar qué mes del rango tiene más coincidencias
        meses_posibles = {}
        
        current_date = fecha_inicio
        while current_date <= fecha_fin:
            mes = current_date.month
            dia = current_date.day
            
            if dia in dias_con_datos:
                if mes not in meses_posibles:
                    meses_posibles[mes] = 0
                meses_posibles[mes] += 1
            
            current_date += timedelta(days=1)
        
        if meses_posibles:
            # Devolver el mes con más coincidencias
            return max(meses_posibles.items(), key=lambda x: x[1])[0]
        
        return None
    
    def obtener_nombre_mes(self, numero_mes):
        """Convierte número de mes a nombre abreviado"""
        meses = {
            1: 'ene', 2: 'feb', 3: 'mar', 4: 'abr', 5: 'may', 6: 'jun',
            7: 'jul', 8: 'ago', 9: 'sep', 10: 'oct', 11: 'nov', 12: 'dic'
        }
        return meses.get(numero_mes, f"mes{numero_mes}")
    
    def combinar_dataframes_con_mes(self, datos_por_mes, fecha_inicio, fecha_fin):
        """Combina dataframes que ya tienen información de mes en los nombres de columnas"""
        if not datos_por_mes:
            return pd.DataFrame()
        
        # Primero, combinar todos los dataframes de cada mes
        dataframes_combinados_por_mes = {}
        
        for mes, lista_dfs in datos_por_mes.items():
            if lista_dfs:
                # Combinar dataframes del mismo mes
                df_mes_combinado = lista_dfs[0].copy()
                
                for i in range(1, len(lista_dfs)):
                    df_actual = lista_dfs[i]
                    
                    # Buscar columna ID común
                    col_id = None
                    for col in df_mes_combinado.columns:
                        if 'id' in str(col).lower():
                            col_id = col
                            break
                    
                    if col_id and col_id in df_actual.columns:
                        for idx, row in df_actual.iterrows():
                            emp_id = row[col_id] if pd.notna(row[col_id]) else None
                            
                            if emp_id is not None:
                                id_str = str(emp_id).strip()
                                mask = df_mes_combinado[col_id].astype(str).str.strip() == id_str
                                
                                if mask.any():
                                    idx_existente = df_mes_combinado[mask].index[0]
                                    
                                    for col in df_actual.columns:
                                        if '-' in str(col):  # Columna con mes-día
                                            valor = row[col]
                                            if pd.notna(valor) and str(valor).strip() and str(valor).strip() not in ['', 'nan']:
                                                df_mes_combinado.at[idx_existente, col] = valor
                                else:
                                    df_mes_combinado = pd.concat([df_mes_combinado, pd.DataFrame([row])], ignore_index=True)
                    else:
                        df_mes_combinado = pd.concat([df_mes_combinado, df_actual], ignore_index=True)
                
                dataframes_combinados_por_mes[mes] = df_mes_combinado
        
        # Ahora combinar entre meses
        meses = sorted(dataframes_combinados_por_mes.keys())
        if not meses:
            return pd.DataFrame()
        
        df_final = dataframes_combinados_por_mes[meses[0]].copy()
        
        for mes in meses[1:]:
            df_mes = dataframes_combinados_por_mes[mes]
            
            col_id = None
            for col in df_final.columns:
                if 'id' in str(col).lower():
                    col_id = col
                    break
            
            if col_id and col_id in df_mes.columns:
                for idx, row in df_mes.iterrows():
                    emp_id = row[col_id] if pd.notna(row[col_id]) else None
                    
                    if emp_id is not None:
                        id_str = str(emp_id).strip()
                        mask = df_final[col_id].astype(str).str.strip() == id_str
                        
                        if mask.any():
                            idx_existente = df_final[mask].index[0]
                            
                            for col in df_mes.columns:
                                if '-' in str(col):  # Columna con mes-día
                                    valor = row[col]
                                    if pd.notna(valor) and str(valor).strip() and str(valor).strip() not in ['', 'nan']:
                                        df_final.at[idx_existente, col] = valor
                        else:
                            df_final = pd.concat([df_final, pd.DataFrame([row])], ignore_index=True)
            else:
                df_final = pd.concat([df_final, df_mes], ignore_index=True)
        
        # Ordenar columnas en orden cronológico
        columnas_fijas = [col for col in df_final.columns if '-' not in str(col)]
        columnas_dias = [col for col in df_final.columns if '-' in str(col)]
        
        # Ordenar columnas de días cronológicamente
        columnas_dias_ordenadas = self.ordenar_columnas_cronologicamente(columnas_dias, fecha_inicio, fecha_fin)
        
        # Columnas finales
        columnas_finales = columnas_fijas + columnas_dias_ordenadas
        columnas_existentes = [col for col in columnas_finales if col in df_final.columns]
        
        return df_final[columnas_existentes]
    
    def ordenar_columnas_cronologicamente(self, columnas_dias, fecha_inicio, fecha_fin):
        """Ordena columnas con formato 'mes-día' cronológicamente"""
        # Crear lista de tuplas (fecha, nombre_columna)
        columnas_con_fecha = []
        
        for col in columnas_dias:
            try:
                # Extraer mes y día del formato "nov-5" o "dic-28"
                partes = str(col).split('-')
                if len(partes) == 2:
                    mes_str, dia_str = partes
                    
                    # Convertir nombre de mes a número
                    meses_dict = {
                        'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
                        'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12
                    }
                    
                    mes = meses_dict.get(mes_str.lower())
                    if mes:
                        dia = int(dia_str)
                        
                        # Asumir año (usar año de fecha_inicio o fecha_fin según el mes)
                        if fecha_inicio.month == mes:
                            año = fecha_inicio.year
                        elif fecha_fin.month == mes:
                            año = fecha_fin.year
                        else:
                            # Si no coincide, usar el año más cercano
                            año = fecha_inicio.year if abs(mes - fecha_inicio.month) < abs(mes - fecha_fin.month) else fecha_fin.year
                        
                        try:
                            fecha = datetime(año, mes, dia)
                            columnas_con_fecha.append((fecha, col))
                        except:
                            pass
            except:
                pass
        
        # Ordenar por fecha
        columnas_con_fecha.sort(key=lambda x: x[0])
        
        # Devolver solo los nombres de columnas ordenados
        return [col for fecha, col in columnas_con_fecha]

class SplashScreen(QSplashScreen):
    """Pantalla de inicio"""
    def __init__(self):
        pixmap = QPixmap(400, 300)
        pixmap.fill(Qt.white)
        super().__init__(pixmap)
        
        self.setWindowTitle("Sistema de Marcaje")
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Logo/título
        title = QLabel("SISTEMA DE UNIFICACIÓN DE MARCAJE")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #2c3e50; margin: 20px;")
        layout.addWidget(title)
        
        # Versión
        version = QLabel("Versión 1.0")
        version.setAlignment(Qt.AlignCenter)
        version.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(version)
        
        # Separador
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #bdc3c7; margin: 20px;")
        layout.addWidget(line)
        
        # Cargando...
        self.loading_label = QLabel("Cargando...")
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_label.setStyleSheet("color: #3498db;")
        layout.addWidget(self.loading_label)
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

class DateDialog(QDialog):
    """Diálogo para seleccionar fechas"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Seleccionar Periodo")
        self.setModal(True)
        self.setFixedSize(400, 250)
        
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("📅 ESPECIFICA EL PERIODO A PROCESAR")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #2c3e50; margin: 10px;")
        layout.addWidget(title)
        
        # Instrucciones
        inst = QLabel("Formato: DD/MM/AAAA\nEjemplo: 26/11/2025")
        inst.setAlignment(Qt.AlignCenter)
        inst.setStyleSheet("color: #7f8c8d; margin: 10px;")
        layout.addWidget(inst)
        
        # Fecha inicio
        layout.addWidget(QLabel("Fecha de inicio:"))
        self.inicio_edit = QLineEdit()
        self.inicio_edit.setPlaceholderText("DD/MM/AAAA")
        layout.addWidget(self.inicio_edit)
        
        # Fecha fin
        layout.addWidget(QLabel("Fecha de fin:"))
        self.fin_edit = QLineEdit()
        self.fin_edit.setPlaceholderText("DD/MM/AAAA")
        layout.addWidget(self.fin_edit)
        
        # Botones
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def get_fechas(self):
        try:
            inicio = datetime.strptime(self.inicio_edit.text(), "%d/%m/%Y")
            fin = datetime.strptime(self.fin_edit.text(), "%d/%m/%Y")
            return inicio, fin
        except:
            return None, None

class MainWindow(QMainWindow):
    """Ventana principal de la aplicación"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Unificación de Marcaje")
        self.setGeometry(100, 100, 900, 700)
        
        # Configurar estilo
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
            QLabel {
                color: #2c3e50;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #bdc3c7;
                border-radius: 3px;
                padding: 5px;
                background-color: white;
            }
            QProgressBar {
                border: 1px solid #bdc3c7;
                border-radius: 3px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 3px;
            }
        """)
        
        self.init_ui()
        self.worker = None
        self.df_final = None
        
    def init_ui(self):
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Título
        title = QLabel("🏢 SISTEMA DE UNIFICACIÓN DE MARCAJE")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2c3e50; margin: 20px;")
        layout.addWidget(title)
        
        # Subtítulo
        subtitle = QLabel("Unificación por periodo de fechas")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("font-size: 14px; color: #7f8c8d; margin-bottom: 30px;")
        layout.addWidget(subtitle)
        
        # Panel de periodo
        periodo_group = QGroupBox("📅 PERIODO A PROCESAR")
        periodo_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        periodo_layout = QVBoxLayout()
        
        # Fechas
        fecha_layout = QHBoxLayout()
        fecha_layout.addWidget(QLabel("Fecha inicio:"))
        self.fecha_inicio_label = QLabel("No establecida")
        self.fecha_inicio_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        fecha_layout.addWidget(self.fecha_inicio_label)
        fecha_layout.addStretch()
        fecha_layout.addWidget(QLabel("Fecha fin:"))
        self.fecha_fin_label = QLabel("No establecida")
        self.fecha_fin_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
        fecha_layout.addWidget(self.fecha_fin_label)
        
        periodo_layout.addLayout(fecha_layout)
        
        # Botón para seleccionar fechas
        self.btn_seleccionar_fechas = QPushButton("📅 Seleccionar Periodo")
        self.btn_seleccionar_fechas.clicked.connect(self.seleccionar_fechas)
        periodo_layout.addWidget(self.btn_seleccionar_fechas)
        
        periodo_group.setLayout(periodo_layout)
        layout.addWidget(periodo_group)
        
        # Panel de archivos
        archivos_group = QGroupBox("📁 ARCHIVOS SELECCIONADOS")
        archivos_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #2ecc71;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        
        archivos_layout = QVBoxLayout()
        
        # Lista de archivos
        self.lista_archivos = QListWidget()
        self.lista_archivos.setMinimumHeight(150)
        archivos_layout.addWidget(self.lista_archivos)
        
        # Botones para archivos
        botones_archivos_layout = QHBoxLayout()
        
        self.btn_agregar_archivos = QPushButton("➕ Agregar Archivos")
        self.btn_agregar_archivos.clicked.connect(self.agregar_archivos)
        botones_archivos_layout.addWidget(self.btn_agregar_archivos)
        
        self.btn_limpiar_archivos = QPushButton("🗑️ Limpiar Lista")
        self.btn_limpiar_archivos.clicked.connect(self.limpiar_archivos)
        botones_archivos_layout.addWidget(self.btn_limpiar_archivos)
        
        archivos_layout.addLayout(botones_archivos_layout)
        
        archivos_group.setLayout(archivos_layout)
        layout.addWidget(archivos_group)
        
        # Panel de progreso
        progreso_group = QGroupBox("⚙️ PROCESAMIENTO")
        progreso_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #f39c12;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
        """)
        
        progreso_layout = QVBoxLayout()
        
        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        progreso_layout.addWidget(self.progress_bar)
        
        # Etiqueta de estado
        self.status_label = QLabel("Listo para procesar")
        self.status_label.setAlignment(Qt.AlignCenter)
        progreso_layout.addWidget(self.status_label)
        
        # Log de mensajes
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)
        progreso_layout.addWidget(QLabel("Mensajes:"))
        progreso_layout.addWidget(self.log_text)
        
        progreso_group.setLayout(progreso_layout)
        layout.addWidget(progreso_group)
        
        # Botón de procesamiento
        self.btn_procesar = QPushButton("🚀 INICIAR PROCESAMIENTO")
        self.btn_procesar.clicked.connect(self.iniciar_procesamiento)
        self.btn_procesar.setEnabled(False)
        self.btn_procesar.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                font-size: 16px;
                padding: 15px;
                margin: 20px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        layout.addWidget(self.btn_procesar)
        
        # Botón para abrir carpeta de resultados
        self.btn_abrir_resultados = QPushButton("📂 Abrir Carpeta de Resultados")
        self.btn_abrir_resultados.clicked.connect(self.abrir_resultados)
        self.btn_abrir_resultados.setEnabled(False)
        layout.addWidget(self.btn_abrir_resultados)
        
        # Estado
        self.estado_label = QLabel("")
        self.estado_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.estado_label)
        
        layout.addStretch()
        
    def seleccionar_fechas(self):
        dialog = DateDialog(self)
        if dialog.exec_():
            inicio, fin = dialog.get_fechas()
            if inicio and fin:
                if fin >= inicio:
                    self.fecha_inicio = inicio
                    self.fecha_fin = fin
                    self.fecha_inicio_label.setText(inicio.strftime("%d/%m/%Y"))
                    self.fecha_fin_label.setText(fin.strftime("%d/%m/%Y"))
                    self.fecha_inicio_label.setStyleSheet("color: #27ae60; font-weight: bold;")
                    self.fecha_fin_label.setStyleSheet("color: #27ae60; font-weight: bold;")
                    self.verificar_estado()
                    self.log("✓ Periodo establecido")
                else:
                    QMessageBox.warning(self, "Error", "La fecha de fin debe ser posterior a la fecha de inicio")
    
    def agregar_archivos(self):
        archivos, _ = QFileDialog.getOpenFileNames(
            self,
            "Seleccionar archivos (.xls, .xlsx)",
            "",
            "Todos los archivos (*.*);;Excel (*.xlsx *.xls);;Excel antiguo (*.xls);;Excel moderno (*.xlsx)"
        )
        
        for archivo in archivos:
            nombre = os.path.basename(archivo)
            if not self.archivo_en_lista(archivo):
                item = QListWidgetItem(f"📄 {nombre}")
                item.setData(Qt.UserRole, archivo)
                self.lista_archivos.addItem(item)
        
        if archivos:
            self.verificar_estado()
            self.log(f"✓ {len(archivos)} archivo(s) agregado(s)")
    
    def archivo_en_lista(self, archivo):
        for i in range(self.lista_archivos.count()):
            item = self.lista_archivos.item(i)
            if item.data(Qt.UserRole) == archivo:
                return True
        return False
    
    def limpiar_archivos(self):
        self.lista_archivos.clear()
        self.verificar_estado()
        self.log("✓ Lista de archivos limpiada")
    
    def verificar_estado(self):
        tiene_fechas = hasattr(self, 'fecha_inicio') and hasattr(self, 'fecha_fin')
        tiene_archivos = self.lista_archivos.count() > 0
        
        self.btn_procesar.setEnabled(tiene_fechas and tiene_archivos)
        
        if tiene_fechas and tiene_archivos:
            self.estado_label.setText("✅ Listo para procesar")
            self.estado_label.setStyleSheet("color: #27ae60; font-weight: bold;")
        else:
            self.estado_label.setText("⚠️ Complete todos los campos")
            self.estado_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
    
    def log(self, mensaje):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {mensaje}")
        QApplication.processEvents()
    
    def iniciar_procesamiento(self):
        # Obtener archivos
        archivos = []
        for i in range(self.lista_archivos.count()):
            item = self.lista_archivos.item(i)
            archivos.append(item.data(Qt.UserRole))
        
        # Deshabilitar controles
        self.btn_procesar.setEnabled(False)
        self.btn_agregar_archivos.setEnabled(False)
        self.btn_seleccionar_fechas.setEnabled(False)
        self.btn_limpiar_archivos.setEnabled(False)
        self.progress_bar.setValue(0)
        
        # Iniciar worker
        self.worker = Worker(archivos, self.fecha_inicio, self.fecha_fin)
        self.worker.progress.connect(self.actualizar_progreso)
        self.worker.message.connect(self.log)
        self.worker.finished.connect(self.procesamiento_completado)
        self.worker.error.connect(self.procesamiento_error)
        self.worker.start()
        
        self.log("🚀 Iniciando procesamiento...")
    
    def actualizar_progreso(self, valor):
        self.progress_bar.setValue(valor)
    
    def procesamiento_completado(self, df):
        self.df_final = df
        
        # Generar nombre de archivo
        inicio_str = self.fecha_inicio.strftime("%Y%m%d")
        fin_str = self.fecha_fin.strftime("%Y%m%d")
        nombre_salida = f"Marcaje_{inicio_str}_{fin_str}_Unificado.xlsx"
        
        # Crear carpeta de resultados si no existe
        if not os.path.exists('Resultados'):
            os.makedirs('Resultados')
        
        ruta_completa = os.path.join('Resultados', nombre_salida)
        
        try:
            # Guardar con formato y filtros
            with pd.ExcelWriter(ruta_completa, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Marcaje', index=False)
                self.aplicar_formato_excel(writer, df)
            
            self.log(f"✅ Archivo guardado: {ruta_completa}")
            
            # Habilitar botones
            self.btn_procesar.setEnabled(True)
            self.btn_agregar_archivos.setEnabled(True)
            self.btn_seleccionar_fechas.setEnabled(True)
            self.btn_limpiar_archivos.setEnabled(True)
            self.btn_abrir_resultados.setEnabled(True)
            
            # Mostrar resumen
            dias = len([c for c in df.columns if str(c).isdigit()])
            QMessageBox.information(
                self,
                "✅ Procesamiento Completado",
                f"✅ PROCESAMIENTO COMPLETADO\n\n"
                f"📊 Total de empleados: {len(df)}\n"
                f"📁 Archivos procesados: {self.lista_archivos.count()}\n"
                f"📅 Periodo: {self.fecha_inicio.strftime('%d/%m/%Y')} - {self.fecha_fin.strftime('%d/%m/%Y')}\n"
                f"🔢 Días mostrados: {dias}\n"
                f"💾 Guardado como: {nombre_salida}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar archivo: {str(e)}")
    
    def aplicar_formato_excel(self, writer, df):
        """Aplica formato y filtros al archivo Excel"""
        workbook = writer.book
        worksheet = writer.sheets['Marcaje']
        
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
        
        # Formatear celdas de datos
        for row in range(2, len(df) + 2):
            max_lineas = 1
            
            for col in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                
                if cell.value and isinstance(cell.value, str):
                    lineas = cell.value.count('\n') + 1
                    if lineas > max_lineas:
                        max_lineas = lineas
            
            worksheet.row_dimensions[row].height = 15 * max_lineas
        
        # Ajustar ancho de columnas
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value:
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
            
            adjusted_width = max_length + 2
            
            if column_letter in ['A', 'B', 'C']:
                adjusted_width = min(max(adjusted_width, 12), 25)
            elif column[0].column > 3:
                adjusted_width = min(max(adjusted_width, 6), 10)
            
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Congelar paneles
        worksheet.freeze_panes = 'D2'
    
    def procesamiento_error(self, mensaje):
        self.log(f"❌ Error: {mensaje}")
        
        # Habilitar controles
        self.btn_procesar.setEnabled(True)
        self.btn_agregar_archivos.setEnabled(True)
        self.btn_seleccionar_fechas.setEnabled(True)
        self.btn_limpiar_archivos.setEnabled(True)
        
        QMessageBox.critical(self, "Error", mensaje)
    
    def abrir_resultados(self):
        try:
            if os.name == 'nt':
                os.startfile('Resultados')
            elif os.name == 'posix':
                if os.name == 'darwin':
                    os.system(f'open "{os.path.abspath("Resultados")}"')
                else:
                    os.system(f'xdg-open "{os.path.abspath("Resultados")}"')
        except:
            QMessageBox.warning(self, "Error", "No se pudo abrir la carpeta de resultados")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Crear y mostrar splash screen
    splash = SplashScreen()
    splash.show()
    
    # Simular carga
    for i in range(101):
        splash.progress_bar.setValue(i)
        splash.loading_label.setText(f"Cargando... {i}%")
        QApplication.processEvents()
        QThread.msleep(20)
    
    # Crear ventana principal
    window = MainWindow()
    
    # Cerrar splash y mostrar ventana principal
    splash.finish(window)
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()