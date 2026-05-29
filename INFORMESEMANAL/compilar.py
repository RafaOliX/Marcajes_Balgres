import os
import subprocess
import sys

def main():
    print("=" * 60)
    print("    COMPILADOR DE SISTEMA DE MARCAJE")
    print("=" * 60)
    
    # Verificar que existe el archivo principal
    if not os.path.exists("informMarcajesUI.py"):
        print("❌ ERROR: No se encuentra 'informMarcajesUI.py'")
        print("   Asegúrate de que el archivo principal tenga ese nombre")
        input("Presiona Enter para salir...")
        return
    
    print("📋 Verificando dependencias...")
    
    # Instalar dependencias si no están
    dependencias = ['pyinstaller', 'pandas', 'openpyxl', 'xlrd', 'PyQt5']
    
    for dep in dependencias:
        try:
            if dep == 'pyinstaller':
                import PyInstaller
            elif dep == 'pandas':
                import pandas
            elif dep == 'openpyxl':
                import openpyxl
            elif dep == 'xlrd':
                import xlrd
            elif dep == 'PyQt5':
                import PyQt5
            print(f"   ✅ {dep}")
        except ImportError:
            print(f"   ⚠️  Instalando {dep}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                print(f"   ✅ {dep} instalado")
            except:
                print(f"   ❌ No se pudo instalar {dep}")
                return
    
    print("\n🔨 Compilando aplicación...")
    
    # Comando de compilación
    cmd = [
        'pyinstaller',
        '--onefile',           # Un solo archivo ejecutable
        '--windowed',          # Sin consola
        '--name=SistemaMarcaje',  # Nombre del ejecutable
        '--icon=icon.ico',     # Icono
        '--add-data=icon.ico;.',  # Incluir icono
        '--clean',             # Limpiar builds anteriores
        '--noconfirm',         # No preguntar para sobrescribir
        # Dependencias ocultas
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=xlrd',
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        'informMarcajesUI.py'
    ]
    
    try:
        print("   Esto puede tomar varios minutos...")
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ Compilación completada exitosamente!")
        
        # Crear carpeta de resultados
        dist_dir = os.path.join('dist', 'SistemaMarcaje')
        resultados_dir = os.path.join(dist_dir, 'Resultados')
        os.makedirs(resultados_dir, exist_ok=True)
        
        print("\n" + "=" * 60)
        print("📁 DIRECTORIOS CREADOS:")
        print(f"   • dist/SistemaMarcaje/ - Contiene el ejecutable")
        print(f"   • dist/SistemaMarcaje/Resultados/ - Para archivos generados")
        print("=" * 60)
        
        # Mostrar ubicación del ejecutable
        exe_path = os.path.join(dist_dir, 'SistemaMarcaje.exe')
        if os.path.exists(exe_path):
            print(f"\n🎉 ¡APLICACIÓN COMPILADA!")
            print(f"📍 Ubicación: {os.path.abspath(exe_path)}")
            print(f"📏 Tamaño: {os.path.getsize(exe_path) / (1024*1024):.2f} MB")
        else:
            # Buscar en build
            for root, dirs, files in os.walk('dist'):
                for file in files:
                    if file.endswith('.exe'):
                        exe_path = os.path.join(root, file)
                        print(f"\n🎉 ¡APLICACIÓN COMPILADA!")
                        print(f"📍 Ubicación: {os.path.abspath(exe_path)}")
                        print(f"📏 Tamaño: {os.path.getsize(exe_path) / (1024*1024):.2f} MB")
                        break
        
        # Crear README para el usuario
        readme_content = """# SISTEMA DE UNIFICACIÓN DE MARCAJE

## 🚀 Cómo usar la aplicación:

1. **Ejecutar**: Haz doble clic en `SistemaMarcaje.exe`
2. **Seleccionar periodo**: Haz clic en "Seleccionar Periodo"
3. **Agregar archivos**: Selecciona los archivos .xls o .xlsx
4. **Procesar**: Haz clic en "INICIAR PROCESAMIENTO"

## 📁 Características:

• Unifica múltiples archivos de marcaje
• Soporta archivos .xls y .xlsx
• Formato de horas en líneas verticales
• Ordena automáticamente por columna A
• Filtros en Excel listos para usar
• Resultados en carpeta "Resultados"

## ⚠️ Notas:

• La primera vez puede tardar unos segundos en iniciar
• Los archivos generados se guardan en la carpeta "Resultados"
• Necesita permisos para leer/crear archivos en la carpeta actual

Desarrollado por RafaelO - Sistema de Marcaje v1.0
"""
        
        with open(os.path.join(dist_dir, 'LEEME.txt'), 'w', encoding='utf-8') as f:
            f.write(readme_content)
        
        print("\n📄 Se creó archivo 'LEEME.txt' con instrucciones")
        
        # Preguntar si abrir carpeta
        print("\n¿Abrir carpeta con el ejecutable? (s/n): ", end='')
        respuesta = input().strip().lower()
        if respuesta == 's':
            os.startfile(os.path.abspath('dist'))
        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Error durante la compilación:")
        print(f"   Código: {e.returncode}")
        if e.stdout:
            print(f"   Salida: {e.stdout[:500]}...")
        if e.stderr:
            print(f"   Error: {e.stderr[:500]}...")
    
    except Exception as e:
        print(f"\n❌ Error inesperado: {str(e)}")
    
    finally:
        input("\nPresiona Enter para salir...")

if __name__ == "__main__":
    main()