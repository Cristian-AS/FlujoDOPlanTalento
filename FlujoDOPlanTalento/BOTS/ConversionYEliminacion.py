# Explica el codigo y su propisito

"""
El código de este archivo se encarga de eliminar los archivos .xlsx que se encuentran en la carpeta PDFCarpeta.
La función Eliminar_Archivos itera sobre los archivos en la carpeta PDFCarpeta y elimina aquellos que tengan la extensión .xlsx.
La función Copia_y_Traslado se encarga de copiar los archivos .xlsx de la carpeta InsumoCambioCargo a la carpeta PDFCarpeta y luego guardarlos como archivos PDF en la misma carpeta.
La función save_as_pdf utiliza la librería win32com para abrir el archivo Excel, exportar la hoja especificada como PDF y cerrar el archivo.
"""

# Importar librerías necesarias
import os
from dotenv import load_dotenv
from openpyxl import load_workbook
import pandas as pd
import win32com.client
from pathlib import Path

# Cargar las variables de entorno
load_dotenv()

# Definir las rutas de los archivos y carpetas
InsumoDesarrolloHumano = os.getenv('InsumoDesarrolloHumano')
InsumoCambioCargo = os.getenv('InsumoCambioCargo')
PDFCarpeta = os.getenv('PDFCarpeta')

# Funcion para copiar los archivos xlsx de InsumoCambioCargo a PDFCarpeta y guardarlos como PDF
def Copia_y_Traslado():
    # Leer el archivo Excel principal
    df = pd.read_excel(InsumoDesarrolloHumano)

    filtered_df = df[df['Automatizacion'] == 'Excel']

    for index, row in filtered_df.iterrows():
        cedula = int(row['Cédula'])
        filename = f"{cedula}.xlsx"
        filepath = os.path.join(InsumoCambioCargo, filename)

        # Verificar que el archivo Excel original existe
        if not os.path.isfile(filepath):
            print(f"El archivo {filepath} no existe.")
            continue
        
        # Cargar el libro de trabajo Excel original
        libro = load_workbook(filepath)

        # Construir una ruta única para guardar la copia
        save_path = os.path.join(PDFCarpeta, filename)

        # Guardar la copia en la nueva ruta
        libro.save(save_path)
        
        # Construir el nombre del archivo PDF
        pdf_filename = f"{cedula}.pdf"
        pdf_path = os.path.join(PDFCarpeta, pdf_filename)
        save_as_pdf(save_path, "PLAN TALENTO", pdf_path)

# Funcion para guardar un archivo Excel como PDF
def save_as_pdf(excel_path, sheet_name, pdf_path):
    
    pdf_file = Path(pdf_path)
    # Verifica si el archivo PDF ya existe
    if pdf_file.exists():
        print(f"El archivo PDF ya existe: {pdf_path}")
        return
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(excel_path)
    try:
        sheet = workbook.Sheets(sheet_name)
        
        # Configurar la hoja para que se ajuste a una sola página en ancho y alto
        sheet.PageSetup.Zoom = 85
        sheet.PageSetup.FitToPagesWide = False
        sheet.PageSetup.FitToPagesTall = False  # Mantener False para evitar comprimir verticalmente demasiado
        
        # Asegúrate de que el área de impresión sea correcta
        sheet.PageSetup.PrintArea = sheet.UsedRange.Address

        # Usa orientación horizontal
        sheet.PageSetup.Orientation = 1  # 2 para orientación horizontal

        # Configurar el tamaño del papel A4
        sheet.PageSetup.PaperSize = 8  # A4

        # Ajusta los márgenes
        sheet.PageSetup.LeftMargin = excel.InchesToPoints(1.15) #1.95
        sheet.PageSetup.RightMargin = excel.InchesToPoints(0.05)
        sheet.PageSetup.TopMargin = excel.InchesToPoints(0.95)
        sheet.PageSetup.BottomMargin = excel.InchesToPoints(0.25)

        # Establecer calidad de impresión
        sheet.PageSetup.PrintQuality = 600
        
        # Ajuste final de escala
        sheet.PageSetup.CenterHorizontally = True
        sheet.PageSetup.CenterVertically = True
        
        sheet.ExportAsFixedFormat(0, pdf_path)
    except Exception as e:
        print(f"Error al guardar {excel_path} como PDF: {e}")
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()

# Funcion para entrar a la carpeta PDFCarpeta y eliminar los archivos xlsx
def Eliminar_Archivos():
    for archivo in os.listdir(PDFCarpeta):
        if archivo.endswith(".xlsx"):
            os.remove(os.path.join(PDFCarpeta, archivo))
            print(f"Archivo {archivo} eliminado")

# Ejecutar las funciones 
if __name__ == "__main__":
    Copia_y_Traslado()
    Eliminar_Archivos()