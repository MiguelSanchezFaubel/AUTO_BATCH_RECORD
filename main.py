# ------------------------------------------------------------------------
# Autor: Miguel Sánchez Faubel
# Descripción: Programa para generar automaticamente los checklist descargando los arhivos 3 documentos de la plataorma de dismuntel.
# ------------------------------------------------------------------------
import os
from win32com import client
from PyPDF2 import PdfMerger
import errno 
import openpyxl
from openpyxl import load_workbook
from shutil import rmtree

#*************** FUNCIONES****************************
def excel2pdf(file_location):

    excel = client.Dispatch("Excel.Application")
    excel.Interactive=False
    excel.Visible=False

    workbook= excel.Workbooks.open(file_location)
    output= os.path.splitext(file_location)[0]

    workbook.ActiveSheet.ExportAsFixedFormat(0,output)
    workbook.Close()

#***************FIN FUNCIONES****************************


# print('Inserte el numero INICIAL de los cecklist que desea unir:(ejemplo:2230001)')
# NS_INIT=int(input())

# print('Inserte el numero FINAL de los cecklist que desea unir:(ejemplo:2230030)')
# NS_FIN=int(input())

# fusionador_final = PdfMerger()
# merged_reports = './merged_reports/merged_reports.pdf'
directorioactual=os.getcwd()

listsequiposdir = os.listdir(directorioactual)
print(listsequiposdir)

for equipo in(listsequiposdir):
    print(equipo[7:])

# try:
#     os.mkdir('./merged_reports/')
    
# except OSError as e:
#     if e.errno != errno.EEXIST:
#         raise

try:
    os.mkdir('./temporaly/')
except OSError as e:
    if e.errno != errno.EEXIST:
        raise

direccion_actual = os.getcwd()
for direquipo in(listsequiposdir):
    if(direquipo[0:3] == 'EPB'):
        os.rename(direccion_actual +'/' + direquipo, direccion_actual +'/REPORT_' + direquipo)
        print(direquipo)

listsequiposdir = os.listdir(directorioactual)
for direquipo in(listsequiposdir):
    print(listsequiposdir)
    if(direquipo[7:10] == 'EPB'):
        fusionador = PdfMerger()
        nombre_archivo_salida = f'{direquipo[7:]}.pdf'
        numero_equipo = str(direquipo[7:])
        nombre_carpeta = f'./Report_' + numero_equipo + '/'

        dir = os.listdir(nombre_carpeta)
        # print(dir)
        
        for file in dir:
            name, ext = os.path.splitext(f'{nombre_carpeta}/{file}')
            print (name, ext)

            if ext == '.pdf':

                fusionador.append(open(f'{nombre_carpeta}/{file}', 'rb'))
                
            if ext == '.xlsx':

                # print(f'{nombre_carpeta}{file}')
                try:
                    wb = load_workbook(f'{nombre_carpeta}{file}')

                except:
                    print(f'Ocurrio un error inesperado en {nombre_carpeta}{file}')

                ws1 = wb['BOARD']
                ws2 = wb['Test']
                ws0 = wb['Portada']

                wb.remove(ws0)
                ws1.delete_rows(40,81)

                ws1.print_area = 'A1:F13'
                ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE # Establecer la orientación de la página a horizontal  
                ws1.page_setup.paperSize = ws1.PAPERSIZE_A3  # Establecer el tamaño del papel en A4 (puedes cambiarlo a A5 o el que necesites)  
                ws1.sheet_view.view = "pageLayout"
                ws1.sheet_properties.pageSetUpPr.fitToPage = True
                # ws1.page_setup.fitToHeight = False
                # ws1.page_setup.fitToPage = True   
                ws1.page_setup.fitToWidth = 1  # Ajustar a 1 página de ancho  
                ws1.page_setup.fitToHeight = 1  # Sin límite en las páginas de alto      

                ws2.print_area = 'A1:F25'
                ws2.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE # Establecer la orientación de la página a horizontal  
                ws2.page_setup.paperSize = ws1.PAPERSIZE_A3  # Establecer el tamaño del papel en A4 (puedes cambiarlo a A5 o el que necesites)
                ws2.sheet_view.view = "pageLayout"
                ws2.sheet_properties.pageSetUpPr.fitToPage = True 
                ws2.page_setup.fitToWidth = 1  # Ajustar a 1 página de ancho  
                ws2.page_setup.fitToHeight = 1  # Sin límite en las páginas de alto   
                
                wb.save(f'./temporaly/{numero_equipo}.xlsx')
                wb.close()
                # print(f'{directorioactual}/temporaly/EPB{numero_equipo_raw}.xlsx')
                excel2pdf(f'{directorioactual}/temporaly/{numero_equipo}.xlsx')

                fusionador.merge(0,open(f'{directorioactual}/temporaly/{numero_equipo}.pdf', 'rb')) ## 0 es la posicion en la que se introduce el pdf

            
        with open(f'{nombre_carpeta}{numero_equipo}.pdf', 'wb') as salida:
            fusionador.write(salida)
            fusionador.close()
            # fusionador_final.append(open(nombre_archivo_salida, 'rb'))
            print(numero_equipo)

            # with open(merged_reports, 'wb') as output:
            #     fusionador_final.write(output)
            #     fusionador_final.close()

# rmtree(f'./temporaly') # Elimina el directorio auxiliar 