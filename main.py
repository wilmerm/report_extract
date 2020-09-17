"""
Módulo fuente para la extarción de informaición de los reportes de Sicflex.
"""
import sys
from reporte.reportes import SicflexReporteVentasPorFechaDetallada



def printt(*args, **kwargs):
    print("------------------------------------------------")
    print(*args, **kwargs)
    print("------------------------------------------------")


def exitt(msg, code):
    print(msg)
    exit(code)


reportes = [
    SicflexReporteVentasPorFechaDetallada,
]


def main(*args):

    printt("Tipo de reporte a generar.")
    
    for n in range(len(reportes)):
        print(n, ":", reportes[n].__name__)

    for n in range(3):
        opt = int(input("Seleccione el reporte: "))

        try:
            reporte_class = reportes[opt]
        except (KeyError) as e:
            print("El reporte no existe... {e}")
            if (n == 2):
                exitt("Debe elegir un reporte por su número.", 1)
            continue
        else:
            break

    printt("Abrir archivo excel que contiene el reporte.")
    for n in range(3):
        filename = str(input("Ruta del archivo: "))

        try:
            reporte = reporte_class(filename)
        except (BaseException) as e:
            print(e)
            if (n == 2):
                exitt("Debe indicar el archivo excel que contiene los datos a extraer.", 1)
            continue
        else:
            break


    printt("Guardar archivo como.")
    for n in range(3):
        filename_out = str(input("Nombre de archivo: "))
        
        try:
            print(reporte.ExportToExcel(filename_out))
        except (BaseException) as e:
            print(e)
            if (n == 2):
                exitt("Debe indicar el nombre del archivo a guardar.", 1)
            continue
        else:
            break

    



if __name__ == "__main__":
    main(sys.argv)