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

    reporte = SicflexReporteVentasPorFechaDetallada("files/ventas.xls")
    reporte.SetVendedores("files/facturas.csv")
    reporte.ExportToExcel("files/out.xls")
    exit()

    printt("Tipo de reporte a generar.")
    for n in range(len(reportes)):
        print(n, ":", reportes[n].__name__)

    opt = int(input("Seleccione el reporte: "))
    reporte_class = reportes[opt]

    printt("Abrir archivo excel que contiene el reporte.")
    filename = str(input("Ruta del archivo: "))
    reporte = reporte_class(filename)

    printt("Datos extras:")
    filename = str(input("Ruta del archivo: "))
    reporte.SetVendedores(filename)

    printt("Guardar archivo como.")
    filename_out = str(input("Nombre de archivo: "))
    print(reporte.ExportToExcel(filename_out))


    



if __name__ == "__main__":
    main(sys.argv)