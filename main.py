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
    print(msg, end="\n\n")
    exit(code)


reportes = [
    SicflexReporteVentasPorFechaDetallada,
]


def main(*args):

    args = args[0]

    if len(args) == 1:
        printt("Error. Faltan argumentos.")
        exitt("Escriba 'python main.py help' para ver la ayuda.", 111)

    if args[1].lower() == "help":
        printt("Ayuda:")
        exitt("Ej.: >> python main.py [ventas.xls] [extras.csv] [out.xls]", 0)
    
    if len(args) < 4:
        printt("Error. faltas argumentos.")
        exitt("Escriba 'python mani.py help' para ver la ayuda.", 112)

    if len(args) > 4:
        printt("Error. demasiados argumentos.")
        exitt("Escriba 'python mani.py help' para ver la ayuda.", 113)

    file_ventas = args[1]
    file_extras = args[2]
    file_out = args[3]

    print("Ventas: ", file_ventas)
    print("Extras: ", file_extras)
    print("Salida: ", file_out)

    printt("Cargado las ventas...")
    reporte = SicflexReporteVentasPorFechaDetallada(file_ventas)
    print("Ventas cargadas")

    printt("Cargado los vendedores...")
    reporte.SetVendedores(file_extras)
    print("Vendedores cargados")
    
    printt("Guardando el reporte...")
    reporte.ExportToExcel(file_out)
    print("Reporte guardado")
    exitt("¡Completado!", 0)




if __name__ == "__main__":
    main(sys.argv)