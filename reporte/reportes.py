from reporte.base import Reporte, Excel, ReporteError






class SicflexReporteVentasPorFechaDetallada(Reporte):
    """
    Reporte Sicflex: Ventas/Ventas por Fecha/Factura (Detallada).

    Con la hoja ya limpia (clean()), la disposición de los datos sería así:
        6: Encabezados (
            1: Artículo.
            7: Descripción.
            12: Cantidad.
            16: Precio.
            20: Venta.
            24: Descuento.
            26: Itbis.
            28: Neto.
        )
        
        Información recurrente a partir de la fila 7:
        ------------------------------------------------------------------------
        7: Información de la factura (
            1: Fecha.
            3: Almacén.
            7: Número ('almacén-tipo-número').
            12: Cliente id.
            16: Cliente nombre.
            28: Moneda.
        )
        8: Movimiento... (
            1: Referencia.
            7: Descripción.
            12: Cantidad.
            16: Precio.
            20: Venta.
            24: Descuento.
            26: Itbis.
            28: Neto.
        )
        Al final de cada detalle, un subtotal (
            7: Cantidad de items diferentes.
            9: Un texto = 'Total  '
            20: Venta total.
            24: Descuento total.
            26: Itbis total.
            28: Neto total.
        )

    El inicio de cada detalle se determina evaluando en cada fila las celdas
    1 (fecha), 3 (almacén). Estas deben ser respectivamente una fecha y un texto,
    y ya que la celda 3 no la ocupa ningún otro dato, solo tenemos que determinar
    si está vacia o no.

    El final de cada detalle (fila de totales) lo determinamos verificando que 
    la celda 1 esté vacia y en la 7 haya un número.

    El final del reporte está determinado cuando, depúes del total del detalle, 
    no se encuentra el inicio del siguiente detalle, y en cambio existe un texto 
    'Total' en la celda 7.
    """

    ROW_ENCABEZADOS = 6
    COL_ENCABEZADOS = {
        1: "articulo",
        7: "descripcion",
        12: "cantidad",
        16: "precio",
        20: "venta",
        24: "descuento",
        26: "itbis",
        28: "neto",
    }

    ROW_FACTURA_DETALLE = 7
    COL_FACTURA_DETALLE = {
        1: "fecha",
        3: "almacen",
        7: "numero",
        12: "clienteid",
        16: "cliente_nombre",
        28: "moneda",
    }

    ENCABEZADOS = ("Fecha", "Almacén", "Número", "ClienteId", "Cliente nombre", 
    "Moneda", "Referencia", "Descripción", "Cantidad", "Precio", "Venta", 
    "Descuento", "Itbis", "Neto")

    ENCABEZADOS_KEYS = ("fecha", "almacen", "numero", "clienteid", "cliente_nombre", 
    "moneda", "articulo", "descripcion", "cantidad", "precio", "venta", 
    "descuento", "itbis", "neto")

    def __init__(self, filename):
        self.Open(filename)
        self.read()

    @classmethod
    def clean(self, data):
        """
        Limpia los datos del reporte Sicflex 'Ventas/Ventas por Fecha/Factura 
        (Detallada)', según la disposición de sus valores y retorna un listado 
        de dicionarios con los nuevos valores ordenados.
        """
        data = Excel.cleanrows(data)
        out = []
        datalength = len(data)
        index = self.ROW_FACTURA_DETALLE
        for index in range(self.ROW_FACTURA_DETALLE, datalength):

            row = data[index]

            # Detalle de la factura.
            if self.IsFacturaDetalle(row):
                item = {x: "" for x in self.COL_FACTURA_DETALLE.values()}
                for colindex, colname in self.COL_FACTURA_DETALLE.items():
                    item[colname] = row[colindex]
                # Movimientos.(por cada detalle puede haber más de un movimiento).
                item["movimientos"] = []
                out.append(item)

            # Movimiento (los movimientos irán dentro del último detalle agregado.)
            elif self.IsMovimiento(row):
                item = {x: "" for x in self.COL_ENCABEZADOS.values()}
                for colindex, colname in self.COL_ENCABEZADOS.items():
                    item[colname] = row[colindex]
                # Agregamos al último detalle.
                out[-1]["movimientos"].append(item)
            
            elif (self.IsMovimiento(row)):
                # En este punto no es posible encontrar un movimiento, de modo 
                # que si encontramos uno, es posible que sea un error.
                raise ReporteError("Al parecer hubo un error leyendo el reporte." \
                    "Se ha encontrado un movimiento huerfano (sin detalle de factura)." \
                    f"{row}, en la fila {index}")
            
            elif (str(row[7]).lower().strip().replace(" ", "") == "total"):
                # Fin del reporte.
                # El final del reporte está determinado cuando, depúes del total 
                # del detalle, no se encuentra el inicio del siguiente detalle, 
                # y en cambio existe un texto 'Total' en la celda 7.
                break
                
            else:
                # En este punto no hemos encontrado nada que coincida con lo que
                # andamos buscando, por lo tanto probamos una nueva iteración.
                continue
        
        return out

    @classmethod
    def IsFacturaDetalle(self, row):
        """
        Confirma si la fila pasada representa el detalle de la factura.

        El inicio de cada detalle se determina evaluando en cada fila las celdas
        1 (fecha), 3 (almacén). Estas deben ser respectivamente una fecha y un texto,
        y ya que la celda 3 no la ocupa ningún otro dato, solo tenemos que determinar
        si está vacia o no.
        """
        if (row[1]) and (row[3]):
            return True
        return False

    @classmethod
    def IsMovimiento(self, row):
        """
        Confirma que la fila pasada representa un movimiento.

        Para confirmar, verificamos la combinación única donde el index 1 y el 12
        no estén vacios, y que el 12 sea un número.
        """
        if (row[1]) and (row[12]):
            if isinstance(row[12], (int, float)):
                if isinstance(row[16], (int, float)):
                    return True
        return False
    
    def ExportToExcel(self, filename=None):
        """
        Exporta la data generada por este reporte, a formato Excel.
        """

        if (not filename):
            filename = f"{self.__class__.__name__}.xls"

        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("blog.unolet.com")
        rowindex = 0

        # Encabezados.
        colindex = 0
        for name in self.ENCABEZADOS:
            sheet.write(rowindex, colindex, name)
            colindex += 1

        rowindex += 1
        for detalle in self.data:

            # Movimientos.
            movimientos = detalle["movimientos"]
            for movimiento in movimientos:
                
                colindex = 0

                # Datos de la factura.
                for colname in self.COL_FACTURA_DETALLE.values():
                    sheet.write(rowindex, colindex, detalle[colname])
                    colindex += 1

                # Datos del movimientos.
                for colname in self.COL_ENCABEZADOS.values():
                    sheet.write(rowindex, colindex, movimiento[colname])
                    colindex += 1

                rowindex += 1
            
        return workbook.save(filename)            
