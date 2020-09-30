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
        super().__init__(filename)

    def clean(self):
        """
        Limpia los datos del reporte Sicflex 'Ventas/Ventas por Fecha/Factura 
        (Detallada)', según la disposición de sus valores y retorna un listado 
        de dicionarios con los nuevos valores ordenados.
        """
        data = self.rawdata
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
        self.data = out
        return out

    
    def SetVendedores(self, filename):
        """
        Los datos extras están contenidos en otro reporte en formato CSV.
        SicflexListadoDeFacturas
        """
        new_data = self.data.copy()
        facturas = SicflexListadoDeFacturas(filename)

        progress = 0
        length = len(new_data)

        print(f"Estableciendo los vendedores: {length} filas en total.")

        for detalle in new_data:
            detalle["vendedor"] = ""
            almacen, tipo, numero = detalle["numero"].split("-")
            numero = int(numero)
            detalle_numero = f"{almacen}-{tipo}-{numero}"

            for factura in facturas:
                try:
                    number = f"{factura['almacen'].zfill(2)}-{factura['tipo']}-{int(factura['numero'])}"
                except (ValueError):
                    continue

                if detalle_numero == number:
                    detalle["vendedor"] = factura["usuario_creo"]
                    break
            
            if not detalle["vendedor"]:
                print(f"Error. No se encontró el vendedor para '{detalle_numero}'.")

            print(f"------ {progress} de {length} | {detalle['vendedor']}", end="\r")
            progress += 1

        self.ENCABEZADOS = list(self.ENCABEZADOS) + ["Vendedor"]
        self.ENCABEZADOS_KEYS = list(self.ENCABEZADOS_KEYS) + ["vendedor"]
        self.data = new_data
        return self.data


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
    
    




class SicflexListadoDeFacturas(Reporte):
    """
    Archivo en formato CSV extraido del listado de facturas en Sicflex.

    Este archivo contiene información de cada factura, número, vendedor, fecha...
    """
    ROW_ENCABEZADOS = 0
    COL_ENCABEZADOS = {
        0: "almacen",
        1: "tipo",
        2: "numero",
        3: "ncf",
        4: "nif",
        5: "usuario_creo",
    }
    ROW_FACTURA_DETALLE = None
    COL_FACTURA_DETALLE = {}
    ENCABEZADOS = ("Almacén", "Tipo", "Número", "NCF", "NIF", "Vendedor")
    ENCABEZADOS_KEYS = ("almacen", "tipo", "numero", "ncf", "nif", "usuario_creo")

    def __init__(self, filename):
        super().__init__(filename)

    def clean(self):
        out = []
        rows = list(self.csv)
        length = len(rows)
        progress = 0
        print(f"Cargando {self.__class__.__name__}. {length} filas a cargar.")

        for row in rows:

            if not row:
                progress += 1
                continue

            print(f"Fila {progress} de {length}.", end="\r")
            item = {}
            for index, name in self.COL_ENCABEZADOS.items():
                try:
                    item[name] = row[index]
                except (IndexError):
                    print(f"IndexError: {progress=}, {index=}, {row=}")
            out.append(item)

            progress += 1
        self.data = out
        return self.data


