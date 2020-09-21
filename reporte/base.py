import csv
import xlrd
import xlwt



class ReporteError(Exception):
    """
    Excepción de este módulo.
    """



class Reporte:
    """
    Instancía de cualquier reporte en formato Excel.
    """
    ROW_ENCABEZADOS = 0
    COL_ENCABEZADOS = dict()
    ROW_FACTURA_DETALLE = None
    COL_FACTURA_DETALLE = dict()
    ENCABEZADOS = list()
    ENCABEZADOS_KEYS = list()
    CSV_DELIMITER = ";"

    def __init__(self, filename):
        self.workbook = None
        self.csv = None
        self.rawdata = [] #
        self.data = []
        self.Open(filename)
        self.Read()
        self.clean()

    def __iter__(self):
        for item in self.data:
            yield item

    def __getitem__(self, index):
        return self.data[index]

    def __len__(self):
        return len(self.data)

    def clean(self):
        """
        Agregue este método a cada reporte para generar el valor de self.data.
        """

    def Open(self, filename):
        ext = filename.split(".")[-1].lower()

        if ext in ("xls",):
            self.workbook = Excel.open(filename)
            return self.workbook
        elif ext in ("csv",):
            f = open(filename, "r")
            self.csv = csv.reader(f, delimiter=self.CSV_DELIMITER)
            return self.csv
        else:
            raise ReporteError("Formato de archivo no soportado.")

    def Read(self):
        if (self.workbook != None):
            self.rawdata = Excel.read(self.workbook)
            self.rawdata = Excel.cleanrows(self.rawdata)
        elif (self.csv != None):
            self.rawdata = self.csv
        else:
            self.rawdata = []

    def ExportToExcel(self, filename=None):
        """
        Exporta la data generada por este reporte, a formato Excel.
        """

        if (not filename):
            filename = f"{self.__class__.__name__}.xls"

        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("blog.unolet.com")
        rowindex = 0

        colindex = 0
        for name in self.ENCABEZADOS:
            sheet.write(rowindex, colindex, name)
            colindex += 1

        rowindex += 1
        for detalle in self.data:
            movimientos = detalle["movimientos"]

            for movimiento in movimientos:
                colindex = 0

                # Datos de la factura.
                for key in self.ENCABEZADOS_KEYS:
                    
                    try:
                        value = detalle[key]
                    except (KeyError):
                        value = movimiento[key]

                    sheet.write(rowindex, colindex, value)
                    colindex += 1
                rowindex += 1
        
        workbook.save(filename)         



class Excel():

    @classmethod
    def open(self, filename):
        """
        Abre el archivo Excel.
        """
        return xlrd.open_workbook(filename)

    @classmethod
    def save(self, filename, data):
        """
        Guarda el archivo en formato Excel.
        """
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("UNOLET")

        for y in range(len(data)):
            row = data[y]
            for x in range(len(row)):
                cell = row[x]
                sheet.write(x + 1, y + 1, cell)

    @classmethod
    def read(self, workbook, sheetindex=0, rowsrange=None):
        """
        Lee la informaición de la hoja Excel en el rango de filas especificado, y 
        retorna un array con dicha informaición.

        Parameters:
            workbook: xlrd.open_workbook()
            sheetindex (int): workbook.sheet_by_index(sheetindex)
            rowrange (int): Cantidad de filas.

        Returns:
            list:
        """
        sheet = workbook.sheet_by_index(sheetindex)

        if (rowsrange is None):
            rowsrange = sheet.nrows

        out = []
        for index in range(rowsrange):
            out.append(sheet.row_values(index))
        return out

    @classmethod
    def clean(self, data):
        """
        Elimida las celdas vacias.
        """
        out = []
        for row in data:
            row = [cell for cell in row if cell != ""]
            if row:
                out.append(row)
        return out

    @classmethod
    def cleanrows(self, data):
        """
        Elimina las filas donde todas sus celdas están en blanco.
        """
        out = []
        for row in data:
            if not [c for c in row if c != ""]:
                continue
            out.append(row)
        return out
