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
    workbook = None
    data = []

    def Open(self, filename):
        self.workbook = Excel.open(filename)
        return self.workbook

    def __read(self):
        self.data = Excel.read(self.workbook)



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
