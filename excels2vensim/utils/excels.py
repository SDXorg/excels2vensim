"""
Excel files manager class.
"""
from openpyxl import load_workbook

class Excels():
    """
    Class to save the read Excel files and thus avoid double reading
    """
    _Excels = {}

    @classmethod
    def read(cls, file):
        """
        Read the Excel file using OpenPyXL or return the previously read one
        """
        if file in cls._Excels:
            return cls._Excels[file]
        else:
            excel = load_workbook(file)
            cls._Excels[file] = excel
            return excel

    @classmethod
    def save_and_close(cls):
        """
        Saves and closes the Excel files
        """
        for file, wb in cls._Excels.items():
            wb.save(file)
            wb.close()

        cls._Excels = {}
