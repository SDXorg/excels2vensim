import re

import numpy as np
import openpyxl
import string

subscript_dict = {'sector': ['A', 'B', 'C', 'D'],
                  'region': ['Region1', 'Region2', 'Region3', 'Region4'],
                  'source': ['Gas', 'Oil', 'Coal'],
                  'out': ['Elec', 'Heat', 'Solid', 'Liquid']}

class ExternalVariable(object):
    def __init__(self, var_name, dims, cell, file, sheet):
        self.var_name = var_name
        self.dims = dims
        self.file = file
        self.sheet = sheet
        self.dims_dict = {}
        self.cell = cell
        self.ref_row, self.ref_col = self._split_excel_cell(cell)

    def add_dimension(self, dim_name, read_along, sep=1):
        self.dims_dict[dim_name] = (read_along, sep)

    @staticmethod
    def add_info(elements, subs, read_along=None, steps=None):
        """
        Combine several list with elements of a given list

        Parameters
        ----------
        elements: dict

        subs: list

        Returns
        -------
        list_out: list of lists

        """
        # add the subs
        list_out = []
        for element1 in elements['subs']:
            for element2 in subs:
                list_out.append(element1 + [element2])

        elements['subs'] = list_out

        # dimension gives the table shape
        if isinstance(steps, int):
            list_out = []
            for current_coord in elements[read_along]:
                list_out.append(current_coord + np.array([0, steps]))

            elements[read_along] = list_out
            return

        # duplicate the rows, cols, file and sheet values if not given
        coords_to_duplicate = [along for along in ['col', 'row', 'file', 'sheet'] if along != read_along]

        for along in coords_to_duplicate:
            list_out = []
            for current_coord in elements[along]:
                for step in steps:
                    list_out.append(current_coord)
            elements[along] = list_out

        list_out = []
        if read_along in ['col', 'row']:
            # udpate cols or rows to read
            for current_coord in elements[read_along]:
                for step in steps:
                    list_out.append(current_coord + step)
            elements[read_along] = list_out
        else:
            # update file or sheet to read
            for current_coord in elements[read_along]:
                for step in steps:
                    list_out.append(step)
            elements[read_along] = list_out


    @staticmethod
    def _col_to_num(col):
        """
        FUNCTION FROM PYSD
        Transforms the column name to int

        Parameters
        ----------
        col: str
          Column name

        Returns
        -------
        int
          Column number
        """
        if len(col) == 1:
            return ord(col.upper()) - ord('A')
        elif len(col) == 2:
            left = ord(col[0].upper()) - ord('A') + 1
            right = ord(col[1].upper()) - ord('A')
            return left * (ord('Z')-ord('A')+1) + right
        else:
            left = ord(col[0].upper()) - ord('A') + 1
            center = ord(col[1].upper()) - ord('A') + 1
            right = ord(col[2].upper()) - ord('A')
            return left * ((ord('Z')-ord('A')+1)**2)\
                + center * (ord('Z')-ord('A')+1)\
                + right

    @staticmethod
    def _num_to_col(num):
        """
        From and old PR to PySD by Julien Malard
        """
        chars = []
        num += 1

        def divmod_excel(n):
            a, b = divmod(n, 26)
            if b == 0:
                return a - 1, b + 26
            return a, b

        while num > 0:
            num, d = divmod_excel(num)
            chars.append(string.ascii_uppercase[d-1])
        return ''.join(reversed(chars))

    def _split_excel_cell(self, cell):
        """
        FUNCTION FROM PYSD
        Splits a cell value given in a string.
        Returns None for non-valid cell formats.

        Parameters
        ----------
        cell: str
          Cell like string, such as "A1", "b16", "AC19"...
          If it is not a cell like string will return None.

        Returns
        -------
        row number, column number: int, int
          If the cell input is valid. Both numbers are given in Python
          enumeration, i.e., first row and first column are 0.

        """
        split = re.findall(r'\d+|\D+', cell)
        try:
            # check that we only have two values [column, row]
            assert len(split) == 2
            # check that the column name has no special characters
            assert not re.compile('[^a-zA-Z]+').search(split[0])
            # check that row number is not 0
            assert int(split[1]) != 0
            # the column name has as maximum 3 letters
            assert len(split[0]) <= 3
            return int(split[1])-1, self._col_to_num(split[0])
        except AssertionError:
            return

class Lookups(ExternalVariable):
    def __init__(self, var_name, dims, file=None, sheet=None):
        super().__init__(var_name, dims, file, sheet)


class Data(ExternalVariable):
    def __init__(self, var_name, dims, file=None, sheet=None, interp=None):
        super().__init__(var_name, dims, file, sheet)
        self.interp = interp

class Constants(ExternalVariable):
    def __init__(self, var_name, dims, cell, file=None, sheet=None):
        super().__init__(var_name, dims, cell, file, sheet)

    def get_vensim(self):
        elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]]
        }
        for dim in self.dims:
            read_along, step = self.dims_dict[dim]
            if step == 1:
                # append only subscript range name
                self.add_info(elements, [dim], read_along, len(subscript_dict[dim])-1)

            elif isinstance(step, int):
                # append list of subscripts in subscript range
                self.add_info(elements, subscript_dict[dim], read_along, range(0, step*len(subscript_dict[dim]), step))
                # steps: [0, step, 2*step, ..., (n_subs-1)*step]
            else:
                # append list of subscripts in subscript range
                self.add_info(elements, subscript_dict[dim], read_along, step)

        if not elements['file'][0]:
            elements['file'] = [self.file] * len(elements['file'])
        if not elements['sheet'][0]:
            elements['sheet'] = [self.sheet] * len(elements['sheet'])

        # convert cols to alpha
        elements['col'] = [[self._num_to_col(col) for col in element] for element in elements['col']]

        # convert rows to excel numering
        elements['row'] = [[row+1 for row in element] for element in elements['row']]

        # writting information
        elements['write'] = [
            '%s!$%s$%s:$%s$%s' % (sheet, cols[0], rows[0], cols[1], rows[1])
            for sheet, cols, rows in\
            zip(elements['sheet'], elements['col'], elements['row'])
        ]

        counter = 1

        for file, sheet, write in zip(elements['file'], elements['sheet'], elements['write']):
            wb = openpyxl.load_workbook(file)
            sheetId = [sheetname_wb.lower() for sheetname_wb
                       in wb.sheetnames].index(sheet.lower())
            new_range = openpyxl.workbook.defined_name.DefinedName(f'{self.var_name}_{counter}', attr_text=write, localSheetId=sheetId)
            wb.defined_names.append(new_range)
            wb.save(file)
            counter += 1






