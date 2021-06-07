import re
import textwrap
import string

import numpy as np
import openpyxl


subscript_dict = {'sector': ['A', 'B', 'C', 'D'],
                  'region': ['Region1', 'Region2', 'Region3', 'Region4'],
                  'source': ['Gas', 'Oil', 'Coal'],
                  'out': ['Elec', 'Heat', 'Solid', 'Liquid'],
                  'age': ['0-4', '5-9', '10-14', '15-19', '20-24', '25-29',
                          '30-34', '35-39', '40-44', '45-49', '50-54', '55-59',
                          '60-64', '65-69', '70-74', '75-79', '80+'],
                  'regions9': ['EU27', 'UK', 'CHINA', 'EASTOC', 'INDIA',
                               'LATAM', 'RUS', 'USMCA', 'LROW'],
                  'gender':  ['female', 'male']
}

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
            excel = openpyxl.load_workbook(file)
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

class ExternalVariable(object):
    def __init__(self, var_name, dims, cell, description, units, file, sheet):
        self.var_name = var_name
        self.description = description
        self.units = units
        self.dims = dims
        self.file = file
        self.sheet = sheet
        self.dims_dict = {}
        self.cell = cell
        self.ref_row, self.ref_col = self._split_excel_cell(cell)

    def add_dimension(self, dim_name, read_along, sep=1):
        """
        Add dimension to an object.

        Parameters
        ----------
        dim_name: str
            Name of the subscript range to add, it must be an existing
            subscript range or subrange.

        read_along: str ('col' or 'row' or 'sheet' or 'file')
            Dimension to read along the subscript range.

        sep: int (for 'col' or 'row') or list (for 'sheet' or 'file')
            The separator between different elements of the subscript range.
            For read_along='row' or 'read_along='col' it must be the number
            of rows or columns between different elements of the subscript
            range. For read_along='sheet' or 'read_along='file' it must
            be the list of sheets or files where the subscript range.

        Returns
        -------
        None

        """
        self.dims_dict[dim_name] = (read_along, sep)

    def add_series(self, name, cell, read_along, length):
        """
        Add series element x or time for LOOKUPS or DATA type objects.

        Parameters
        ----------
        name: str
            Name of the series cellrange.

        cell: str
            Reference cell.

        read_along: str ('col' or 'row')
            Dimension to read along the series.

        length: int
            The length of the series.

        Returns
        -------
        None

        """
        self.series = {
            'name': name,
            'cell': cell,
            'read_along': read_along,
            'length': length}

        ref_row, ref_col = self._split_excel_cell(cell)
        if read_along == 'row':
            rows = [ref_row, ref_row + length - 1]
            cols = [ref_col, ref_col]
        elif read_along == 'col':
            rows = [ref_row, ref_row]
            cols = [ref_col, ref_col + length -1]
        else:
            raise ValueError(
                "read_along must be 'row' or 'col'."
            )

        # add series cellrange name without specifiying the sheet
        self.series['cellrange'] = '$%s$%s:$%s$%s' % (
            self._num_to_col(cols[0]), rows[0] + 1 ,
            self._num_to_col(cols[1]), rows[1] + 1)

    def update_series_cellranges(self, sheets, files):
        """
        Add sheets to the cell ranges of the series

        Parameters
        ----------
        sheets: set
            The set of the sheets to add cellranges.

        files: set
            The set of the files to add cellranges.

        Returns
        -------
        None

        """
        self.series['cellrange'] = [
            sheet + '!' + self.series['cellrange']
            for sheet in sheets for file in files]

        self.series['file'] = [
            file
            for sheet in sheets for file in files]

        self.series['sheet'] = [
            sheet
            for sheet in sheets for file in files]

    def build_boxes(self, elements, visited):
        """
        Using the information of the dims_dict, builds the cellrange
        boxes.

        Parameters
        ----------
        elements: dict
            Dictionary to save the information of the object. This
            object will be mutated by this function.
        visited: list
            List of the visited read_along elements with sep=1. It is
            used for specify the dimension of the series in DATA and LOOKUPS.

        Returns
        -------
        visited: list
            List of the visited read_along elements with sep=1,
            excluding 'file' and 'sheet'. Is needed for constants to
            add '*' for transpositions.

        """
        for dim in self.dims:
            read_along, step = self.dims_dict[dim]
            if step == 1:
                # append only subscript range name
                self.add_info(elements, [dim],
                              read_along,
                              len(subscript_dict[dim])-1)

                visited.append(read_along)

            elif isinstance(step, int):
                # append list of subscripts in subscript range
                self.add_info(elements, subscript_dict[dim],
                              read_along,
                              range(0, step*len(subscript_dict[dim]), step))
                # steps: [0, step, 2*step, ..., (n_subs-1)*step]
            else:
                # read along file of sheet
                # append list of subscripts in subscript range
                self.add_info(elements, subscript_dict[dim],
                              read_along,
                              step)
                visited.append(read_along)

        for dim in ['col', 'row']:
            if visited.count(dim) > 1:
                raise ValueError(
                    f"Two or more dimensions are defined along {dim}"
                    " with step 1.")

        for dim in ['file', 'sheet']:
            if visited.count(dim) == 0:
                # no dimension defined along sheet or file
                elements[dim] = [getattr(self, dim)] * len(elements[dim])
            elif visited.count(dim) == 1:
                # 1 dimension defined along sheet or file, remove it from
                # dim for transpositions in CONSTANTS
                visited.remove(dim)
            else:
                # 2 or more dimensions defined along sheet or file
                raise ValueError(
                    f"Two or more dimensions are defined along {dim}")

        # dictionary to manage cellrange names
        if len(elements['sheet']) * len(elements['file'])\
          / len(set(elements['sheet'])) * len(set(elements['file'])) != 1:
          # elements are repeated in the same sheet of the same file
          # need to add numbers to the end of the cellrange name
            added = {file: {sheet: 0 for sheet in set(elements['sheet'])}
                     for file in set(elements['file'])}
        else:
            added = None

        # convert cols to alpha
        elements['col'] = [[self._num_to_col(col) for col in element]
                           for element in elements['col']]

        # convert rows to excel numbering
        elements['row'] = [[row+1 for row in element]
                           for element in elements['row']]

        # writting information
        elements['cellrange'] = [
            '%s!$%s$%s:$%s$%s' % (sheet, cols[0], rows[0], cols[1], rows[1])
            for sheet, cols, rows in\
            zip(elements['sheet'], elements['col'], elements['row'])
        ]

        return visited, added

    @staticmethod
    def add_info(elements, subs, read_along, steps=None):
        """
        Combine several list with elements of a given list

        Parameters
        ----------
        elements: dict
            Dictionary with element information.

        subs: list
            List of current subscripts.

        read_along: str ('col' or 'row' or 'sheet' or 'file')
            Dimension to read along the subscript range.

        steps: int, ndarray or None (optional)
            Default is None.


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
        coords_to_duplicate = [
            along for along in ['col', 'row', 'file', 'sheet']
            if along != read_along]

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

    def write_cellranges(self, base_name, files, sheets, cellranges,
                         added=None):
        """
        Loop for writting several cellranges in excel file.

        Parameters
        ----------
        base_name: str
            The base name of the cellrange. If added is provided, a
            number will be append to the name. Otherwise, the base name
            will be used for the cellrange.

        files: list
            List of files to write each cellrange in.

        sheets: list
            List of files to write each cellrange in.

        cellranges: list
            List of cellranges to write.

        added: dict or None (optional)
            If dict it will be used to add number for the cellranges to
            avoid repeating them.

        Returns
        -------
        written_names: list
            The list of the written cell range names.

        """
        written_names = []
        for file, sheet, cellrange in zip(files,
                                          sheets,
                                          cellranges):
            if added:
                added[file][sheet] += 1
                name = f'{base_name}_{added[file][sheet]}'
            else:
                name = base_name

            self.write_cellrange(name, file, sheet, cellrange)
            written_names.append(name)

        return written_names

    @staticmethod
    def write_cellrange(name, file, sheet, cellrange):
        """
        Writes cellranges using openpyxl

        Parameters
        ----------
        name: str
            The name of the cellrange.

        file: str
            The name of the file to write the cellrange in.

        sheet: str
            The name of the sheet to write the cellrange in.

        cellrange: str
            The cellrange to write.

        Returns
        -------
        None

        """
        wb = Excels.read(file)
        sheetId = [sheetname_wb.lower() for sheetname_wb
                    in wb.sheetnames].index(sheet.lower())

        existing_names = wb.defined_names.localnames(sheetId)
        if name in existing_names and\
          wb.defined_names.get(name, sheetId).attr_text == cellrange:
            # cellrange already defined with same name and coordinates
            return

        new_range = openpyxl.workbook.defined_name.DefinedName(
            name, attr_text=cellrange, localSheetId=sheetId)
        wb.defined_names.append(new_range)

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
    def __init__(self, var_name, dims, cell, description='', units='',
                 file=None, sheet=None):
        super().__init__(var_name, dims, cell, description, units, file, sheet)

    def add_x(self, name, cell, read_along, length):
        """
        Add x for LOOKUPS type objects.

        Parameters
        ----------
        name: str
            Name of the x series cellrange.

        cell: str
            Reference cell.

        read_along: str ('col' or 'row')
            Dimension to read along the x series.

        length: int
            The length of the x series.

        Returns
        -------
        None

        """
        super().add_series(name, cell, read_along, length)

    def get_vensim(self):
        elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]]
        }

        elements[self.series['read_along']][0][1] += self.series['length'] - 1

        _, added = super().build_boxes(elements, [self.series['read_along']])

        super().update_series_cellranges(set(elements['sheet']),
                                         set(elements['file']))

        # write series cellranges
        super().write_cellranges(
            self.series['name'], self.series['file'],
            self.series['sheet'], self.series['cellrange'],
            None)

        # write data cellranges
        elements['cellname'] = super().write_cellranges(
            self.var_name, elements['file'],
            elements['sheet'], elements['cellrange'],
            added)

        # save changes and close Excel files
        Excels.save_and_close()

        # generate Vensim equations
        vensim_eqs = ""
        for subs, file, sheet, cellname in zip(elements['subs'],
                                               elements['file'],
                                               elements['sheet'],
                                               elements['cellname']):
            vensim_eq = f"""
            {self.var_name}[{', '.join(map(str, subs))}]=
            \tGET_DIRECT_LOOKUPS('{file}', '{sheet}', '{self.series['name']}', '{cellname}') ~~|"""

            vensim_eqs += vensim_eq

        vensim_eqs = textwrap.dedent(vensim_eqs)
        vensim_eqs = vensim_eqs[:-4]\
                     + f'\n\t~\t{self.units}'\
                     + f'\n\t~\t{self.description}'\
                     + '\n\t|'
        print(vensim_eqs)

class Data(ExternalVariable):
    def __init__(self, var_name, dims, cell, description='', units='',
                 file=None, sheet=None, interp=None):
        super().__init__(var_name, dims, cell, description, units, file, sheet)
        self.interp = interp

    def add_time(self, name, cell, read_along, length):
        """
        Add time for DATA type objects.

        Parameters
        ----------
        name: str
            Name of the time series cellrange.

        cell: str
            Reference cell.

        read_along: str ('col' or 'row')
            Dimension to read along the time series.

        length: int
            The length of the time series.

        Returns
        -------
        None

        """
        super().add_series(name, cell, read_along, length)

    def get_vensim(self):
        elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]]
        }

        elements[self.series['read_along']][0][1] += self.series['length'] - 1

        _, added = super().build_boxes(elements, [self.series['read_along']])

        super().update_series_cellranges(set(elements['sheet']),
                                         set(elements['file']))

        # write series cellranges
        super().write_cellranges(
            self.series['name'], self.series['file'],
            self.series['sheet'], self.series['cellrange'],
            None)

        # write data cellranges
        elements['cellname'] = super().write_cellranges(
            self.var_name, elements['file'],
            elements['sheet'], elements['cellrange'],
            added)

        # save changes and close Excel files
        Excels.save_and_close()

        # generate Vensim equations
        vensim_eqs = ""
        for subs, file, sheet, cellname in zip(elements['subs'],
                                               elements['file'],
                                               elements['sheet'],
                                               elements['cellname']):
            vensim_eq = f"""
            {self.var_name}[{', '.join(map(str, subs))}]"""
            if self.interp:
                # add keyword for interpolation method
                vensim_eq += f":{self.interp.upper()}::="
            else:
                vensim_eq += ":="
            vensim_eq += f"""
            \tGET_DIRECT_DATA('{file}', '{sheet}', '{self.series['name']}', '{cellname}') ~~|"""

            vensim_eqs += vensim_eq

        vensim_eqs = textwrap.dedent(vensim_eqs)
        vensim_eqs = vensim_eqs[:-4]\
                     + f'\n\t~\t{self.units}'\
                     + f'\n\t~\t{self.description}'\
                     + '\n\t|'
        print(vensim_eqs)

class Constants(ExternalVariable):
    def __init__(self, var_name, dims, cell, description='', units='',
                 file=None, sheet=None, **kwargs):
        super().__init__(var_name, dims, cell, description, units, file, sheet)
        self.transpose = False

    def get_vensim(self):
        elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]]
        }
        visited, added = self.build_boxes(elements, [])

        # transpose with *
        if visited in [["row"], ["col", "row"]]:
            self.transpose = True

        # write data cellranges
        elements['cellname'] = super().write_cellranges(
            self.var_name, elements['file'],
            elements['sheet'], elements['cellrange'],
            added)

        # save changes and close Excel files
        Excels.save_and_close()

        # generate Vensim equations
        vensim_eqs = ""
        for subs, file, sheet, cellname in zip(elements['subs'],
                                               elements['file'],
                                               elements['sheet'],
                                               elements['cellname']):
            if self.transpose:
                cellname += '*'
            vensim_eq = f"""
            {self.var_name}[{', '.join(map(str, subs))}]=
            \tGET_DIRECT_CONSTANTS('{file}', '{sheet}', '{cellname}') ~~|"""

            vensim_eqs += vensim_eq

        vensim_eqs = textwrap.dedent(vensim_eqs)
        vensim_eqs = vensim_eqs[:-4]\
                     + f'\n\t~\t{self.units}'\
                     + f'\n\t~\t{self.description}'\
                     + '\n\t|'
        print(vensim_eqs)






