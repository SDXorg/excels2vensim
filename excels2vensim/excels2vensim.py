"""
Functions for parsing the subscript from a .mdl file using PySD.
"""
import warnings
import re
import textwrap
import string
import json

import numpy as np
from openpyxl.workbook.defined_name import DefinedName

from .utils.excels import Excels
from .utils.subscripts import Subscripts


class ExternalVariable(object):
    def __init__(self, var_name, dims, cell, description, units, file, sheet):
        self.var_name = var_name.strip()
        self.base_name = self._clean_identifier(self.var_name )
        if self.base_name != self.var_name:
            warnings.warn(
                f"The name of the variable '{self.var_name}' has special "
                + f"characters. '{self.base_name}' will be used for "
                + "cellrange names.")
        self.description = description.strip()
        self.units = units.strip()
        self.dims = [dim.strip() for dim in dims]
        self.file = file
        self.sheet = sheet
        self.dims_dict = {}
        self.cell = cell
        self.ref_row, self.ref_col = self._split_excel_cell(cell)
        self.subscripts_warns = set()

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
        if dim_name.strip() not in Subscripts.get_ranges():
            raise ValueError(
                f"\n'{dim_name}' is not in the list of subscript ranges:\n\t"
                + str(Subscripts.get_ranges()))
        elif read_along not in ['col', 'row', 'sheet', 'file']:
            raise ValueError(
                "\nread_along must be 'row', 'col', 'sheet' or 'file'."
            )

        self.dims_dict[dim_name.strip()] = (read_along, sep)

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
        cname = self._clean_identifier(name)
        if name.strip() != cname:
            warnings.warn(
                f"The name of the interpolation dimension '{name.strip()}'"
                f" has special characters. '{cname}' will be used for "
                + "cellrange names.")

        self.series = {
            'name': cname,
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
                "\nread_along must be 'row' or 'col'."
            )

        # add series cellrange name without specifiying the sheet
        self.series['cellrange'] = '$%s$%s:$%s$%s' % (
            self._num_to_col(cols[0]), rows[0] + 1 ,
            self._num_to_col(cols[1]), rows[1] + 1)

    def _update_series_cellranges(self, sheets, files):
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

        self.series['name'] =\
             [self.series['name']] * len(self.series['cellrange'])

    def _build_boxes(self, visited):
        """
        Using the information of the dims_dict, builds the cellrange
        boxes.

        Parameters
        ----------
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
                self._add_info([dim],
                               read_along,
                               len(Subscripts.get(dim))-1)

                visited.append(read_along)

            elif isinstance(step, int):
                # append list of subscripts in subscript range
                self._add_info(Subscripts.get(dim),
                               read_along,
                               range(0, step*len(Subscripts.get(dim)), step))
                # steps: [0, step, 2*step, ..., (n_subs-1)*step]
            else:
                # read along file of sheet
                # append list of subscripts in subscript range
                self._add_info(Subscripts.get(dim),
                               read_along,
                               step)
                visited.append(read_along)

        # raise warnings only once per dimension
        for swarn in self.subscripts_warns:
            warnings.warn(swarn)

        for dim in ['col', 'row']:
            if visited.count(dim) > 1:
                raise ValueError(
                    f"\nTwo or more dimensions are defined along {dim}"
                    " with step 1.")

        for dim in ['file', 'sheet']:
            if visited.count(dim) == 0:
                # no dimension defined along sheet or file
                self.elements[dim] =\
                    [getattr(self, dim)] * len(self.elements[dim])
            elif visited.count(dim) == 1:
                # 1 dimension defined along sheet or file, remove it from
                # dim for transpositions in CONSTANTS
                visited.remove(dim)
            else:
                # 2 or more dimensions defined along sheet or file
                raise ValueError(
                    f"\nTwo or more dimensions are defined along {dim}.")

        # convert cols to alpha
        self.elements['col'] = [[self._num_to_col(col) for col in element]
                           for element in self.elements['col']]

        # convert rows to excel numbering
        self.elements['row'] = [[row+1 for row in element]
                           for element in self.elements['row']]

        # writting information
        self.elements['cellrange'] = [
            '%s!$%s$%s:$%s$%s' % (sheet, cols[0], rows[0], cols[1], rows[1])
            for sheet, cols, rows in\
            zip(self.elements['sheet'],
                self.elements['col'], self.elements['row'])
        ]

        return visited

    def _add_info(self, subs, read_along, steps=None):
        """
        Combine several list with elements of a given list

        Parameters
        ----------
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
        for element1 in self.elements['subs']:
            for element2 in subs:
                list_out.append(element1 + [element2])

        self.elements['subs'] = list_out

        # dimension gives the table shape
        if isinstance(steps, int):
            list_out = []
            for current_coord in self.elements[read_along]:
                list_out.append(current_coord + np.array([0, steps]))

            self.elements[read_along] = list_out
            return

        # duplicate the rows, cols, file and sheet values if not given
        coords_to_duplicate = [
            along for along in ['col', 'row', 'file', 'sheet']
            if along != read_along]

        for along in coords_to_duplicate:
            list_out = []
            for current_coord in self.elements[along]:
                for step in steps:
                    list_out.append(current_coord)
            self.elements[along] = list_out

        list_out = []
        names_out = []
        if read_along in ['col', 'row']:
            # udpate cols or rows to read
            for current_coord in self.elements[read_along]:
                for step in steps:
                    list_out.append(current_coord + step)
            self.elements[read_along] = list_out
            for current_name in self.elements['cellname']:
                for sub in subs:
                    subc = self._clean_identifier(sub)
                    if subc != sub.strip():
                        self.subscripts_warns.add(
                             f"The name of the subscript '{sub.strip()}'"
                             + f" has special characters. '{subc}' will be"
                             + " used for cellrange names.")
                    names_out.append(current_name + '_' + subc)
            self.elements['cellname'] = names_out
        else:
            # update file or sheet to read
            for (current_name, current_coord) in zip(self.elements['cellname'],
                                                     self.elements[read_along]):
                for step in steps:
                    list_out.append(step)
                    names_out.append(current_name)
            self.elements[read_along] = list_out
            self.elements['cellname'] = names_out

    def _write_cellranges(self, names, files, sheets, cellranges):
        """
        Loop for writting several cellranges in excel file.

        Parameters
        ----------
        names: list
            List of names of cellranges.

        files: list
            List of files to write each cellrange in.

        sheets: list
            List of sheets to write each cellrange in.

        cellranges: list
            List of cellranges to write.

        Returns
        -------
        written_names: list
            The list of the written cell range names.

        """
        written_names = []
        for name, file, sheet, cellrange in zip(names,
                                                files,
                                                sheets,
                                                cellranges):

            self._write_cellrange(name, file, sheet, cellrange, self.force)
            written_names.append(name)

        return written_names

    @staticmethod
    def _write_cellrange(name, file, sheet, cellrange, force):
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
        if name in existing_names:
            if wb.defined_names.get(name, sheetId).attr_text == cellrange:
                # cellrange already defined with same name and coordinates
                return
            elif force:
                wb.defined_names.delete(name, sheetId)
            else:
                raise ValueError(
                    f"\nTrying to write a cellrange with name '{name}' at "
                    + f"'{cellrange}'. However, '{name}' already exist in "
                    + f"'{wb.defined_names.get(name, sheetId).attr_text}'\n"
                    + "Use force=True to overwrite it.")

        new_range = DefinedName(
            name, attr_text=cellrange, localSheetId=sheetId)
        wb.defined_names.append(new_range)

    @staticmethod
    def _col_to_num(col):
        """
        Transforms the column name to int.

        Parameters
        ----------
        col: str
          Column name.

        Returns
        -------
        int
          Column number starting from 0.

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
        Transforms the column number to name.

        Parameters
        ----------
        num: int
          Column number starting from 0.

        Returns
        -------
        str
          Column name.

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

    @classmethod
    def _split_excel_cell(cls, cell):
        """
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
            return int(split[1])-1, cls._col_to_num(split[0])
        except AssertionError:
            return

    @staticmethod
    def _clean_identifier(string):
        """
        Remove invalid characters and spaces from a string.

        Parameters
        ----------
        string: str
            Original string.

        Returns
        -------
        str
            Clean string.
        """
        return re.sub('[^A-Za-z0-9]+', '_', string).strip('_')


class Lookups(ExternalVariable):
    """
    Class for creating GET DIRECT/XLS LOOKUPS equations and cellranges.

    Parameters
    ----------
    var_name: str
        The name of the variable in Vensim code and basestring for
        cellrange names.

    dims: list
        List of the dimensions of the variable in the same order that
        will be created the variable in the Vensim code.

    cell: str
        Reference cell of the data. First cellw with numeric values
        (upper-left corner).

    description: str (optional)
        Description to include in the Vensim equations. By default no
        description will be included.

    units: str (optional)
        Units to include in the Vensim equations. By default no
        units will be included.

    file: str (optional)
        File where the data is. This argument is mandatory unless a
        subscript range is defined across several files. Default is None.

    sheet: str (optional)
        Sheet where the data is. This argument is mandatory unless a
        subscript range is defined across several sheets. Default is None.

    """
    def __init__(self, var_name, dims, cell, description='', units='',
                 file=None, sheet=None, **kwargs):
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

    def execute(self, force=False, loading='DIRECT'):
        """
        Get vensim equations and write cell range names in the Excel file.

        Parameters
        ----------
        force: bool (optional)
            If True and trying and tryting to write a cell range name
            that already exist in other positions it will overwrite it
            (not recommended). If False it will return and error when
            trying to write the new cellrange name. Default is False.

        loading: str (optional)
            Vensing GET loading type it can be 'DIRECT' or 'XLS'.
            Default is 'DIRECT'.

        Returns
        -------
        vensim_eqs: str
            The string of Vensim equations to copy in the model .mdl file.

        """
        # force removal of conflicting cellrange names
        self.force = force

        vensim_eqs = self.get_vensim(loading)

        # write series cellranges
        super()._write_cellranges(
            self.series['name'], self.series['file'],
            self.series['sheet'], self.series['cellrange'])

        # write data cellranges
        super()._write_cellranges(
            self.elements['cellname'], self.elements['file'],
            self.elements['sheet'], self.elements['cellrange'])

        # save changes and close Excel files
        Excels.save_and_close()

        return vensim_eqs

    def get_vensim(self, force=False, loading='DIRECT'):
        """
        Get vensim equations and write cell range names in the Excel file.

        Parameters
        ----------
        force: bool (optional)
            If True and trying and tryting to write a cell range name
            that already exist in other positions it will overwrite it
            (not recommended). If False it will return and error when
            trying to write the new cellrange name. Default is False.

        loading: str (optional)
            Vensing GET loading type it can be 'DIRECT' or 'XLS'.
            Default is 'DIRECT'.

        Returns
        -------
        vensim_eqs: str
            The string of Vensim equations to copy in the model .mdl file.

        """
        # force removal of conflicting cellrange names
        self.force = force

        self.elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]],
            'cellname': [self.base_name]
        }

        self.elements[self.series['read_along']][0][1] +=\
            self.series['length'] - 1

        super()._build_boxes([self.series['read_along']])

        super()._update_series_cellranges(set(self.elements['sheet']),
                                         set(self.elements['file']))

        # generate Vensim equations
        vensim_eqs = ""
        for subs, file, sheet, cellname in zip(self.elements['subs'],
                                               self.elements['file'],
                                               self.elements['sheet'],
                                               self.elements['cellname']):
            vensim_eq = f"""
            {self.var_name}[{', '.join(map(str, subs))}]=
            \tGET_{loading}_LOOKUPS('{file}', '{sheet}', '{self.series['name'][0]}', '{cellname}') ~~|"""

            vensim_eqs += vensim_eq

        vensim_eqs = textwrap.dedent(vensim_eqs)
        vensim_eqs = vensim_eqs[:-4]\
                     + f'\n\t~\t{self.units}'\
                     + f'\n\t~\t{self.description}'\
                     + '\n\t|'

        return vensim_eqs

class Data(ExternalVariable):
    """
    Class for creating GET DIRECT/XLS DATA equations and cellranges.

    Parameters
    ----------
    var_name: str
        The name of the variable in Vensim code and basestring for
        cellrange names.

    dims: list
        List of the dimensions of the variable in the same order that
        will be created the variable in the Vensim code.

    cell: str
        Reference cell of the data. First cellw with numeric values
        (upper-left corner).

    description: str (optional)
        Description to include in the Vensim equations. By default no
        description will be included.

    units: str (optional)
        Units to include in the Vensim equations. By default no
        units will be included.

    file: str (optional)
        File where the data is. This argument is mandatory unless a
        subscript range is defined across several files. Default is None.

    sheet: str (optional)
        Sheet where the data is. This argument is mandatory unless a
        subscript range is defined across several sheets. Default is None.

    interp: str or None (optional)
            Keyword of the interpolation method to use with DATA. It can be
            any keyword accepted by Vensim 'interpolate', 'look forward',
            'keep backward' or 'raw'. If None, no keyword will be added,
            Vensim will use the default interpolation method ('interpolate').
            Default is None.

    """
    def __init__(self, var_name, dims, cell, description='', units='',
                 file=None, sheet=None, interp=None, **kwargs):
        super().__init__(var_name, dims, cell, description, units, file, sheet)

        if interp:
            # conver interp to vensim notation
            self.interp = interp.strip().upper().replace('_', ' ')
            if self.interp not in ['INTERPOLATE', 'RAW',
                                   'HOLD BACKWARD', 'LOOK FORWARD']:
                raise ValueError(
                    "\ninterp must be 'interpolate', 'raw', "
                    "'hold backward' or 'look forward'")
        else:
            self.interp = None

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

    def execute(self, force=False, loading='DIRECT'):
        """
        Get vensim equations and write cell range names in the Excel file.

        Parameters
        ----------
        force: bool (optional)
            If True and trying and tryting to write a cell range name
            that already exist in other positions it will overwrite it
            (not recommended). If False it will return and error when
            trying to write the new cellrange name. Default is False.

        loading: str (optional)
            Vensing GET loading type it can be 'DIRECT' or 'XLS'.
            Default is 'DIRECT'.

        Returns
        -------
        vensim_eqs: str
            The string of Vensim equations to copy in the model .mdl file.

        """
        # force removal of conflicting cellrange names
        self.force = force

        vensim_eqs = self.get_vensim(loading)

        # write series cellranges
        super()._write_cellranges(
            self.series['name'], self.series['file'],
            self.series['sheet'], self.series['cellrange'])

        # write data cellranges
        super()._write_cellranges(
            self.elements['cellname'], self.elements['file'],
            self.elements['sheet'], self.elements['cellrange'])

        # save changes and close Excel files
        Excels.save_and_close()

        return vensim_eqs

    def get_vensim(self, force=False, loading='DIRECT'):
        """
        Get vensim equations and write cell range names in the Excel file.

        Parameters
        ----------
        force: bool (optional)
            If True and trying and tryting to write a cell range name
            that already exist in other positions it will overwrite it
            (not recommended). If False it will return and error when
            trying to write the new cellrange name. Default is False.

        loading: str (optional)
            Vensing GET loading type it can be 'DIRECT' or 'XLS'.
            Default is 'DIRECT'.

        Returns
        -------
        vensim_eqs: str
            The string of Vensim equations to copy in the model .mdl file.

        """
        # force removal of conflicting cellrange names
        self.force = force

        self.elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]],
            'cellname': [self.base_name]
        }

        self.elements[self.series['read_along']][0][1] +=\
            self.series['length'] - 1

        super()._build_boxes([self.series['read_along']])

        super()._update_series_cellranges(set(self.elements['sheet']),
                                         set(self.elements['file']))

        # generate Vensim equations
        vensim_eqs = ""
        for subs, file, sheet, cellname in zip(self.elements['subs'],
                                               self.elements['file'],
                                               self.elements['sheet'],
                                               self.elements['cellname']):
            vensim_eq = f"""
            {self.var_name}[{', '.join(map(str, subs))}]"""
            if self.interp:
                # add keyword for interpolation method
                vensim_eq += f":{self.interp}::="
            else:
                vensim_eq += ":="
            vensim_eq += f"""
            \tGET_{loading}_DATA('{file}', '{sheet}', '{self.series['name'][0]}', '{cellname}') ~~|"""

            vensim_eqs += vensim_eq

        vensim_eqs = textwrap.dedent(vensim_eqs)
        vensim_eqs = vensim_eqs[:-4]\
                     + f'\n\t~\t{self.units}'\
                     + f'\n\t~\t{self.description}'\
                     + '\n\t|'

        return vensim_eqs

class Constants(ExternalVariable):
    """
    Class for creating GET DIRECT/XLS CONSTANTS equations and cellranges.

    Parameters
    ----------
    var_name: str
        The name of the variable in Vensim code and basestring for
        cellrange names.

    dims: list
        List of the dimensions of the variable in the same order that
        will be created the variable in the Vensim code.

    cell: str
        Reference cell of the data. First cellw with numeric values
        (upper-left corner).

    description: str (optional)
        Description to include in the Vensim equations. By default no
        description will be included.

    units: str (optional)
        Units to include in the Vensim equations. By default no
        units will be included.

    file: str (optional)
        File where the data is. This argument is mandatory unless a
        subscript range is defined across several files. Default is None.

    sheet: str (optional)
        Sheet where the data is. This argument is mandatory unless a
        subscript range is defined across several sheets. Default is None.

    """
    def __init__(self, var_name, dims, cell, description='', units='',
                 file=None, sheet=None, **kwargs):
        super().__init__(var_name, dims, cell, description, units, file, sheet)
        self.transpose = False

    def execute(self, force=False, loading='DIRECT'):
        """
        Get vensim equations and write cell range names in the Excel file.

        Parameters
        ----------
        force: bool (optional)
            If True and trying and tryting to write a cell range name
            that already exist in other positions it will overwrite it
            (not recommended). If False it will return and error when
            trying to write the new cellrange name. Default is False.

        loading: str (optional)
            Vensing GET loading type it can be 'DIRECT' or 'XLS'.
            Default is 'DIRECT'.

        Returns
        -------
        vensim_eqs: str
            The string of Vensim equations to copy in the model .mdl file.

        """
        # force removal of conflicting cellrange names
        self.force = force

        vensim_eqs = self.get_vensim(loading)

        # write data cellranges
        super()._write_cellranges(
            self.elements['cellname'], self.elements['file'],
            self.elements['sheet'], self.elements['cellrange'])

        # save changes and close Excel files
        Excels.save_and_close()

        return vensim_eqs

    def get_vensim(self, loading='DIRECT'):
        """
        Get vensim equations.

        Parameters
        ----------
        loading: str (optional)
            Vensing GET loading type it can be 'DIRECT' or 'XLS'.
            Default is 'DIRECT'.

        Returns
        -------
        vensim_eqs: str
            The string of Vensim equations to copy in the model .mdl file.

        """
        self.elements = {
            'row': [np.array([self.ref_row, self.ref_row], dtype=int)],
            'col': [np.array([self.ref_col, self.ref_col], dtype=int)],
            'subs': [[]],
            'sheet': [[]],
            'file': [[]],
            'cellname': [self.base_name]
        }

        visited = self._build_boxes([])

        # transpose with *
        if visited in [["row"], ["col", "row"]]:
            self.transpose = True



        # generate Vensim equations
        vensim_eqs = ""
        for subs, file, sheet, cellname in zip(self.elements['subs'],
                                               self.elements['file'],
                                               self.elements['sheet'],
                                               self.elements['cellname']):
            if self.transpose:
                cellname += '*'
            vensim_eq = f"""
            {self.var_name}[{', '.join(map(str, subs))}]=
            \tGET_{loading}_CONSTANTS('{file}', '{sheet}', '{cellname}') ~~|"""

            vensim_eqs += vensim_eq

        vensim_eqs = textwrap.dedent(vensim_eqs)
        vensim_eqs = vensim_eqs[:-4]\
                     + f'\n\t~\t{self.units}'\
                     + f'\n\t~\t{self.description}'\
                     + '\n\t|'

        return vensim_eqs


def load_from_json(json_file):
    """
    Run the features using a JSON file.

    Prameters
    ---------
    json_file: str
        Name of the JSON file with the needed information.

    Returns
    -------
    str
        The equations to copy in the Vensim model file.

    """
    with open(json_file) as file:
        vars_dict = json.load(file)

    return execute(vars_dict)


def execute(vars_dict):
    """
    Run the features using a dictionary.

    Prameters
    ---------
    vars_dict
        Python dictionary with the needed information.

    Returns
    -------
    str
        The equations to copy in the Vensim model file.

    """
    eqs = []

    for var, info in vars_dict.items():

        if info['type'].lower() == 'constants':
            # create object
            obj = Constants(var, **info)

        elif info['type'].lower() == 'lookups':
            # create object
            obj = Lookups(var, **info)
            # add x series
            obj.add_x(**info['x'])

        elif info['type'].lower() == 'data':
            # create object
            obj = Data(var, **info)
            # add time series
            obj.add_time(**info['time'])

        else:
            raise ValueError(
                f"\n Invalid type of variable '{info['type']}' for '{var}'."
                + " It must be 'constants', 'lookups' or 'data'.")

        # add dimensions
        for dimension, along in info['dimensions'].items():
            obj.add_dimension(dimension, *along)

        if hasattr(info, 'force'):
            force = info['force']
        else:
            force = False
        if hasattr(info, 'loading'):
            loading = info['loading']
        else:
            loading = 'DIRECT'
        eqs.append(obj.execute(force=force, loading=loading))

    return '\n'.join(eqs)