import os
import json
import tkinter as tkk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from excels2vensim import Subscripts, execute

class Application(tkk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(pady = 10, padx = 10)
        self.all_subs = list()
        self.non_selected_subs = self.all_subs.copy()
        self.subs = []
        self.description = ''
        self.element_dict = dict()
        self.create_vars()
        self.subscript_selector()
        self.main_window()


    def create_vars(self):

        self.var_info = tkk.StringVar(self)

        # Type of variable
        self.var_type = tkk.StringVar(self)
        self.var_type.trace(
            "w", lambda name, index, mode: self.update_type())

        # type of interp for data
        self.var_type_data = tkk.StringVar(self)
        self.var_type_data.set("none")
        self.var_type_data.trace(
            "w", lambda name, index, mode: self.update_info())
        self.var_type_data_drop_down = None

        # Name
        self.var_name = tkk.StringVar(self)
        self.var_name.trace(
            "w", lambda name, index, mode: self.update_info())

        # Loading
        self.var_loading = tkk.StringVar(self)
        self.var_loading.set("DIRECT")
        self.var_loading.trace(
            "w", lambda name, index, mode: self.update_info())


        # Subscript name
        self.var_sub = tkk.StringVar(self)
        self.var_sub.trace(
            "w", lambda name, index, mode: self.update_list(self.non_selected_subs))

        self.var_subs_info = tkk.StringVar(self)

        # Units
        self.var_units = tkk.StringVar(self)

        # File
        self.var_file = tkk.StringVar(self)
        self.var_file.trace(
            "w", lambda name, index, mode: self.update_info())

        # Sheet
        self.var_sheet = tkk.StringVar(self)
        self.var_sheet.trace(
            "w", lambda name, index, mode: self.update_info())

        # Reference cell
        self.var_cell = tkk.StringVar(self)

        # Subscript along
        self.var_along = tkk.StringVar(self)

        # Force
        self.var_force = tkk.BooleanVar(self)
        self.var_force.set(False)
        self.var_force.trace(
            "w", lambda name, index, mode: self.update_force())

        self.force_warning = None

    def subscript_selector(self):
        mdl_types = '*.mdl *.MDL'
        json_types = '*.json *.JSON'

        file_name = askopenfilename(
                title="Select external data file",
                filetypes=(('.mdl and .json files',
                            ' '.join([mdl_types, json_types])),
                           ('.mdl files', mdl_types),
                           ('.json files', json_types),
                           ("All files", '*'))
        )

        if file_name.lower().endswith('.mdl'):
            tkk.Label(self,
                      text=f"Reading {file_name}...\n"
                           + "This may take some seconds").grid(
                row = 0, column = 0, padx=10, pady=10)

        cwd = os.path.dirname(file_name)
        print(f"Setting current working directory to: {cwd}")
        os.chdir(cwd)

        Subscripts.read(file_name)
        self.clean()
        self.all_subs = list(Subscripts.get_ranges())
        self.non_selected_subs = self.all_subs.copy()

    def main_window(self):

        tkk.Label(self, textvariable=self.var_info).grid(row=0, column=0, columnspan=5, padx=10, pady=10)

        # var_type
        self.var_type_start = 1

        self.var_type_drop_down = tkk.OptionMenu(
            self, self.var_type, "Constants", "Data", "Lookups")

        tkk.Label(self, text="Variable type:").grid(
            row = self.var_type_start, column = 0, padx=10, pady=10)
        self.var_type_drop_down.grid(row=self.var_type_start, column=1)

        # name
        name_start = 2
        tkk.Label(self, text="Variable name:").grid(
            row = name_start, column = 0, padx=10, pady=10)
        self.entry_name = tkk.Entry(self, textvariable=self.var_name)
        self.entry_name.grid(row=name_start, column=1)

        # loading
        self.var_loading_drop_down = tkk.OptionMenu(
            self, self.var_loading, "DIRECT", "XLS")

        tkk.Label(self, text="Loading:").grid(
            row = name_start, column = 2, padx=10, pady=10)

        self.var_loading_drop_down.grid(row=name_start, column=3)

        # subs
        subs_start = 3
        tkk.Label(self, text="Subscript:").grid(
            row = subs_start, column = 0, padx=10, pady=10)

        tkk.Label(self, text="search").grid(
            row = subs_start+2, column = 0)
        self.entry_subs = tkk.Entry(self, textvariable=self.var_sub)
        self.entry_subs.grid(row=subs_start+3, column=0)

        self.right_scrollbar = tkk.Scrollbar(self, orient=tkk.VERTICAL)
        self.lstbox = tkk.Listbox(
            self, yscrollcommand=self.right_scrollbar.set, height=6)
        self.right_scrollbar.grid(
            row=subs_start, column=1, rowspan=4, padx=10)
        self.lstbox.grid(
            row=subs_start, column=1, rowspan=4, padx=10)
        self.right_scrollbar.config(command=self.lstbox.yview)
        self.lstbox.bind('<<ListboxSelect>>', self.add_subs)
        self.populate_list()

        tkk.Label(self, textvariable=self.var_subs_info).grid(
            row=subs_start+1, column=2)
        self.clean_button = tkk.Button(self)
        self.clean_button['text'] = 'Clean'
        self.clean_button['command'] = self.clean_subs
        self.clean_button.grid(row=subs_start+2, column=2)

        # units and descr
        units_start = 7
        tkk.Label(self, text="Variable units:").grid(
            row = units_start, column=0, padx=10, pady=10)
        self.entry_unit = tkk.Entry(self, textvariable=self.var_units)
        self.entry_unit.grid(
            row=units_start, column=1, columnspan=1)

        tkk.Label(self, text="Variable description:").grid(
            row = units_start+1, column=0, padx=10, pady=10)
        self.entry_desc = tkk.Text(self)
        self.entry_desc.grid(
            row=units_start+1, column=1, rowspan=2, columnspan=3)
        self.entry_desc.bind(
            '<KeyRelease>', lambda *args: self.update_description())

        # file and sheet
        units_start = 10
        tkk.Label(self, text="File name:").grid(
            row = units_start, column=0, padx=10, pady=10)
        self.entry_file = tkk.Entry(self, textvariable=self.var_file)
        self.entry_file.grid(
            row=units_start, column=1, columnspan=1)

        tkk.Label(self, text="Sheet name:").grid(
            row = units_start+1, column=0, padx=10, pady=10)
        self.entry_file = tkk.Entry(self, textvariable=self.var_sheet)
        self.entry_file.grid(
            row=units_start+1, column=1, columnspan=1)

        # reference cell
        tkk.Label(self, text="Reference cell:").grid(
            row = units_start, column=2, padx=10, pady=10)
        self.entry_file = tkk.Entry(self, textvariable=self.var_cell)
        self.entry_file.grid(
            row=units_start, column=3, columnspan=1)

        # check boxes
        self.force_start = 12
        self.force_cb = tkk.Checkbutton(
            self, text="Force overwriting cellrange names",
            variable=self.var_force)
        self.force_cb.grid(
            row = self.force_start, column=0, columnspan=2, padx=10, pady=10)

        # next
        next_start = 13
        self.next_button = tkk.Button(self)
        self.next_button['text'] = 'next'
        self.next_button['command'] = self.next
        self.next_button.grid(row=next_start, column=4, padx=10, pady=10)

    def second_window(self):
        tkk.Label(self, textvariable=self.var_info).grid(row=0, column=0, columnspan=5, padx=10, pady=10)

        if self.var_type.get() == "Constants":
            subs_start = 2
        else:
            subs_start = 5
            if self.var_type.get() == "Data":
                x_name = "time"
            else:
                x_name = "x"
            self.series_name = tkk.StringVar(self)
            self.series_along = tkk.StringVar(self)
            self.series_cell = tkk.StringVar(self)
            self.series_len = tkk.IntVar(self)
            tkk.Label(self, text=f"Interpolation dimension ({x_name})").grid(
                row=1, column=0, padx=10, pady=10)

            tkk.Label(self, text="Name:").grid(
                row=2, column = 0, padx=10, pady=10)
            self.entry_dim_name = tkk.Entry(self, textvariable=self.series_name)
            self.entry_dim_name.grid(row=2, column=1)

            tkk.Label(self, text="Along:").grid(
                row=3, column=0, padx=10, pady=10)
            option_menu = tkk.OptionMenu(
                self, self.series_along, "row", "col")

            option_menu.grid(row=3, column=1, padx=10, pady=10)

            tkk.Label(self, text="Reference cell:").grid(
                row=2, column=2, padx=10, pady=10)
            self.entry_dim_cell = tkk.Entry(self, textvariable=self.series_cell)
            self.entry_dim_cell.grid(row=2, column=3)

            tkk.Label(self, text="Length:").grid(
                row=3, column=2, padx=10, pady=10)
            self.entry_dim_cell = tkk.Entry(self, textvariable=self.series_len)
            self.entry_dim_cell.grid(row=3, column=3)

        if self.subs:
            self.create_subs_vars()
            tkk.Label(self, text=f"Subscripts information").grid(
                row=subs_start-1, column=0, padx=10, pady=10)
            tkk.Label(self, text="Along").grid(
                row=subs_start, column=1, padx=10, pady=10)
            tkk.Label(self, text="Sep").grid(
                row=subs_start, column=2, padx=10, pady=10)

            for i, sub in enumerate(self.subs):
                tkk.Label(self, text=f"{sub}: ").grid(
                    row=subs_start+i+2, column=0, padx=10, pady=10)
                option_menu = tkk.OptionMenu(
                    self, self.var_along[sub], "row", "col", "sheet", "file")
                option_menu.grid(row=subs_start+i+2, column=1, padx=10, pady=10)
                entry_sep = tkk.Entry(self, textvariable=self.var_sep[sub])
                entry_sep.grid(row=subs_start+i+2, column=2, padx=10, pady=10)
        else:
            i = -1


        # go
        end_start = subs_start + len(self.subs) + 2
        self.next_button = tkk.Button(self)
        self.next_button['text'] = 'previous'
        self.next_button['command'] = self.previous
        self.next_button.grid(row=end_start, column=3, padx=10, pady=10)

        self.next_button = tkk.Button(self)
        self.next_button['text'] = 'add'
        self.next_button['command'] = self.add
        self.next_button.grid(row=end_start, column=4, padx=10, pady=10)

    def last_window(self):
        tkk.Label(self, text=f"Current elements:").grid(
                row=0, column=0, columnspan=4, padx=10, pady=10)
        for i, (element, val) in enumerate(self.element_dict.items()):
            tkk.Label(self, text=f"{element} ({val['type']})").grid(
                row=i+1, column=0, columnspan=4, padx=10, pady=10)

        # go
        end_start = i + 2
        self.next_button = tkk.Button(self)
        self.next_button['text'] = 'new element'
        self.next_button['command'] = self.new_element
        self.next_button.grid(row=end_start, column=2, padx=10, pady=10)

        self.next_button = tkk.Button(self)
        self.next_button['text'] = 'save'
        self.next_button['command'] = self.save
        self.next_button.grid(row=end_start, column=3, padx=10, pady=10)

        self.next_button = tkk.Button(self)
        self.next_button['text'] = 'execute'
        self.next_button['command'] = self.execute
        self.next_button.grid(row=end_start, column=4, padx=10, pady=10)

    def update_type(self, *args):
        if self.var_type.get() == 'Data':
            self.var_type_data_drop_down = tkk.OptionMenu(
                self, self.var_type_data,
                "none", "interpolate", "look forward", "hold backward", "raw")

            self.keyword_label = tkk.Label(self, text="Keyword:")
            self.keyword_label.grid(
                row = self.var_type_start, column = 2)
            self.var_type_data_drop_down.grid(row=self.var_type_start,
                                              column=3)
        elif self.var_type_data_drop_down:
            self.var_type_data.set("none")
            self.var_type_data_drop_down.destroy()
            self.keyword_label.destroy()

        self.update_info()

    def update_force(self, *args):
        if self.var_force.get():
            self.force_warning = tkk.Label(
                self,
                text="Warning this may remove existing\n "
                     + "cellranges from the file")
            self.force_warning.grid(
                row = self.force_start+1, column = 0, columnspan=2)
        elif self.force_warning:
            self.force_warning.destroy()

    def update_list(self, *args):
        """
        Updates the list of variables in lstbox based on the user input
        :param list_:
        :return:
        """
        search_term = self.var_sub.get()

        self.lstbox.delete(0, tkk.END)

        for item in self.non_selected_subs:
            if search_term.lower() in item.lower():
                self.lstbox.insert(tkk.END, item)

    def populate_list(self):
        self.lstbox.delete(0, tkk.END)
        for item in self.non_selected_subs:
            self.lstbox.insert(tkk.END, item)

    def add_subs(self, sub):
        w = sub.widget
        new_sub = w.get(w.curselection()[0])
        self.subs.append(new_sub)
        self.non_selected_subs.remove(new_sub)
        self.entry_subs.delete(0, 'end')
        self.update_subs_info()

    def clean_subs(self):
        self.non_selected_subs = self.all_subs.copy()
        self.subs = []
        self.update_subs_info()


    def create_subs_vars(self):
        self.var_along = {sub: tkk.StringVar(self) for sub in self.subs}
        self.var_sep = {sub: tkk.StringVar(self) for sub in self.subs}

    def update_subs_info(self):
        """
        Print information about the subscripts
        """
        if self.subs:
            self.var_subs_info.set(f"[{', '.join(self.subs)}]")
        else:
            self.var_subs_info.set("")

        try:
            self.populate_list()
        except tkk.TclError:
            pass
        self.update_info()

    def update_info(self):
        """
        Print information about the variable (pseudo-equation)
        """
        var_info = self.var_name.get()

        if self.subs:
            var_info += f"[{', '.join(self.subs)}]"

        if self.var_type.get() == "Data":
            if self.var_type_data.get() != "none":
                var_info += ":" + self.var_type_data.get().upper() + ":"
            var_info += f":=\t\n\nGET_{self.var_loading.get()}_DATA("
        elif self.var_type.get() == "Lookups":
            var_info += f"=\t\n\nGET_{self.var_loading.get()}_LOOKUPS("
        elif self.var_type.get() == "Constants":
            var_info += f"=\t\n\nGET_{self.var_loading.get()}_CONSTANTS("

        if self.var_type.get():
            if self.var_file.get():
                var_info += f"'{self.var_file.get()}', "
            else:
                var_info += "..., "
            if self.var_sheet.get():
                var_info += f"'{self.var_sheet.get()}', "
            else:
                var_info += "..., "
            if self.var_type.get() == "Constants":
                var_info += "...)"
            else:
                var_info += "..., ...)"

        self.var_info.set(var_info)

    def update_description(self):
        try:
            self.description = self.entry_desc.get("1.0", "end-1c")
        except IndexError:
            self.description = ''

    def next(self):
        self.clean()
        self.second_window()

    def add(self):
        name = self.var_name.get()
        var_type = self.var_type.get()

        self.element_dict[name] = {
            "type": var_type,
            "force": self.var_force.get(),
            "loading": self.var_loading.get(),
            "dims": self.subs,
            "cell": self.var_cell.get(),
            "description": self.description,
            "units": self.var_units.get(),
            "file": self.var_file.get(),
            "sheet": self.var_sheet.get(),
            "dimensions": {
                sub: [self.var_along[sub].get(), None]
                for sub in self.subs
            }
        }

        for sub, (along, _) in self.element_dict[name]["dimensions"].items():
            if along in ['row', 'col']:
                self.element_dict[name]["dimensions"][sub][1] =\
                    int(self.var_sep[sub].get())
            else:
                self.element_dict[name]["dimensions"][sub][1] =\
                    [el.strip() for el in self.var_sep[sub].get().split(',')]

        if var_type == 'Lookups':
            self.element_dict[name]['x'] = {
                "name": self.series_name.get(),
                "cell": self.series_cell.get(),
                "read_along": self.series_along.get(),
                "length": self.series_len.get()
                }
        elif var_type == 'Data':
            self.element_dict[name]['time'] = {
                "name": self.series_name.get(),
                "cell": self.series_cell.get(),
                "read_along": self.series_along.get(),
                "length": self.series_len.get()
                }
            if self.var_type_data.get() != "none":
                self.element_dict[name]['interp'] = self.var_type_data.get()

        self.clean_all()
        self.last_window()

    def new_element(self):
        self.clean()
        self.main_window()

    def save(self):
        outname = asksaveasfilename(
                title="Save configuration",
                filetypes=(('.json files',
                            '*.json *.JSON')))
        with open(outname, 'w') as outfile:
            json.dump(self.element_dict, outfile)

    def previous(self):
        self.clean()
        self.main_window()

    def clean(self):
        for l in self.grid_slaves():
            l.destroy()

    def clean_all(self):
        self.var_name.set('')
        self.var_type.set('')
        self.var_type_data.set('none')
        self.var_file.set('')
        self.var_sheet.set('')
        self.var_cell.set('')
        self.var_force.set(False)
        self.non_selected_subs = self.all_subs.copy()
        self.description = ''
        self.var_units.set('')
        self.subs = []
        self.update_subs_info()

        self.clean()

    def execute(self):
        outname = asksaveasfilename(
                title="Save vensim equations",
                filetypes=(('.txt files', '*.txt'),
                           ('All files', '*'))
                )
        vensim_eqs = execute(self.element_dict)

        with open(outname, 'w') as file:
            file.write(vensim_eqs)

        self.destroy


def start_gui():
    root = tkk.Tk()
    root.title("excels2vensim")
    app = Application(master=root)
    app.mainloop()
