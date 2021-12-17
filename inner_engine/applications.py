from .import sheet, errors
from .get import *
from .validation import *
from dataxl.templates import standard_template
from dataxl.analysis import email


class DataAnalysis(object):
    def save(self, name=None):
        if name is None:
            self.__lib.save(self.file_name)
        else:
            self.__lib.save(name)

    def __init__(self, file_xlsx, template=None, intelligence=None):
        self.file_name = file_xlsx + '.xlsx' if not file_xlsx[-5:] == '.xlsx' else file_xlsx
        try:
            self.__lib = load_workbook(self.file_name)
        except FileNotFoundError:
            self.__lib = Workbook()
        finally:
            self.__ws = self.__lib.active
            self.__design = lambda: sheet.Design(self.__ws)
            self.__template = standard_template(self, template=template, get=True)
            standard_template(self, template, intelligence)

    def create_spreadsheet(self, spreadsheet_name: str):
        self.__lib.create_sheet(spreadsheet_name)
        self.save()

    def rename_spreadsheet(self, spreadsheet_name: str):
        self.__lib.active.title = spreadsheet_name

    def cell(self, cell: str, value: str = None, text_size: int = None, text_font: str = None,
             bold=False, italic=False, underline_index=None, text_color=None, background_color=None,
             column_width=None, line_height=None, alignment=None):
        cell = cell.upper()

        # Design applications
        if self.__template is not None:
            errors.A1_error(cell, self.__ws)  # check if the cell has column A or row one
            if column_width is None:
                column_width = self.__template.DEFAULT_COLUMN_WIDTH
            if line_height is None:
                line_height = self.__template.DEFAULT_LINE_HEIGHT

            design = self.__design()
            design.text_size(cell, text_size)
            design.text_font(cell, text_font, bold, italic, underline_index)
            design.color(cell, text_color, background_color)
            design.column_width(cell, column_width)
            design.line_height(cell, line_height)
            design.alignment(cell, alignment)

        self.__ws[cell] = value
        self.save()

        if DEFAULT_STARTUP_BORDER:
            self.__border(*EDGE_PARAMETERS)

    def insert_unlimited(self, *values: dict, starting_cell: str = '\x42\x33', sort=DEFAULT_ORGANIZE):
        cell = starting_cell.upper()

        # Errors
        if self.__template is not None:
            errors.A1_error(cell, self.__ws)  # check if the cell has column A or row one
        errors.values_package(cell, values)  # check if the values is a package

        values = values[0] if values[0].__class__ is list or \
                              values[0].__class__ is tuple or \
                              values[0].__class__ is dict else values
        if sort:
            for row in range(int(cell[1:]), self.__ws.max_row + 1):
                temp_cell = cell[0] + str(row)
                if self.__ws[temp_cell].value is None:
                    cell = temp_cell
                    break
                else:
                    cell = temp_cell[0] + str(int(temp_cell[1:]) + 1)

        for index in range(len(values)):
            if values.__class__ is not dict:  # if it's a package
                self.__ws[get_column_by_index(get_column_index(cell[0]) + index) + str(int(cell[1:]))] = values[index]
                if self.__template is not None:
                    self.__design().color(get_column_by_index(get_column_index(cell[0]) + index) + str(int(cell[1:])),
                                          self.__template.DEFAULT_TEXT_COLOR,
                                          self.__template.DEFAULT_BACKGROUND_COLOR)
            elif values.__class__ is dict:  # if it's a dictionary
                for reference in values.items():
                    self.__ws[cell[0] + cell[1:]] = reference[0]  # dictionary key
                    self.__design().color(cell[0] + cell[1:], self.__template.DEFAULT_TEXT_COLOR,
                                          self.__template.DEFAULT_BACKGROUND_COLOR)
                    for element in range(len(reference[1])):  # dictionary values as a list
                        self.__ws[get_column_by_index(get_column_index(cell[0]) + (element + 1)) + cell[1:]] = \
                            reference[1][element]
                        self.__design().color(get_column_by_index(get_column_index(cell[0]) + (element + 1)) + cell[1:],
                                              self.__template.DEFAULT_TEXT_COLOR,
                                              self.__template.DEFAULT_BACKGROUND_COLOR)
            self.save()

            if DEFAULT_STARTUP_BORDER:
                self.__border(*EDGE_PARAMETERS)

    def remove_cell(self, cell):
        setattr(self.__ws[cell], 'value', None)
        self.save()

    def privateFunction(self):
        return self.__ws

    def border(self, index: int, select_cells=None, start=None, finish=None, istemplate=False):
        global EDGE_PARAMETERS, DEFAULT_STARTUP_BORDER
        EDGE_PARAMETERS = index, select_cells, start, finish, DEFAULT_AUTOMATIC_FINISH
        if istemplate:
            DEFAULT_STARTUP_BORDER = True
        else:
            self.__border(*EDGE_PARAMETERS)

    def __border(self, *args: tuple):
        args = list(args)
        if args:
            args[-1] = args[-1] if args[1] is None and args[3] is None else False
            for i in [1, 2, 3]:
                try:
                    args[i] = args[i].upper()
                except AttributeError:
                    pass
            else:
                args[2] = self.__balance_cell(args[2], 'A1')
                errors.edge_error(args[2], args[3])  # check if the starting edge is greater than the ending edge
                self.__design().border(args[0], args[1], args[2], args[3], automatic_finish=args[-1])
        self.save()

    @staticmethod
    def __balance_cell(cell, cell_exception, add_row=1):
        if cell is not None:
            if cell[1:] == '\x30':
                cell = cell[0] + str('\x31')
            if cell[0] == cell_exception[0]:
                cell = ascii_uppercase[ascii_uppercase.find(cell[0]) + 1] + cell[1:]
            if cell[1:] == cell_exception[1:]:
                cell = cell[0] + str(int(cell_exception[1:]) + add_row)
            return cell
