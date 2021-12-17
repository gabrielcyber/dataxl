from .get import *
from .validation import *


class Design(object):
    def __init__(self, workbook_active):
        self.__ws = workbook_active
        self.__color_names = dict(white='FFFFFF')
        self.__borders = styles.borders
        self.__font = styles.fonts.Font()
        self.__background = styles.fills.PatternFill()

    def text_size(self, cell, size: int):
        setattr(self.__font, 'sz', size)
        setattr(self.__ws[cell], 'font', self.__font)

    def text_font(self, cell, font: str, bold: bool, italic: bool, underline: int):
        __elements__ = ['singleAccounting', 'double', 'single', 'doubleAccounting']
        setattr(self.__font, 'name', font)
        setattr(self.__font, 'bold', bold)
        setattr(self.__font, 'italic', italic)
        try:
            setattr(self.__font, 'underline', __elements__[underline])
        except (IndexError, TypeError):
            pass
        setattr(self.__ws[cell], 'font', self.__font)

    def color(self, cell, font_color, background_color):
        if font_color is not None:
            setattr(self.__font, 'color', font_color)
            setattr(self.__ws[cell], 'font', self.__font)
        if background_color is not None:
            setattr(self.__background, 'start_color', background_color)
            setattr(self.__background, 'fill_type', 'solid')
            setattr(self.__ws[cell], 'fill', self.__background)

    def column_width(self, cell, width: float):
        setattr(self.__ws.column_dimensions[cell[0]], 'width', width)

    def line_height(self, cell, height: float):
        setattr(self.__ws.row_dimensions[int(cell[1:])], 'height', height)

    def alignment(self, cell, position):
        if type(position) is tuple:
            horizontal = position[0]
            vertical = position[1]
        else:
            horizontal = vertical = position
        setattr(self.__ws[cell], 'alignment', styles.alignment.Alignment(horizontal=horizontal, vertical=vertical))

    def border(self, *cells, automatic_finish):
        __elements__ = ['dashDot', 'dashDotDot', 'dashed', 'dotted',
                        'double', 'hair', 'medium', 'mediumDashDot', 'mediumDashDotDot',
                        'mediumDashed', 'slantDashDot', 'thick', 'thin']

        side = self.__borders.Side(style=__elements__[cells[0]])
        if automatic_finish:
            if cells[2] is None:
                cells = list(cells)
                cells[2] = get_column_by_index(self.__ws.min_column - 1)
                cells[2] += str(self.__ws.min_row)
            for row in range(int(cells[2][1:]), self.__ws.max_row + 1):
                for column in range(get_column_index(cells[2][0]), self.__ws.max_column):
                    cell = get_column_by_index(column) + str(row)
                    if self.__ws[cell].value is not None:
                        self.__ws[cell].border = self.__borders.Border(left=side, right=side, top=side, bottom=side)
        else:
            if cells[1] is None:
                for row in range(int(cells[2][1:]), int(cells[3][1:]) + 1):
                    for column in range(get_column_index(cells[2][0]),
                                        get_column_index(cells[3][0]) + 1):
                        self.__ws[get_column_by_index(column) + str(row)].border = \
                            self.__borders.Border(left=side, right=side, top=side, bottom=side)
            else:
                highlighter = [':', ';', ',', '-', '/', '.']
                for divisor in highlighter:
                    find_divisor = cells[1].find(divisor)
                    if find_divisor > 0:
                        for row in range(int(cells[1][:find_divisor][1:]), int(cells[1][find_divisor + 1:][1:]) + 1):
                            for column in range(get_column_index(cells[1][:find_divisor][0]),
                                                get_column_index(cells[1][find_divisor + 1:][0]) + 1):
                                self.__ws[get_column_by_index(column) + str(row)].border = \
                                    self.__borders.Border(left=side, right=side, top=side,
                                                        bottom=side)
