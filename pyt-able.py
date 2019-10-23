# pyt-able.py
# module for EXCEL exports

from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def create_demo(filename="example.xlsx", title="Example"):
    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = title

    # merge cells for header
    worksheet.merge_cells('A1:F1')
    worksheet["A1"] = title

    # set_columns_width(worksheet, (20, 20, 20, 15))

    workbook.save(filename=filename)


def set_columns_width(worksheet, widths):
    for no, width in enumerate(widths, 1):
        worksheet.column_dimensions[get_column_letter(no)].width = width


create_demo()
