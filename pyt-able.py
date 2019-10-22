# pyt-able.py
# module for EXCEL exports

from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def create_demo(filename="example.xlsx", title="Example"):
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = title

    sheet["A1"] = "This is EXAMPLE"

    set_columns_width(sheet, (20, 20, 20, 15))

    workbook.save(filename=filename)


def set_columns_width(worksheet, widths):
    for no, width in enumerate(widths, 1):
        worksheet.column_dimensions[get_column_letter(no)].width = width


# create_demo()
