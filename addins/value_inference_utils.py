import re
import numpy as np
import xlwings as xw
from balance_table_utils import get_pease_sheet


def column(num, res=""):
    return (
        column((num - 1) // 26, "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[(num - 1) % 26] + res)
        if num > 0
        else res
    )


def get_tb_row_col(sheet):
    max_ind = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row
    size_list = []
    for i in range(1, max_ind):
        s = sheet.range((i, 1)).expand("down").size
        size_list.append(s)

    row_index = np.argmax(np.array(size_list)) + 1
    if (
        sheet.range((row_index, 1)).value == None
        or sheet.range((row_index, 2)).value == None
    ):
        row_index += 1

    col_index = sheet.range((row_index, 1)).end("right").column

    return row_index, col_index


def get_dates(sheet):
    header_str = " ".join(
        [el for el in sum(sheet["A1:D5"].value, []) if type(el) == str]
    )
    date_from = re.findall("From ([0-9]{2}\/[0-9]{2}\/[0-9]{4})", header_str)
    date_to = re.findall("To ([0-9]{2}\/[0-9]{2}\/[0-9]{4})", header_str)

    return date_from, date_to


def extract_data_for_map_dict(xlsx_files):
    date, till = "", ""
    for i, xlsx_file in enumerate(xlsx_files):
        wbapp = xw.App(visible=False)
        wb = wbapp.books.open(xlsx_file)
        sheet = get_pease_sheet(wb)
        if i == 0:
            row_index, col_index = get_tb_row_col(sheet)
            separate_num_from_name = not all(
                map(str.isdigit, sheet.range((row_index + 1, 1)).value[-4:])
            )
            if separate_num_from_name:
                acc_num_start, acc_name_start = None, "A" + str(row_index)
            else:
                acc_num_start, acc_name_start = (
                    "A" + str(row_index),
                    "B" + str(row_index),
                )

            balance_start = column(col_index) + str(row_index)

        if not (date and till):
            date_from, date_to = get_dates(sheet)
            if date_from:
                date = date_from[0]

            if date_to:
                till = date_to[0]

        else:
            break

        wb.close()
        wbapp.quit()

    if not date:
        if till:
            date = "01/01/" + till[-4:]
        else:
            date = ""

    return acc_num_start, acc_name_start, balance_start, date, till


if __name__ == "__main__":
    print(extract_data_for_map_dict(r"E:\Projects\accounting\test"))

