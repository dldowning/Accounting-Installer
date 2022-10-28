import re
import numpy as np
import pandas as pd
import xlwings as xw
import datetime
from balance_table_utils import get_pease_sheet


def column(num, res=""):
    return (
        column((num - 1) // 26, "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[(num - 1) % 26] + res)
        if num > 0
        else res
    )


def get_tb_row_col(sheet):

    max_col = 50
    max_row = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row

    tb_data = pd.DataFrame(sheet.range((1, 1), (max_row, max_col)).value)
    tb_data = tb_data.loc[:, tb_data.isnull().sum() <= len(tb_data) // 2]
    col_index = len(tb_data.columns)
    tb_data["types"] = (
        pd.concat(
            [
                tb_data.loc[:, i].apply(lambda x: str(type(x))[7:-1])
                for i in range(len(tb_data.columns))
            ],
            1,
        )
        .agg(" ".join, axis=1)
        .str.replace("'", "")
    )
    most_freq = tb_data["types"].mode().values[0]
    row_index = (
        np.argwhere(tb_data["types"].str.contains(most_freq).to_numpy()).flatten()[0]
        + 1
    )

    return row_index, col_index


def get_dates(sheet):
    sheet_vals = sheet["A1:D5"].value
    datetime_list = [
        el for el in np.array(sheet_vals).flatten() if type(el) == datetime.datetime
    ]

    if datetime_list:
        datetime_list = list(map(lambda x: x.strftime("%m/%d/%Y"), datetime_list))
        if len(datetime_list) == 1:
            date_from, date_to = [datetime_list[0]], []
        else:
            date_from, date_to = [datetime_list[0]], [datetime_list[1]]
    else:
        header_str = " ".join([el for el in sum(sheet_vals, []) if type(el) == str])
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
                map(
                    str.isdigit,
                    re.sub(
                        "[^A-Za-z0-9]+", "", sheet.range((row_index + 1, 1)).value[-4:]
                    ),
                )
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
