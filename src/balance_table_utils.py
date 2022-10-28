import xlwings as xw
import pandas as pd


def get_pease_sheet(wb):
    sheets = wb.sheets
    if len(sheets) > 1:
        pease_index = [
            i
            for i, el in enumerate(list(map(lambda x: str(x).upper(), list(sheets))))
            if "PEASE" in el
        ]
        if pease_index:
            pease_index = pease_index[0]
            sheet = sheets[pease_index]
        else:
            sheet = sheets[0]
    else:
        sheet = sheets[0]

    return sheet


def separate_number_from_name(df_table):
    df_table = df_table["Account Name"].str.split(" ", n=1, expand=True)
    df_table.columns = ["Account Number", "Account Name"]
    return df_table


def get_start_data(sheet, start, name):
    return pd.DataFrame(sheet.range(start).expand("down").value, columns=[name])


def get_df_table(sheet, balance_start, acc_name_start, acc_num_start=None):
    df_table = get_start_data(sheet, acc_name_start, "Account Name")
    if acc_num_start == None:
        df_table = separate_number_from_name(df_table)
    else:
        acc_numbers = get_start_data(sheet, acc_num_start, "Account Number")
        df_table = pd.concat([acc_numbers, df_table], axis=1).dropna()

    balance_values = get_start_data(sheet, balance_start, "Closing Balance").astype(
        float
    )
    df_table = pd.concat([df_table, balance_values], axis=1).dropna()

    return df_table


def get_balance_table(xlsx_path, balance_start, acc_name_start, acc_num_start=None):
    wbapp = xw.App(visible=False)
    wb = wbapp.books.open(xlsx_path)
    sheet = get_pease_sheet(wb)
    df_table = get_df_table(
        sheet,
        balance_start=balance_start,
        acc_name_start=acc_name_start,
        acc_num_start=acc_num_start,
    )
    wb.close()
    wbapp.quit()
    df_table["Account Number"] = df_table["Account Number"].astype(str)

    return df_table
