import xlwings as xw
import pandas as pd
import os


def save_mapping_dictionary(dict_file, mapping_dictionary, out_path, header=False):
    wbapp = xw.App(visible=False)
    wb = wbapp.books.open(dict_file)
    ws = wb.sheets["Sheet1"]
    ws["A1"].options(
        pd.DataFrame, header=header, index=False, expand="table"
    ).value = mapping_dictionary
    wb.save(out_path)
    wb.close()
    wbapp.quit()


def save_based_on_template(template_path, df, out_path, header=False):
    wbapp = xw.App(visible=False)
    wb = wbapp.books.open(template_path)
    ws = wb.sheets["Sheet1"]
    ws["A1"].options(
        pd.DataFrame, header=header, index=False, expand="table"
    ).value = df
    wb.save(out_path)
    wb.close()
    wbapp.quit()


def save_new_summary(template_path, summaries, out_path, sheet_names):
    out = os.path.join(out_path, "New summary.xlsx")
    wbapp = xw.App(visible=False)
    wb = wbapp.books.open(template_path)

    ws = wb.sheets["Sheet1"]
    for _ in range(len(summaries) - 1):
        ws.api.Copy(Before=ws.api)

    for ws, summary, sheet_name in zip(wb.sheets, summaries, sheet_names):
        ws.name = sheet_name[-31:]
        ws["A1"].options(
            pd.DataFrame, header=False, index=False, expand="table"
        ).value = summary

    wb.save(out)
    wb.close()
    wbapp.quit()
