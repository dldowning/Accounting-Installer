import xlwings as xw
import pandas as pd
from styleframe import StyleFrame, Styler, utils
import os


def save_reviewers_aid(reviewers_aid, out_path):
    sf = StyleFrame(
        reviewers_aid,
        Styler(
            border_type=None,
            fill_pattern_type=None,
            font=utils.fonts.calibri,
            font_size=11,
        ),
    )
    sf.set_column_width_dict({0: 28, 1: 31, 2: 11, 3: 19, 4: 50, 5: 15})
    ew = StyleFrame.ExcelWriter(out_path)
    sf.to_excel(ew, header=False)
    ew.save()
    ew.close()


def save_summary(summary, out_path):
    sf = StyleFrame(
        summary,
        Styler(
            border_type=None,
            fill_pattern_type=None,
            font=utils.fonts.calibri,
            font_size=11,
        ),
    )
    cols = summary.columns
    sf.set_column_width_dict({cols[0]: 30, cols[1]: 30})
    ew = StyleFrame.ExcelWriter(out_path)
    sf.to_excel(ew)
    ew.save()
    ew.close()


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


"""def save_new_summary(summaries, out_path, sheet_names):
    ew = StyleFrame.ExcelWriter(os.path.join(out_path, "New summary.xlsx"))
    for summary, sheet_name in zip(summaries, sheet_names):
        sf = StyleFrame(
            summary,
            Styler(
                border_type=None,
                fill_pattern_type=None,
                font=utils.fonts.calibri,
                font_size=11,
                horizontal_alignment="left",
                vertical_alignment="center",
            ),
        )
        cols = summary.columns
        sf.set_column_width_dict(
            {
                cols[0]: 30,
                cols[1]: 50,
                cols[2]: 30,
                cols[3]: 30,
                cols[4]: 5,
                cols[5]: 30,
                cols[6]: 5,
                cols[7]: 30,
                cols[8]: 5,
                cols[9]: 45,
            }
        )
        sf.to_excel(ew, header=None, sheet_name=sheet_name[-31:])
    ew.save()
    ew.close()
"""

