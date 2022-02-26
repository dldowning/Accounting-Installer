from http import client
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

        # client_data_upper_left_corner_row, 120 is the length of the footer
        ul_row = len(summary) - 120

        col_name_dict = {
            "AA": "Amortization",
            "AB": "Depreciation",
            "AC": "Accrual to Cash",
            "AD": "Nondeductible",
            "AE": "Book Income (Loss)",
            "K": "Ownership Percentages",
            "L": "Beginning Tax Capital",
            "M": "Contribution",
            "N": "Distribution",
            "O": "Tax Capital Subtotal",
            "P": "Taxable Income (Loss)",
            "Q": "Tier 1",
            "R": "Tier 2",
            "S": "Special Allocated Taxable Income (Loss)",
            "T": "Ending Tax Basis",
            "U": "Qualified Nonrecourse Debt",
            "V": "Recourse Debt",
            "W": "CY Allocation Percentages",
            "X": "Ordinary Income (Loss)",
            "Y": "Interest Income",
            "Z": "Charitable Contributions",
        }

        for k, v in col_name_dict.items():
            ws[k + str(ul_row)].value = v

        sum_upper_cells_let = ["K", "L", "M", "N", "O", "Q", "S", "T", "V", "W"]
        sum_upper_cells = {
            let + str(ul_row + 3): f"=SUM({let}{ul_row+1}:{let}{ul_row+2})"
            for let in sum_upper_cells_let
        }
        for k, v in sum_upper_cells.items():
            ws[k].value = v

        ws["R" + str(ul_row + 3)].value = f"=P{ul_row+3}-Q{ul_row+3}"
        ws["X" + str(ul_row + 3)].value = f"=P{ul_row+3}-Y{ul_row+3}-Z{ul_row+3}"

        ws["O" + str(ul_row + 1)].value = f"=L{ul_row+1}+M{ul_row+1}+N{ul_row+1}"
        ws["O" + str(ul_row + 2)].value = f"=L{ul_row+2}+M{ul_row+2}+N{ul_row+2}"

        ws["P" + str(ul_row + 1)].value = f"=$P${ul_row+3}*K{ul_row+1}"
        ws["P" + str(ul_row + 2)].value = f"=$P${ul_row+3}*K{ul_row+2}"

        ws["Q" + str(ul_row + 1)].value = f"=P{ul_row+1}"
        ws["Q" + str(ul_row + 2)].value = f"=P{ul_row+2}"

        ws["T" + str(ul_row + 1)].value = f"=O{ul_row+1}+S{ul_row+1}+AD{ul_row+1}"
        ws["T" + str(ul_row + 2)].value = f"=O{ul_row+2}+S{ul_row+2}+AD{ul_row+2}"

        ws["U" + str(ul_row + 1)].value = f"=P{ul_row+1}+T{ul_row+1}+AE{ul_row+1}"
        ws["U" + str(ul_row + 2)].value = f"=P{ul_row+2}+T{ul_row+2}+AE{ul_row+2}"

        ws["W" + str(ul_row + 1)].value = f"=S{ul_row+1}/$S${ul_row+3}"
        ws["W" + str(ul_row + 2)].value = f"=S{ul_row+2}/$S${ul_row+3}"

        mul_list = ["X", "Y", "Z", "AA", "AB", "AC", "AD", "AE"]
        for char in mul_list:
            ws[char + str(ul_row + 1)].value = f"=${char}${ul_row+3}*W{ul_row+1}"
            ws[char + str(ul_row + 2)].value = f"=${char}${ul_row+3}*W{ul_row+2}"

    wb.save(out)
    wb.close()
    wbapp.quit()


"""


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

def save_new_summary(summaries, out_path, sheet_names):
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

