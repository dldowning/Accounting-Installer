import os
import glob
import xlwings as xw

from datetime import datetime

from mapping_utils import get_mapping, get_mapping_dictionary
from summary_utils import get_summary, get_new_summary
from review_utils import get_reviewers_aid_df
from balance_table_utils import get_balance_table
from value_inference_utils import extract_data_for_map_dict
from savers import (
    save_new_summary,
    save_based_on_template,
)


def check_user_inputs(acc_name_start, balance_start, date, till):
    status_line = ""
    cont = False
    if acc_name_start == None:
        status_line += "specify account name column; "
        cont = True
    if balance_start == None:
        status_line += "specify balance column; "
        cont = True
    if date == None:
        status_line += "specify date; "
        cont = True
    if till == None:
        status_line += "specify thru; "
        cont = True

    if cont:
        raise ValueError(status_line)


def get_path_data(path, func):
    if os.path.isdir(path):
        dir_path = path
    else:
        dir_path = os.path.dirname(os.path.realpath(path))

    xlsx_files = glob.glob(os.path.join(dir_path, "*.xlsx"))

    xlsx_files = [
        el
        for el in xlsx_files
        if "dictionary" not in el.lower()
        and "_summary" not in el.lower()
        and "new summary" not in el.lower()
        and "_review" not in el.lower()
        and "~$" not in el.lower()
    ]
    # raise Exception(xlsx_files)
    if func != "map_dict":
        assert "Dictionary.xlsx" in os.listdir(
            dir_path
        ), f"Dictionary.xlsx does not exist in {dir_path}"
        dict_file = os.path.join(dir_path, "Dictionary.xlsx")

    if os.path.isfile(path):
        xlsx_files = [path]

    if func != "map_dict":
        return xlsx_files, dir_path, dict_file

    else:
        return xlsx_files, dir_path, None


def gen_docs(
    xlsx_files,
    dict_file,
    new_summary_template_path,
    reviewers_aid_template,
    summary_template,
    func,
    date,
    till,
    balance_start,
    acc_name_start,
    acc_num_start,
):
    acc_to_class, classes, x12mm_classes = get_mapping(dict_file)
    if func == "new_summary":
        summary_fname = os.path.join(os.path.dirname(xlsx_files[0]), "New summary.xlsx")
        if os.path.exists(summary_fname):
            os.remove(summary_fname)
        summaries, out_paths, sheet_names = [], [], []

    for xlsx_path in xlsx_files:
        comp_name = os.path.basename(xlsx_path)[:-5]
        table = get_balance_table(
            xlsx_path,
            balance_start=balance_start,
            acc_name_start=acc_name_start,
            acc_num_start=acc_num_start,
        )
        table["Class Mapping"] = table["Account Number"].map(acc_to_class)

        if func == "review":
            reviewers_aid = get_reviewers_aid_df(
                table, date, x12mm_classes, classes, comp_name=comp_name
            )
            save_based_on_template(
                reviewers_aid_template, reviewers_aid, xlsx_path[:-5] + "_review.xlsx"
            )

        elif func == "summary":
            summary, _ = get_summary(table, comp_name, classes)
            save_based_on_template(
                summary_template, summary, xlsx_path[:-5] + "_summary.xlsx", header=True
            )

        elif func == "new_summary":
            new_summary = get_new_summary(
                new_summary_template_path,
                dict_file,
                xlsx_path,
                comp_name,
                date,
                till,
                balance_start,
                acc_name_start,
                acc_num_start,
            )
            summaries.append(new_summary)
            sheet_names.append(os.path.basename(xlsx_path[:-5]))

    if func == "new_summary":
        out_path = os.path.dirname(xlsx_files[0])
        save_new_summary(new_summary_template_path, summaries, out_path, sheet_names)


def generate(
    new_summary_template_path,
    reviewers_aid_template,
    summary_template,
    dictionary_template_path,
    func,
):
    wbapp = xw.apps(xw.apps.keys()[0])
    wb = wbapp.books[0]
    sheet = wb.sheets[0]
    path = os.path.dirname(wb.fullname)
    xlsx_files, dir_path, dict_file = get_path_data(path, func)

    cur_stat = "B1"
    try:
        sheet[cur_stat].value = "Waiting..."

        if func == "map_dict":
            dict_full_path = os.path.join(dir_path, "Dictionary.xlsx")
            (
                acc_num_start,
                acc_name_start,
                balance_start,
                date,
                till,
            ) = extract_data_for_map_dict(xlsx_files)
            mapping_dict = get_mapping_dictionary(
                xlsx_files,
                dictionary_template_path,
                path=path,
                date=date,
                till=till,
                balance_start=balance_start,
                acc_name_start=acc_name_start,
                acc_num_start=acc_num_start,
            )
            save_based_on_template(
                dictionary_template_path, mapping_dict, dict_full_path
            )
            wb.close()
            wbapp.quit()

            os.system(f"start excel.exe {dict_full_path}")
        else:
            acc_num_start, acc_name_start, balance_start, date, till = list(
                map(lambda x: sheet[x].value, ["C2", "E2", "G2", "C3", "E3"])
            )
            check_user_inputs(acc_name_start, balance_start, date, till)
            if type(date) == str:
                date = datetime.strptime(date, "%m/%d/%Y")
            if type(till) == str:
                till = datetime.strptime(till, "%m/%d/%Y")
            gen_docs(
                xlsx_files,
                dict_file,
                new_summary_template_path,
                reviewers_aid_template,
                summary_template,
                func,
                date,
                till,
                balance_start,
                acc_name_start,
                acc_num_start,
            )
            if func == "review":
                sheet[cur_stat].value = "Review generated!"
            elif func == "summary":
                sheet[cur_stat].value = "Summary generated!"
            elif func == "new_summary":
                sheet[cur_stat].value = "Big summary generated!"

    except Exception as e:
        print(e)
        sheet[cur_stat].value = str(e)
