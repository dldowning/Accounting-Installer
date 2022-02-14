import string
import pandas as pd
import numpy as np

from balance_table_utils import get_balance_table
from mapping_utils import get_mapping


def get_summary(df, name, classes):
    cls_summary_vals, not_used_unidentified_dict = [], {}
    for i in range(len(classes)):
        if classes[i] == "NOT USED":
            not_used_unidentified_dict[classes[i]] = dict(
                zip(
                    list(df[df["Class Mapping"] == i]["Account Number"].values),
                    list(df[df["Class Mapping"] == i]["Account Name"].values),
                )
            )

        cls_summary_vals.append(df[df["Class Mapping"] == i]["Closing Balance"].sum())

    comp_df = pd.DataFrame.from_dict({name: cls_summary_vals})
    class_df = pd.DataFrame.from_dict({"CLASS": classes})
    summary = pd.concat([class_df] + [comp_df], 1).iloc[1:]

    return summary, not_used_unidentified_dict


def get_new_summary_meta(template_path, header_values, since, till, col_names):

    df_footer = pd.read_excel(template_path, skiprows=50, names=col_names)
    df_header = pd.read_excel(template_path, nrows=10, names=col_names)

    df_header[["AI", "AJ"]] = 0
    df_header.loc[6, "A"] = f'Book Income (Loss) - Accrual {since.strftime("%m/%d/%Y")}'
    df_header.loc[
        8, "A"
    ] = f'Book Income (Loss) thru {till.strftime("%m/%d/%Y")} - Accrual'
    for i in range(0, len(header_values)):
        df_header.loc[i, "A"] = header_values[i]

    return df_header, df_footer


def create_summary_block(summary, year, with_balance, col_names):

    summary_info = np.empty((summary.shape[0] + 2, 36))
    summary_info[:] = np.nan
    summary_info = pd.DataFrame(summary_info, columns=col_names)

    summary_info["B"] = summary["CLASS"]
    summary_info.loc[0, "B"] = year
    if with_balance:
        summary_info["D"] = summary[summary.columns[1]]

    return summary_info


def create_summary_body(summaries, years, with_balance_vals, col_names):
    summary_blocks = []
    for summary, year, with_balance in zip(summaries, years, with_balance_vals):
        sb = create_summary_block(summary, year, with_balance, col_names)
        summary_blocks.append(sb)

    return pd.concat(summary_blocks)


def get_new_summary(
    new_summary_template_path,
    dict_file,
    xlsx_path,
    comp_name,
    date,
    till,
    balance_start,
    acc_name_start,
    acc_num_start,
):
    col_names = list(string.ascii_uppercase) + list(
        map(lambda x: "A" + x, list(string.ascii_uppercase)[0:10])
    )

    acc_to_class, classes, x12mm_classes = get_mapping(dict_file)
    table = get_balance_table(
        xlsx_path,
        balance_start=balance_start,
        acc_name_start=acc_name_start,
        acc_num_start=acc_num_start,
    )
    table["Class Mapping"] = table["Account Number"].map(acc_to_class)
    summary, _ = get_summary(table, comp_name, classes)
    header_values = [
        comp_name,
        f"{date.year} YETP Projection",
        "Cash Method",
        f"Tax Year {date.year}",
    ]

    df_header, df_footer = get_new_summary_meta(
        new_summary_template_path, header_values, date, till, col_names
    )
    df_body = create_summary_body(
        [summary] * 2, [date.year - 1, date.year], [False, True], col_names
    )
    return pd.concat([df_header, df_body, df_footer]).reset_index(drop=True)
