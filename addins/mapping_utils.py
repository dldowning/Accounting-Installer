import pandas as pd
import numpy as np
from balance_table_utils import get_balance_table


def get_mapping(dict_file):
    df_mapping = pd.read_excel(
        dict_file,
        usecols=["Acc Number", "Acc Name", "Acc Labeled"],
        skiprows=3,
        converters={"Acc Number": str},
    )

    df_mapping = df_mapping.rename(
        columns={
            df_mapping.columns[0]: "Account Number",
            df_mapping.columns[1]: "Name",
            df_mapping.columns[2]: "Class Number ID",
        }
    )

    df_mapping["Class Number ID"] = df_mapping["Class Number ID"].fillna(0).astype(int)
    df_mapping["Account Number"] = df_mapping["Account Number"].astype(str)

    df_classes = pd.read_excel(
        dict_file, usecols=["For Label Name", "Use Label", "x(12/mm)"], skiprows=3
    )
    df_classes = df_classes[~df_classes["For Label Name"].isnull()]
    df_classes["Use Label"] = df_classes["Use Label"].astype(int)
    classes = ["NOT USED"] + list(
        df_classes["For Label Name"].values[df_classes["Use Label"].values - 1]
    )

    x12mm_classes = df_classes[df_classes["x(12/mm)"] == "Y"]["For Label Name"].values
    acc_to_class = dict(df_mapping[["Account Number", "Class Number ID"]].values)

    return acc_to_class, classes, x12mm_classes


def get_mapping_dictionary(
    xlsx_files,
    dictionary_template_path,
    path,
    date,
    till,
    balance_start,
    acc_name_start,
    acc_num_start=None,
):

    dictionary_template = pd.read_excel(dictionary_template_path, skiprows=3)
    head = pd.read_excel(dictionary_template_path, nrows=4, header=None)
    head.loc[1, 2] = np.nan if acc_num_start == None else acc_num_start
    head.loc[1, 4] = acc_name_start
    head.loc[0, 5] = path
    head.loc[2, 2] = date
    head.loc[2, 4] = till
    head.columns = dictionary_template.columns

    head.insert(6, 6, [np.nan, balance_start, np.nan, np.nan])

    account_data = pd.concat(
        [
            get_balance_table(
                xlsx_file,
                balance_start=balance_start,
                acc_name_start=acc_name_start,
                acc_num_start=acc_num_start,
            )
            for xlsx_file in xlsx_files
        ]
    )
    account_data["Account Number"] = account_data["Account Number"].astype(str)

    account_data = (
        account_data.drop(["Closing Balance"], axis=1)
        .drop_duplicates()
        .sort_values(by=["Account Number"])
    )
    account_data = account_data.rename(
        columns={
            account_data.columns[0]: "Acc Number",
            account_data.columns[1]: "Acc Name",
        }
    )

    return pd.concat([head, dictionary_template, account_data], axis=0).reset_index(
        drop=True
    )


if __name__ == "__main__":
    xlsx_folder = r"E:\Projects\accounting\same_format"
    dictionary_template_path = r"Dictionary_template.xlsx"
    balance_start, acc_name_start, acc_num_start = "B5", "A5", None

    get_mapping_dictionary(
        r"E:\Projects\accounting\same_format",
        dictionary_template_path,
        balance_start,
        acc_name_start,
        acc_num_start,
    )

