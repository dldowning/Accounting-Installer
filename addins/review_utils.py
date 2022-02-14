import numpy as np
import pandas as pd


def get_meta_df(comp_name, date):
    year = date.year
    date = date.strftime("%B %d, %Y")

    meta_array = np.zeros((9, 6)).astype(str)

    meta_array[0][0] = comp_name
    meta_array[1][0] = "YETP Summary"
    meta_array[2][0] = date
    meta_array[4][0] = "Reviewer's Aid"
    meta_array[6][3] = "YETP"
    meta_array[8][1] = year
    return pd.DataFrame(meta_array).replace({"0.0": np.nan})


def get_class_info(df_class, month, cls, x12mm_classes):

    df_class_cb_sum = df_class["Closing Balance"].sum()
    df_class["nans_1"] = np.nan
    df_class["nans_2"] = np.nan
    if cls in x12mm_classes:
        df_class["operation"] = "x (12/mm)"
        df_class_cb_sum *= 12 / month
    else:
        df_class["operation"] = "Sum"
    df_class = df_class[
        [
            "nans_1",
            "nans_2",
            "operation",
            "Account Number",
            "Account Name",
            "Closing Balance",
        ]
    ].reset_index(drop=True)
    df_class.columns = np.arange(len(df_class.columns))
    df_class_header = pd.DataFrame(
        [[np.nan, cls, np.nan, df_class_cb_sum, np.nan, np.nan]]
    )

    class_info = pd.concat([df_class_header, df_class])
    return class_info


def get_reviewers_aid_df(df, date, x12mm_classes, classes, comp_name="No Name"):

    meta_df = get_meta_df(comp_name, date)
    month = date.month

    class_infos = []
    for i in range(len(classes)):
        df_class = df[df["Class Mapping"] == i][
            ["Account Number", "Account Name", "Closing Balance"]
        ]
        class_info = get_class_info(df_class, month, classes[i], x12mm_classes)
        class_infos.append(class_info)

    class_infos.append(class_infos.pop(0))
    reviewers_aid = pd.concat([meta_df] + class_infos)
    return reviewers_aid.reset_index(drop=True)
