
# importing pandas library for dealing with excel file
from os import write
import random
import pandas as pd
# importing json for handling config.json file
import json

# loading config file
config_file = open('config.json')
config_data = json.load(config_file)

# print(config_data)
file1_config = config_data[0]["file1_configuration"]
file2_config = config_data[1]["file2_configuration"]

file1 = pd.read_excel(
    file1_config["file_path"], dtype=str, header=1 if file1_config["contains_header"] == "y" else 0)
file2 = pd.read_excel(
    file2_config["file_path"], dtype=str, header=1 if file2_config["contains_header"] == "y" else 0)

# print(file1)
# print(file2)
# # print(type(file1))
# print(file1.columns)
# print(file2.columns)
# ========================================color function ====================


def color_fun(row):

    if len(row[len(row)-1]) != 0:
        return ['background-color: #f9b5ac']*(len(row))
    # else:
    #     return ['background-color: #90be6d']*(len(row))


print("=====================result========================")

num_of_cols_file1 = len(file1_config["columns"]) if len(
    file1_config["columns"]) > 0 else len(file1.columns)
num_of_cols_file2 = len(file2_config["columns"]) if len(
    file2_config["columns"]) > 0 else len(file2.columns)
if(num_of_cols_file1 != num_of_cols_file2):
    raise Exception("columns in both file should be equal")
else:
    result_data = []
    file1_col_name = file1_config["columns"] if len(
        file1_config["columns"]) > 0 else list(file1.columns)
    file2_col_name = file2_config["columns"] if len(
        file2_config["columns"]) > 0 else list(file2.columns)
    for row_index in file1.index:
        comment = []
        for col_index in range(num_of_cols_file1):
            file1_cell_data = str(file1[file1_col_name[col_index]][row_index])
            file2_cell_data = str(file2[file2_col_name[col_index]][row_index])
            if(file1_cell_data != file2_cell_data):
                comment.append(str(file2_col_name[col_index]))
        comment = ", ".join(comment)
        if len(comment) > 0:
            comment += " not matching"
        data_row = []
        data_row.extend([str(cd) for cd in list(file1.iloc[row_index])])
        data_row.extend([str(cd) for cd in list(file2.iloc[row_index])])
        data_row.append(comment)
        result_data.append(data_row)
    result_data_columns = list()
    zipped_col = list(zip(file1_col_name, file2_col_name))
    result_data_columns.extend([file1_col if file1_col!=file2_col else f'{file1_col}_1'for file1_col,file2_col in zipped_col])
    result_data_columns.extend([file2_col if file2_col!=file1_col else f'{file2_col}_2' for file1_col, file2_col in zipped_col])
    result_data_columns.append("Comment")
    # convert list to dataframe
    # result_data_df = pd.DataFrame(result_data)
    result_data_df = pd.DataFrame(
        result_data, dtype=str, columns=result_data_columns).astype(str)
    result_data_df_styled = result_data_df.style.apply(color_fun, axis=1)
    # print(result_data)
    with pd.ExcelWriter("result.xlsx") as writer:
        file1.to_excel(writer, sheet_name="File1", index=False)
        file2.to_excel(writer, sheet_name="File2", index=False)
        result_data_df_styled.to_excel(
            writer, sheet_name="Result", index=False)
