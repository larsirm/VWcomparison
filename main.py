# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/



import pandas as pd
import numpy as np


# new_worksheet = pd.DataFrame()
# result_worksheet = pd.DataFrame()
# data_worksheet = pd.read_excel("20221111-VWBank-Fachkonzept-ECBLoanTape - Copy (working version).xlsx", sheet_name="Loan tape – Data field list ")
# table_names = data_worksheet["Table"].unique()

# lst_0_1 = []
# lst_1_1 = []
# lst_2 = []
# for table_name in table_names:
#     datafields = data_worksheet[data_worksheet["Table"] == table_name]["Data field"].unique()
#     for datafield in datafields:
#         work_unit = data_worksheet[(data_worksheet["Table"]==table_name) & (data_worksheet["Data field"]==datafield)]
#         nr_retail = len(work_unit[work_unit['Retail/Non Retail'] == 'Retail'])
#         nr_nonretail = len(work_unit[work_unit['Retail/Non Retail'] == 'Non Retail'])
#         if (nr_retail == 0 and nr_nonretail == 1) or (nr_retail == 1 and nr_nonretail == 0):
#             lst_0_1.append([table_name, datafield, nr_retail, nr_nonretail])
#         if nr_retail == 1 and nr_nonretail == 1:
#             lst_1_1.append([table_name, datafield, nr_retail, nr_nonretail, (work_unit['Field Name'] == work_unit['Field Name'].iloc[0]).all(), (work_unit['Field Definition / Additional Guidance'] == work_unit['Field Definition / Additional Guidance'].iloc[0]).all(), (work_unit['DFs'] == work_unit['DFs'].iloc[0]).all(), (work_unit['Pseudocode'] == work_unit['Pseudocode'].iloc[0]).all()])
#         if nr_retail > 1 or nr_nonretail > 1:
#             lst_2.append([table_name, datafield, nr_retail, nr_nonretail])
#             print(lst_2)




dictionary_worksheet = pd.read_excel("2022-09-19_DG_OMI_FRI_C1_Credit_Risk_Loan_Tape_Data_Dictionary_v1.3.xlsb", sheet_name="Overview")
data_worksheet = pd.read_excel("20221111-VWBank-Fachkonzept-ECBLoanTape - Copy (working version).xlsx", sheet_name="Loan tape – Data field list ")

dictionary_worksheet_selected = dictionary_worksheet[["Table","Data field", "RRE", "Retail SME", "Retail Other", "CRP", "CRE", "LF"]]
data_worksheet_selected = data_worksheet[["Table","Data field", "RRE", "Retail SME", "Retail Other", "CRP", "CRE", "LF"]]


df_all = dictionary_worksheet_selected.merge(data_worksheet_selected.drop_duplicates(), on=["Table","Data field", "RRE", "Retail SME", "Retail Other", "CRP", "CRE", "LF"],
                   how='outer', indicator=True)

df_all[df_all["_merge"]!="both"]
df_all_goup = df_all.groupby( ['Table', 'Data field'] ).size().to_frame(name = 'count').reset_index()
different_records = df_all_goup[df_all_goup['count']>1].drop_duplicates()


result_one = different_records.merge(data_worksheet_selected.drop_duplicates(), on=["Table","Data field"], how='left', indicator=True)
result_one = result_one.assign(_merge='tianyu merge result')
result_one = result_one.drop(columns=['count'])

result_two = different_records.merge(dictionary_worksheet_selected.drop_duplicates(), on=["Table","Data field"], how='left', indicator=True)
result_two = result_two.assign(_merge='Dictionary_v1.3 result')
result_two = result_two.drop(columns=['count'])

concat_result = pd.concat([result_one, result_two])
concat_result = concat_result.sort_values(['Table', 'Data field'],
              ascending = [True, True])
concat_result.to_excel("different_rows.xlsx")
