import sys
import openpyxl as op
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell

# TO DO
# complete the checking one by one
# add the logging all mismatches per row
# add colouring the cells
# rework needed to delete the returning index, could be not necessary

# make sure that OUR table has exactly the same column order as mapping tupple

# looking for the mapped column based on the given header
def search_in_tuple(header_title):
    x = 0
    y = 0
    while (x < len(mapping_tuple)):
        if str(mapping_tuple[x][0]) == str(header_title):
            return str(mapping_tuple[x][1])
            break
        else:
            x = x + 1
        # print(mapping_tuple[x][y-1])

#  returns the
def find_column_number_in_their(column_name):
    list_index = 1
    print("list = " + str(their_column_names) )
    while (list_index <= len(their_column_names)):
        if str(their_column_names[list_index-1]) == str(column_name):
            return list_index-1
            break
        else:
            list_index = list_index + 1

def looking_for_value_in_column(key_value, column_index):
    column_name = their_column_names[column_index]
    print(column_name)
    print(key_value)
    return_tuple = tuple()
    for column_cell in loantape_sh.iter_cols(1, loantape_sh.max_row):
        if column_cell[0].value == column_name:
            for data in column_cell[1:]:
                # print("XXXXXXXXXXXXXX  "+ data.value)
                if data.value == key_value:
                    print("I've found it in column " + column_name + " " + str(data.coordinate))
                    # data.fill = greencolor
                    return_tuple = (True, str(data.row))
                    print(return_tuple)
                    return return_tuple
                    break
                # else:
                #     print("I'm in else")
                #     return False


# stores number of rows where all instr_id taken from our file was not found in second table
not_found_key = 0
# stores number of rows where all of values matches between files
all_correct_values = 0
# stores number of rows where at least 1 value differs
not_correct = 0
# stores the number of values matched within a single row
local_match = 0
# stores the number of values not-matched within a single row
local_fail = 0
# number of searches returning more than one rows with a instr_id
duplicated_key = 0
# number of rows in our file
number_of_rows_ours = 0
# number of rowss in their file
number_of_rows_their = 0

# output file to save the logs
logFile = "output.txt"
f = open(logFile, "w")

# 1. read file our
testing_worksheet = op.load_workbook("Our.xlsx")
testing_sh = testing_worksheet.worksheets[0]
# 2. read file their
loantape_worksheet = op.load_workbook("Their.xlsx")
loantape_sh = loantape_worksheet.worksheets[0]
# 3. define mapping tuple - order should be the same as in our table
mapping_tuple = (("KDW0001", "BIC/KLL1"), ("KDW0002", "BIC/LK2"), ("KDW0003", "BIC/77382"))
# 4. check number of rows
row_number_ours = testing_sh.max_row
row_number_their = loantape_sh.max_row
print("our rows no: " + str(row_number_ours) + " ; their rows no: " + str(row_number_their))

# go through their column list and save it
their_column_names = list()
for column_cell in loantape_sh.iter_cols():
    their_column_names.append(str(column_cell[0].value))

key_value = "A1"
i = 2

# while (i <= row_number_ours):
for row_cell in testing_sh.iter_rows(2, ):
    print("*****************************************")
    key_column_name = testing_sh[key_value].value
    print("#105 I'm looking for: " + testing_sh[key_value].value)
    column_returned = search_in_tuple(testing_sh[key_value].value)

    # I'm looking now in their table column with that name
    index_their_column_names = find_column_number_in_their(column_returned)
    our_key_index = "A"+str(i)
    our_key_value = testing_sh[our_key_index].value

    print("Current value I'll be looking at their table as key: " + our_key_value)
    key_value_tuple = (False, "A1")
    key_value_tuple = looking_for_value_in_column(our_key_value, index_their_column_names)

    #

    if key_value_tuple is not None and key_value_tuple[0] == True:
        print("I've found the key value in their column")
        # I've found the key, now I can go through the rest of columns
        for cell in row_cell[1:]:
            column_index = 2
            print("matched row_our: " + str(i) + " and row_their: " + key_value_tuple[1])
            # make sure that OUR table has exactly the same column order as mapping tupple

            print(mapping_tuple[column_index-1][0])
            # 1. read the column_header of that value from tuple
            # 2. find the mapped column in their table
            # 3. check if that cell.value matches to returned value
            #     3.1 if yes, then local_match =+ 1
            #     3.2 if no, then add to local_fail =+1, add the difference to the logs

            column_index = column_index + 1
            print(column_index)

    else:
        print("I've not found it")
        not_found_key = not_found_key + 1

    print("Not found keys: " + str(not_found_key))


    # check numbers in local_match  and local_fail
    i= i+1



