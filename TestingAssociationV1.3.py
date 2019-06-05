import pandas as pd
import xlrd

from mlxtend.preprocessing import OnehotTransactions
from mlxtend.frequent_patterns import apriori

# ----------------------------------------------------------------------------------------------------------------------
# Reading in the Spreadsheet
# ----------------------------------------------------------------------------------------------------------------------
# path, name, and extension for the file being read in
# Note: Set path to nothing if the file is contained within the same folder as this script
path = ""
name_of_spreadsheet = "Breast Cancer"
extension = (".xlsx")

# location / name of spreadsheet based on the above criteria
loc = (path + name_of_spreadsheet + extension)

# Open workbook at the location we set and set which sheet
# Note: Leaving this value as 0 sets it to the first sheet in the Excel file
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# Which column do you want the program to read in? Note: The first column is column 0, second is 1, etc.
spreadsheet_column = 44

# Creating list that will contain all the cells
all_cells = ""

counter = 0

all_cells_list = []
# Loop that sets cell_list to the value of each sell appended onto each other
for cell in sheet.col(spreadsheet_column):
    # Sets cell_value to the value found in the cell
    cell_value = cell.value
    # Removes unneeded characters to reduce clutter
    cell_value = cell_value.replace('"',"")
    cell_value = cell_value.replace(r'[', '')
    cell_value = cell_value.replace(r']', '')
    # Splits each value apart based on spaces and appends them to the cell list
    cell_value_parsed = cell_value.split(',')
    all_cells_list.append(cell_value_parsed)
print(all_cells_list)

# Creation of the data frame based on the cell_list
oht = OnehotTransactions()
oht_ary = oht.fit(all_cells_list).transform(all_cells_list)
df = pd.DataFrame(oht_ary, columns=oht.columns_)
print(df)

# ----------------------------------------------------------------------------------------------------------------------
# Configuration for apriori algorithm
# ----------------------------------------------------------------------------------------------------------------------
# minimum value of the coefficient
min_co = 0.085
# True or False whether or not to include the column names in the output
use_colnames_bool = True
# max number of associations
max_len_value = None
frequent_itemsets = apriori(df, min_support=min_co, use_colnames=use_colnames_bool, max_len=max_len_value)


frequent_itemsets.to_csv(name_of_spreadsheet + ' Data Association.csv')
print(frequent_itemsets)
print("done")