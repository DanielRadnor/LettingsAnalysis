import variables as vb
import pandas as pd
import os

# DATA IMPORT
# -----------

# Finds all Excel files in the specified directory
path = vb.path
files = os.listdir(path)
# Create list of excel files in directory "path"
files_xls = [f for f in files if f[-3:] == 'xls']
files_xls += [f for f in files if f[-4:] == 'xlsx']

df1 = pd.DataFrame()
df2 = pd.DataFrame()

rent_payments_ws = 'Rent Payments'
booking_summary_ws = 'Booking Summary'

# Import excel files into 2 dataframes (one for each excel worksheet)
for f in files_xls:
    data1 = pd.read_excel(path + "\\" + f, sheet_name = rent_payments_ws)
    data2 = pd.read_excel(path + "\\" + f, sheet_name = booking_summary_ws)
    df1 = df1.append(data1)
    df2 = df2.append(data2)

# Merge on "User Id", "Address" and "Room Number" to rule out merging people with different names
combined_df = df1.merge(df2, how='left',left_on=['User Id','Address','Room Number', 'Contract Start', 'Contract End', 'Contract State'],
                        right_on=['User Id','Address','Room Number','Contract Start', 'Contract End', 'Contract State'], suffixes=('', '_y'))

# Removes columns which have been duplicated (i.e. which were in both excel worksheets)
combined_df.drop(combined_df.filter(regex='_y$').columns.tolist(),axis=1, inplace=True)