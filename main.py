import pandas as pd
import numpy as np
import os
from datetime import date, datetime, timedelta
import time
import re

# Start time to calculate running time
start_time = time.perf_counter()

# INPUT VARIABLES
# ---------------

# Minimum number of weeks considered as Long Term
lt_weeks = 43.5

# Max days a contract can be Pending Signatures before being removed
expire_pending = 14

# Switch to run excel export or not
run_export = True

# Dates for 20/21 Academic year
start_date_2020_21 = "2020-09-05"
end_date_2020_21 = '2021-09-03'

# FUNCTIONS
# ---------

# Convert Address to official property name and add as new column
def property_name(address):
    if "Renny" in address:
        return "DRM-Ernest Place"
    elif "Howard Street" in address:
        return "NCL-Mansion Tyne"
    elif "BS1 4TR" in address:
        return "BRS-Centre Gate"
    elif "BS1 4ST" in address:
        return "BRS-The Colston"
    elif "Cromwell Range" in address:
        return "MCR-Mansion Point"
    elif "Boston Street" in address:
        return "NOT-Mansion Place"
    elif "Union Street" in address:
        return "SHF-Redvers Tower"
    else:
        return ""

# Total capacity at each property
capacity_dict = {"LDS-Leodis Residence": 719,
            "DRM-Ernest Place": 362,
            "MCR-Mansion Point": 206,
            "NCL-Mansion Tyne": 416,
            "BRS-Centre Gate": 85,
            "BRS-The Colston": 77,
            "NOT-Mansion Place": 139,
            "SHF-Redvers Tower": 227}

# Returns the 'Capacity" field in the first row of the dataframe
def capacity(dataframe):
    cp = dataframe['Capacity'].loc[0]
    return cp

# Returns the 'Capacity" in the dataframe for a specified property name
def capacity2(dataframe, property_name):
    cp2 = dataframe[dataframe['Property Name'] == property_name].max()["Capacity"]
    return cp2

# Dict for the total potential rent at each property
total_potential_rent = {"LDS-Leodis Residence": 3422562.23,
             "DRM-Ernest Place": 2846820,
             "MCR-Mansion Point": 1727370,
             "NCL-Mansion Tyne": 2268684,
             "BRS-Centre Gate": 779586,
             "BRS-The Colston": 711552,
             "NOT-Mansion Place": 1036422,
             "SHF-Redvers Tower": 1994049}

# Returns the total gross rent achieved for a property
def gross_rent(dataframe, property_name):
    gr = dataframe[dataframe['Property Name'] == property_name].sum()["Contract Total Rent"]
    return gr

# Returns the total confirmed gross rent for a property (i.e. contract status = signed)
def gross_rent_confirmed(dataframe, property_name):
    gr = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract State'] == 'Signed')].sum()["Contract Total Rent"]
    return gr

# Returns the total Long Term rent achieved for a property
def total_lt_rent(dataframe, property_name):
    tltr = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] >= lt_weeks)].sum()["Contract Total Rent"]
    return tltr

# Returns the total Short Term rent achieved  for a property
def total_st_rent(dataframe, property_name):
    tstr = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < lt_weeks)].sum()["Contract Total Rent"]
    return tstr

# Returns the total bookings achieved for a property by counting the number of times there is a Contract Total Rent
def total_bookings(dataframe, property_name):
    tb = dataframe[(dataframe['Property Name'] == property_name)
                   & (dataframe['Contract Total Rent'] > 0)].nunique()["Contract ID"]
    return tb

# Total achieved Long Term bookings for a property (counts twin occupancy as 1 bed)
def total_lt_bookings(dataframe, property_name):
    tltb = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] >= lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].nunique()["Contract ID"]
    return tltb

# Total achieved Short Term bookings for a property (counts twin occupancy as 1 bed)
def total_st_bookings(dataframe, property_name):
    tstb = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].nunique()["Contract ID"]
    return tstb

# Total achieved Short Term bookings for a property converted to long term equivalents
def total_st_bookings_lteqv(dataframe, property_name):
    tstblteqv = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].sum()["Contract Weeks"] / lt_weeks
    return tstblteqv

# Total achieved Short Term bookings for a property in terms of total weeks booked
def total_st_bookings_weeks(dataframe, property_name):
    tstbwk = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].sum()["Contract Weeks"]
    return tstbwk

# Check whether a booking is in 2020/21 based on 50% or more of the booking days after 5th September
def booking_in_2020_21(start_date,end_date):
    start_date_2020_21 = '2020-09-05'
    start_date_2020_21 = pd.to_datetime(start_date_2020_21)
    if start_date > start_date_2020_21:
        return True
    elif (end_date - start_date_2020_21) >= (start_date_2020_21 - start_date):
        return True
    else:
        return False

british_country_list = ["English",
                        "British",
                        "Welsh",
                        "Northern Irish",
                        "Scottish"]

eu_country_list = ['Austrian',
                   'Belgian',
                   'Bulgarian',
                   'Croatian',
                   'Cypriot',
                   'Czech',
                   'Danish',
                   'Estonian',
                   'Finnish',
                   'French',
                   'German',
                   'Greek',
                   'Hungarian',
                   'Irish',
                   'Italian',
                   'Latvian',
                   'Lithuanian',
                   'Luxembourgan',
                   'Luxembourger',
                   'Maltese',
                   'Dutch',
                   'Polish',
                   'Portugese',
                   'Romanian',
                   'Slovakian',
                   'Slovenian',
                   'Spainish',
                   'Swedish']

# Check whether a student is Domestic or International
def domestic_or_intl (country):
    if any(x in country for x in british_country_list):
        return "Domestic"
    else:
        return "International"

# Check whether a student is Domestic or EU or International
def eu_or_intl (country):
    if domestic_or_intl(country) == "Domestic":
        return "Domestic"
    elif "Irish" in country:
        if "Northern" in country:
            return "Domestic"
        else:
            return "EU"
    else:
        if any(x in country for x in eu_country_list):
            return "EU"
        else:
            return "International"

# Prints dictionaries in a pretty manner
def pretty(d, indent=0):
   for key, value in d.items():
      if isinstance(value, dict):
          print('\t' * indent + str(key))
          pretty(value, indent+1)
      else:
          print('\t' * (indent) + str(key) + ":  " + str(value))

# Prints dictionaries in a pretty manner (only indents once)
def pretty_one_step(d, indent=0):
   for key, value in d.items():
      print('\t' * indent + str(key))
      print('\t' * (indent+1) + str(value))

def extract_date_from_str(text):
    match = re.search(r'\d{2}/\d{2}/\d{4}', text)
    date = datetime.strptime(match.group(), '%d/%m/%Y').date()
    return date


# ---------------
# ---------------
# -- MAIN CODE --
# ---------------
# ---------------

# DATA IMPORT
# -----------

# Finds all Excel files in the specified directory
path = "C:\\Users\\daniel.radnor\\PycharmProjects\\Lettings\\Lettings Report Automation\\Raw Files"
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

# Add Property Name to use as part of merge as some properties have 2 spellings
df1['Property Name'] =df1.apply(lambda x: property_name(x['Address']), axis=1)
df2['Property Name'] =df2.apply(lambda x: property_name(x['Address']), axis=1)

# Merge on "User Id", "Address" and "Room Number" to rule out merging people with different names
combined_df = df1.merge(df2, how='left',left_on=['User Id','Property Name','Room Number', 'Contract Start', 'Contract End', 'Contract State'],
                        right_on=['User Id','Property Name','Room Number','Contract Start', 'Contract End', 'Contract State'], suffixes=('', '_y'))

# Removes columns which have been duplicated (i.e. which were in both excel worksheets)
combined_df.drop(combined_df.filter(regex='_y$').columns.tolist(),axis=1, inplace=True)

# Bring Property Name to the first column position
prop_name = combined_df['Property Name']
combined_df.drop(labels=['Property Name'], axis=1,inplace = True)
combined_df.insert(0, 'Property Name', prop_name)

# Remove all records where the rent is 0
combined_df = combined_df[(combined_df['Contract Total Rent'] > 0)]

# ADD EXTRA COLUMNS TO THE DATA
# -----------------------------

# Converts all room numbers to strings
combined_df['Room Number'] = combined_df['Room Number'].apply(str)

# Add property bed capacity for each row
combined_df['Capacity'] = combined_df['Property Name'].map(capacity_dict)

# Add property total potential rent for each row
combined_df['Total Property Potential Rent'] = combined_df['Property Name'].map(total_potential_rent)

# Add Daily Rent rate
combined_df['Contract Daily Rent'] = combined_df.apply(lambda x: (x['Contract Total Rent'] / (x['Contract Days'])), axis =1)

# Add Cluster or Studio column
combined_df['Cluster or Studio'] = combined_df.apply(lambda x: ("Studio" if 'Studio' in x['Room Type'] or 'Twodio' in x['Room Type'] or 'Threedio' in x['Room Type'] or 'Apartment' in x['Room Type'] else "Cluster"), axis = 1)

# Add ST/LT column
combined_df['LT or ST'] = combined_df.apply(lambda x: ("LT" if x['Contract Weeks'] >= lt_weeks else "ST"), axis = 1)

# Change "-" to "No" for Guarantor Signed
combined_df['Guarantor Signed Yes or No'] = combined_df.apply(lambda x: ("No Guarantor" if x['Guarantor Signed'] == "-" else "Guarantor"), axis = 1)

# Fill in Blanks for University
combined_df['University'].fillna(value="Unknown", inplace=True)

# Fill in Blanks for Year of Study
combined_df['Year of Study'].fillna(value="Unknown", inplace=True)

# Fill in Blanks for Bookings Source
combined_df['Booking Source'].fillna(value="Unknown", inplace=True)

# Add Column for Domestic / International
combined_df['Nationality'].fillna(value="Unknown", inplace=True)
combined_df['Domestic or International'] = combined_df.apply(lambda x: domestic_or_intl(x['Nationality']), axis = 1)
combined_df['Domestic EU or International'] = combined_df.apply(lambda x: eu_or_intl(x['Nationality']), axis = 1)

# Add Column for In Arrears
combined_df['In Arrears'] = combined_df.apply(lambda x: (True if x['Total Rent Arrears'] > 0 else False), axis = 1)

# Add Column for Never Paid
combined_df['Never Paid'] = combined_df.apply(lambda x: (True if x['Total Rent Paid'] == 0 else False), axis = 1)

# Add Column for whether contract is live (i.e. today is between the start and end dates
combined_df['Contract Live'] = combined_df.apply(lambda x: (True if (date.today() > x['Contract Start'] and date.today() < x['Contract End']) else False ), axis = 1)
combined_df['Contract Live Count'] = combined_df.apply(lambda x: (1 if x['Contract Live'] == True else 0), axis = 1)

# Days Accrued since start of contract to today
combined_df['Rent Liable Days Accrued'] = combined_df.apply(lambda x: ((datetime.now() - x['Contract Start']) if x['Contract Live'] == True else (x['Contract End'] - x['Contract Start'])), axis = 1)
combined_df['Rent Liable Days Accrued'] = combined_df['Rent Liable Days Accrued'].dt.days
combined_df['Accrued Rent Total'] = combined_df.apply(lambda x: (x['Rent Liable Days Accrued'] * x['Contract Daily Rent']), axis = 1)

# Add Column for High Risk of Defaulting based on: Not checked in, not paid, and start date expired
combined_df['Checked In'] = combined_df.apply(lambda x: (False if (pd.isnull(x['Check In'])) else True), axis = 1)
combined_df['Checked Out'] = combined_df.apply(lambda x: (False if (pd.isnull(x['Check Out'])) else True), axis = 1)
combined_df['In Room'] = combined_df.apply(lambda x: (True if x['Checked In'] == True and x['Checked Out'] == False else False), axis = 1)
combined_df['In Room Count'] = combined_df.apply(lambda x: (1 if x['In Room'] == True else 0), axis = 1)
combined_df['High Risk of Default'] = combined_df.apply(lambda x: (True if (x['Checked In'] == False and x['Never Paid'] == True and x['Total Rent Arrears'] > 0 and x['Contract Start'] < date.today()) else False), axis = 1)

# Label Pending Signatures contracts which are old and expire
combined_df['Unsigned Contract Days'] = combined_df.apply(lambda x: ((datetime.now() - x['Created']) if x['Contract State'] == "Pending Signatures" else 0), axis = 1)
combined_df['Pending Expired'] = combined_df.apply(lambda x: (True if (x['Contract State'] == "Pending Signatures" and ((datetime.now() - x['Created']) > timedelta(days=expire_pending))) else False), axis = 1)

# Store a copy of the dataframe before cleaning (e.g. removing duplicates etc).
original_combined_df = combined_df.copy()

# Remove all records where the Contract State is Cancelled/Expired/Draft
cancelled_df = combined_df[combined_df['Contract State'].str.contains("Cancelled")]
combined_df = combined_df[~combined_df['Contract State'].str.contains("Cancelled")]
expired_df = combined_df[combined_df['Contract State'].str.contains("Expired")]
combined_df = combined_df[~combined_df['Contract State'].str.contains("Expired")]
draft_df = combined_df[combined_df['Contract State'].str.contains("Draft")]
combined_df = combined_df[~combined_df['Contract State'].str.contains("Draft")]

# Extract Cancellation Date and create new column
cancelled_df['Cancellation Date'] = cancelled_df.apply(lambda x: extract_date_from_str(x['Contract State']), axis = 1)

# This removes Pending Signature which are more than "expire_pending" days old
expired_pending_df = combined_df[((combined_df['Contract State'] == "Pending Signatures") & ((datetime.now() - combined_df['Created']) > timedelta(days=expire_pending)))]
combined_df.drop(expired_pending_df.index.tolist(), axis = 0, inplace = True)

# Add column to mark which contracts are 2020_21 and then remove from the main DataFrame
combined_df['2020_21 Contracts'] = combined_df.apply(lambda x: booking_in_2020_21(x['Contract Start'], x['Contract End']), axis=1)
not_2020_21_df = combined_df[combined_df['2020_21 Contracts'] == False] # Create dataframe for  rows not for 2020_21 to new dataframe (these will be removed)
list_to_delete_not_2020_21 = combined_df.index[combined_df['2020_21 Contracts'] == False].tolist()
combined_df.drop(list_to_delete_not_2020_21, axis=0, inplace=True)

# TWIN OCCUPANCY
# --------------
# Duplicates by contract ID remain as the Rent fields are split across the tenants which share a bed.  Just need to identify
#twin_occupancy_df = combined_df[combined_df.duplicated(['Contract ID'], keep=False) &
#                           combined_df.duplicated(['Property Name', 'Room Number'], keep=False)]

twin_occupancy_df = combined_df[combined_df.duplicated(['Contract ID','Property Name', 'Room Number', 'Contract Start', 'Contract End'], keep=False)]

# Create list of rows which are twin occupancy and mark them in new Twin Occupancy column
twin_occupancy_list = twin_occupancy_df.index.tolist()
combined_df['Twin Occupancy'] = combined_df.apply(lambda x: (True if x.name in twin_occupancy_list else False), axis = 1)

# LT EQUIVALENT
# -------------

# If LT then 1 else fraction of 43.5 making sure twin occupancy are halved
combined_df['LT Equivalent'] = combined_df.apply(lambda x: (1 if x['LT or ST'] == 'LT' else x['Contract Weeks'] / lt_weeks), axis = 1)
combined_df['LT Equivalent'] = combined_df.apply(lambda x: ((x['LT Equivalent'] / 2 ) if x['Twin Occupancy'] == True else x['LT Equivalent']), axis = 1)

# DUPLICATES BY ROOM
# ------------------
# All rows where there are duplicates by Room Number (i.e. 2 contracts for the same room)
# =========================================================================================================================
# If there is the odd case of a Twin Occupancy which is a duplicate with a normal occupancy then this will not be reported
# =========================================================================================================================
dups_by_room_df = combined_df[combined_df['Twin Occupancy'] == False]
dups_by_room_df = dups_by_room_df[dups_by_room_df.duplicated(['Property Name', 'Room Number'], keep=False)]

# This section checks the date compared the row above to see whether the contracts for the same room overlap
dups_by_room_df = dups_by_room_df.sort_values(['Property Name', 'Room Number', 'Contract Start', 'Contract End'],
                                              ascending=[True, True, True, True])

dups_by_room_df['Prev Property Name'] = dups_by_room_df['Property Name'].shift(1)
dups_by_room_df['Prev Room Number'] = dups_by_room_df['Room Number'].shift(1)
dups_by_room_df['Prev Contract Start'] = dups_by_room_df['Contract Start'].shift(1)
dups_by_room_df['Prev Contract End'] = dups_by_room_df['Contract End'].shift(1)

# Save original index as we will reset the index
dups_by_room_df['Original Index'] = dups_by_room_df.index

dups_by_room_df = dups_by_room_df.reset_index()

dups_by_room_df['Room Duplicate'] = False
room_duplicate_id = 10000
dups_by_room_df['Room Duplicate ID'] = 0
dups_by_room_df['Date Overlap'] = False

# This section loops through all of the duplicates by room and then assigns a duplicate ID so that they are linked
for index,row in dups_by_room_df.iterrows():
    #if dups_by_room_df.loc[index, 'Contract Start'] != np.datetime64('NaT'):
    if row['Property Name'] == row['Prev Property Name'] and row['Room Number'] == row['Prev Room Number'] and\
            row['Contract Start'] < row['Prev Contract End']:
        dups_by_room_df.loc[index - 1, 'Room Duplicate'] = True
        dups_by_room_df.loc[index, 'Room Duplicate'] = True
        dups_by_room_df.loc[index - 1, 'Date Overlap'] = True
        dups_by_room_df.loc[index, 'Date Overlap'] = True

        if dups_by_room_df.loc[index - 1, 'Room Duplicate ID'] == 0:
            room_duplicate_id += 1
            dups_by_room_df.loc[index - 1, 'Room Duplicate ID'] = room_duplicate_id
            dups_by_room_df.loc[index, 'Room Duplicate ID'] = room_duplicate_id
        else:
            dups_by_room_df.loc[index, 'Room Duplicate ID'] = dups_by_room_df.loc[index - 1, 'Room Duplicate ID']

# Removed all valid duplicates (dates not overlapping etc) leaving only ones with overlapping dates to be cleaned
dups_by_room_df = dups_by_room_df[dups_by_room_df['Room Duplicate'] == True]

# Sort duplicates by Room Duplicate ID, Created and Contract State as this determines what we select to remove
dups_by_room_df = dups_by_room_df.sort_values(['Room Duplicate ID', 'Created', 'Contract State'],
                                              ascending=[True, True, True])

# Save original index as we will reset the index
#dups_by_room_df['Original Index'] = dups_by_room_df.index

dups_by_room_df = dups_by_room_df.reset_index()

previous_room_id = 0
list_to_delete_room = []
dict_to_keep_room = {}

for index,row in dups_by_room_df.iterrows():
    if row['Room Duplicate ID'] != previous_room_id:
        dups_by_room_df_temp = dups_by_room_df[dups_by_room_df['Room Duplicate ID'] == row['Room Duplicate ID']]
        previous_room_id = row['Room Duplicate ID']
        #print(row['Room Duplicate ID'], "\n", "------------- \n", dups_by_room_df_temp)

        # Checks whether any of the duplicates for the room have Contract State = "Signed"
        contains_signed = False
        for i, r in dups_by_room_df_temp.iterrows():
            if r['Contract State'] == "Signed":
                contains_signed = True

        # If any of the duplicates for the room have Contract State = "Signed", drop all others (add to list)
        if contains_signed == True:
            for i, r in dups_by_room_df_temp.iterrows():
                if r['Contract State'] != "Signed":
                    list_to_delete_room.append(r['Original Index'])
                    #dups_by_room_df_temp2.drop(i, axis = 0, inplace = True)

        # Drops all rows other than the oldest room based on Created date (add to llist)
        first_pass = 0
        for i, r in dups_by_room_df_temp.iterrows():
            dups_by_room_df_temp = dups_by_room_df_temp.sort_values(['Created'],ascending=[True])
            if first_pass != 0:
                if r['Original Index'] not in list_to_delete_room:
                    list_to_delete_room.append(r['Original Index'])
                    dict_to_keep_room[r['Original Index']] = False
            else:
                dict_to_keep_room[r['Original Index']] = True
            first_pass += 1

# Map the dict_to_keep_room to the dups_by_room_df so that the user can see which ones are being kept
dups_by_room_df['Room Duplicate To Keep'] = dups_by_room_df['Original Index'].map(dict_to_keep_room)
dups_by_room_df['Room Duplicate To Keep'] = dups_by_room_df['Room Duplicate To Keep'].fillna(value=False)

# Drop all rows which were identified as needing dropping
combined_df.drop(list_to_delete_room, axis=0, inplace=True)

# DUPLICATES BY TENANT
# --------------------
# All rows where there are duplicates by Tenant (i.e. 2 contracts for the same tenant)
dups_by_tenant_df = combined_df[combined_df['Twin Occupancy'] == False]
dups_by_tenant_df = dups_by_tenant_df[dups_by_tenant_df.duplicated(['Property Name', 'User Id'], keep=False)]

# This section checks the date compared the row above to see whether the contracts for the same tenant overlap
dups_by_tenant_df = dups_by_tenant_df.sort_values(['Property Name', 'User Id', 'Contract Start', 'Contract End'],
                                              ascending=[True, True, True, True])

dups_by_tenant_df['Prev Property Name'] = dups_by_tenant_df['Property Name'].shift(1)
dups_by_tenant_df['Prev Tenant'] = dups_by_tenant_df['User Id'].shift(1)
dups_by_tenant_df['Prev Contract Start'] = dups_by_tenant_df['Contract Start'].shift(1)
dups_by_tenant_df['Prev Contract End'] = dups_by_tenant_df['Contract End'].shift(1)

# Save original index as we will reset the index
dups_by_tenant_df['Original Index'] = dups_by_tenant_df.index

dups_by_tenant_df = dups_by_tenant_df.reset_index()

dups_by_tenant_df['Tenant Duplicate'] = False
tenant_duplicate_id = 20000
dups_by_tenant_df['Tenant Duplicate ID'] = 0
dups_by_tenant_df['Date Overlap'] = False

# This section loops through all of the duplicates by tenant and then assigns a duplicate ID so that they are linked
for index,row in dups_by_tenant_df.iterrows():
    #if dups_by_tenant_df.loc[index, 'Contract Start'] != np.datetime64('NaT'):
    if row['User Id'] == row['Prev Tenant'] and row['Contract Start'] < row['Prev Contract End']:
        dups_by_tenant_df.loc[index - 1, 'Tenant Duplicate'] = True
        dups_by_tenant_df.loc[index, 'Tenant Duplicate'] = True
        dups_by_tenant_df.loc[index - 1, 'Date Overlap'] = True
        dups_by_tenant_df.loc[index, 'Date Overlap'] = True

        if dups_by_tenant_df.loc[index - 1, 'Tenant Duplicate ID'] == 0:
            tenant_duplicate_id += 1
            dups_by_tenant_df.loc[index - 1, 'Tenant Duplicate ID'] = tenant_duplicate_id
            dups_by_tenant_df.loc[index, 'Tenant Duplicate ID'] = tenant_duplicate_id
        else:
            dups_by_tenant_df.loc[index, 'Tenant Duplicate ID'] = dups_by_tenant_df.loc[index - 1, 'Tenant Duplicate ID']

# Removed all valid duplicates (dates not overlapping etc) leaving only ones with overlapping dates to be cleaned
dups_by_tenant_df = dups_by_tenant_df[dups_by_tenant_df['Tenant Duplicate'] == True]

# Sort duplicates by Tenant Duplicate ID, Created and Contract State as this determines what we select to remove
dups_by_tenant_df = dups_by_tenant_df.sort_values(['Tenant Duplicate ID', 'Created', 'Contract State'],
                                              ascending=[True, True, True])

# Save original index as we will reset the index
#dups_by_tenant_df['Original Index'] = dups_by_tenant_df.index

dups_by_tenant_df = dups_by_tenant_df.reset_index()

previous_tenant_id = 0
list_to_delete_tenant = []
dict_to_keep_tenant = {}

for index,row in dups_by_tenant_df.iterrows():
    if row['Tenant Duplicate ID'] != previous_tenant_id:
        dups_by_tenant_df_temp = dups_by_tenant_df[dups_by_tenant_df['Tenant Duplicate ID'] == row['Tenant Duplicate ID']]
        previous_tenant_id = row['Tenant Duplicate ID']
        #print(row['Tenant Duplicate ID'], "\n", "------------- \n", dups_by_tenant_df_temp)

        # Checks whether any of the duplicates for the tenant have Contract State = "Signed"
        contains_signed = False
        for i, r in dups_by_tenant_df_temp.iterrows():
            if r['Contract State'] == "Signed":
                contains_signed = True

        # If any of the duplicates for the tenant have Contract State = "Signed", drop all others (add to list)
        if contains_signed == True:
            for i, r in dups_by_tenant_df_temp.iterrows():
                if r['Contract State'] != "Signed":
                    list_to_delete_tenant.append(r['Original Index'])
                    #dups_by_tenant_df_temp2.drop(i, axis = 0, inplace = True)

        # Drops all rows other than the oldest tenant based on Created date (add to llist)
        first_pass = 0
        for i, r in dups_by_tenant_df_temp.iterrows():
            dups_by_tenant_df_temp = dups_by_tenant_df_temp.sort_values(['Created'],ascending=[True])
            if first_pass != 0:
                if r['Original Index'] not in list_to_delete_tenant:
                    list_to_delete_tenant.append(r['Original Index'])
                    dict_to_keep_tenant[r['Original Index']] = False
            else:
                dict_to_keep_tenant[r['Original Index']] = True
            first_pass += 1

# Map the dict_to_keep_tenant to the dups_by_tenant_df so that the user can see which ones are being kept
dups_by_tenant_df['Tenant Duplicate To Keep'] = dups_by_tenant_df['Original Index'].map(dict_to_keep_tenant)
dups_by_tenant_df['Tenant Duplicate To Keep'] = dups_by_tenant_df['Tenant Duplicate To Keep'].fillna(value=False)

# Drop all rows which were identified as needing dropping
combined_df.drop(list_to_delete_tenant, axis=0, inplace=True)
# At this stage we have cleaned the data and need to create a summary of the data

# --------------------
# SUMMARY AND ANALYSIS
# --------------------

# New way of doing this (by iterating through properties and creating a dictionary with summary results)
property_list = ["DRM-Ernest Place", "NCL-Mansion Tyne", "BRS-Centre Gate", "BRS-The Colston", "MCR-Mansion Point", "NOT-Mansion Place", "SHF-Redvers Tower"]

summary_dict = {}

for p in property_list:
    # By Beds
    site_capacity = capacity_dict[p]
    total_lt_bookings = combined_df[(combined_df['Property Name'] == p) & (combined_df['LT or ST'] == 'LT')].nunique()["Contract ID"]
    total_st_bookings = combined_df[(combined_df['Property Name'] == p) & (combined_df['LT or ST'] == 'ST')].nunique()["Contract ID"]
    total_bookings = total_lt_bookings + total_st_bookings
    total_bookings_lt_eqv = combined_df[(combined_df['Property Name'] == p)]['LT Equivalent'].sum()
    occupancy_by_bed = total_bookings_lt_eqv / site_capacity
    # By Rent
    quoting_rent = total_potential_rent[p]
    total_lt_income = combined_df[(combined_df['Property Name'] == p) & (combined_df['LT or ST'] == 'LT')]['Contract Total Rent'].sum()
    total_st_income = combined_df[(combined_df['Property Name'] == p) & (combined_df['LT or ST'] == 'ST')]['Contract Total Rent'].sum()
    total_income = total_lt_income + total_st_income
    occupancy_by_income = total_income / quoting_rent
    # By Nationality
    total_bookings_uk = combined_df[(combined_df['Property Name'] == p) & (combined_df['Domestic or International'] == 'Domestic')]['Contract State'].count()
    total_income_uk = combined_df[(combined_df['Property Name'] == p) & (combined_df['Domestic or International'] == 'Domestic')]['Contract Total Rent'].sum()
    total_bookings_intl = combined_df[(combined_df['Property Name'] == p) & (combined_df['Domestic or International'] == 'International')]['Contract State'].count()
    total_income_intl = combined_df[(combined_df['Property Name'] == p) & (combined_df['Domestic or International'] == 'International')]['Contract Total Rent'].sum()
    # High Risk of Default
    total_bookings_high_risk = combined_df[(combined_df['Property Name'] == p) & (combined_df['High Risk of Default'] == True)]['Contract State'].count()
    total_income_high_risk = combined_df[(combined_df['Property Name'] == p) & (combined_df['High Risk of Default'] == True)]['Contract Total Rent'].sum()
    # By bookings
    total_records = combined_df[combined_df['Property Name'] == p]['Contract State'].count()
    total_beds_currently_occupied = combined_df[(combined_df['Property Name'] == p) & (combined_df['In Room'] == True)].nunique()["Contract ID"]
    total_beds_currently_occupied_pct = total_beds_currently_occupied / site_capacity
    #total_records2 = combined_df['Contract State'][combined_df['Property Name'] == p].count()
    # Averages
    total_days = combined_df[(combined_df['Property Name'] == p)]['Contract Days'].sum()
    total_weeks = combined_df[(combined_df['Property Name'] == p)]['Contract Weeks'].sum()
    avg_wks = total_weeks / total_records
    avg_weekly_rent = total_income / total_weeks
    # Extra Analysis
    total_rent_paid = combined_df[(combined_df['Property Name'] == p)]['Total Rent Paid'].sum()
    total_rent_due = combined_df[(combined_df['Property Name'] == p)]['Total Rent Due'].sum()
    total_arrears_value = combined_df[(combined_df['Property Name'] == p)]['Total Rent Arrears'].sum()
    total_arrears_number = combined_df[(combined_df['Property Name'] == p) & (combined_df['Total Rent Arrears'] > 0)]['Total Rent Arrears'].count()

    summary_dict.update({p: ({
        'Site Capacity': site_capacity,
        'Total LT Bookings': total_lt_bookings,
        'Total ST Bookings': total_st_bookings,
        'Total Bookings': total_bookings,
        'Total Bookings LT Eqv': total_bookings_lt_eqv,
        'Occupancy By Bed': occupancy_by_bed,
        'Quoting Rent': quoting_rent,
        'Total Income': total_income,
        'Total LT Income': total_lt_income,
        'Total ST Income': total_st_income,
        'Occupancy By Income': occupancy_by_income,
        'Average Weeks Per Contract': avg_wks,
        'Average Rent Per Week': avg_weekly_rent,
        'Total Domestic Bookings': total_bookings_uk,
        'Total Domestic Income': total_income_uk,
        'Total International Bookings': total_bookings_intl,
        'Total International Income': total_income_intl,
        'Total High Risk Bookings': total_bookings_high_risk,
        'Total High Risk Income': total_income_high_risk,
        'Total Records': total_records,
        'Total Beds Currently Occupied': total_beds_currently_occupied,
        'Total Beds Currently Occupied %': total_beds_currently_occupied_pct,
        'Total Rent Paid': total_rent_paid,
        'Total Rent Due': total_rent_due,
        'Total Arrears': total_arrears_value,
        'Total Students in Arrears': total_arrears_number})
    })

summary_dict_to_df = pd.DataFrame.from_dict(summary_dict, orient='index')

def custom_group_by_sum (df, filter_col, filter_val, group_col, agg_col):
    return df[df[filter_col] == filter_val].groupby(group_col)[agg_col].sum()

# Groupby Results

grouped_multiple = combined_df.groupby(['Property Name', 'Checked In', 'Contract State', 'Domestic EU or International', 'Guarantor Signed Yes or No']).agg({'Contract Total Rent': ['mean', 'min', 'max']})

pivot_multiple1 = pd.pivot_table(combined_df, index=['Property Name','Checked In', 'Domestic EU or International'],values=['Contract Total Rent','Total Rent Paid','Total Rent Due', 'Total Rent Arrears'],aggfunc=np.sum)
pivot_multiple2 = pd.pivot_table(combined_df, index=['Property Name','Checked In', 'Contract State', 'Domestic EU or International', 'Guarantor Signed Yes or No'],values=['Contract Total Rent','Total Rent Paid','Total Rent Due', 'Total Rent Arrears'],aggfunc=np.sum)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Need to spend more time on this section
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Function used to create a pivot table for a list of indexes and calculate the suma and count of rent and LTeqv
def pivot_table_rent_lteqv(df,p,index_list):
    pivot = pd.pivot_table(df[df['Property Name'] == p],
                   index=index_list,
                   values=['Contract Total Rent', 'LT Equivalent'],
                   aggfunc={'Contract Total Rent': (np.sum, (lambda x: x.sum() / df[df['Property Name'] == p]['Contract Total Rent'].sum())),
                            'LT Equivalent': (np.sum, (lambda x: x.sum() / df[df['Property Name'] == p]['LT Equivalent'].sum()))},
                   dropna=False)

    pivot.columns.set_levels(["%", "sum"], level =1, inplace=True)

    #print(type(pivot), "\n", pivot.shape, "\n", pivot.columns, "\n", pivot.columns.levels[1], "\n", pivot.index)
    #pivot.columns = pivot.columns.droplevel(1)

    return pivot

def pivot_table_rent_lteqv_occupied(df,p,index_list):
    pivot = pd.pivot_table(df[df['Property Name'] == p],
                   index=index_list,
                   values=['Contract Total Rent', 'LT Equivalent', 'In Room Count'],
                   aggfunc={'Contract Total Rent': (np.sum, (lambda x: x.sum() / df[df['Property Name'] == p]['Contract Total Rent'].sum())),
                            'LT Equivalent': (np.sum, (lambda x: x.sum() / df[df['Property Name'] == p]['LT Equivalent'].sum())),
                            'In Room Count': (np.sum, (lambda x: x.sum() / df[df['Property Name'] == p]['In Room Count'].sum()))
                            },
                   dropna=False)
    pivot.columns = pivot.columns.droplevel(1)
    #pivot.rename(columns={pivot.columns[0]: "Total Rent", pivot.columns[1]: "% Based on  Rent", pivot.columns[2]: "Beds", pivot.columns[3]: "% Based on Beds",}, inplace=True)
    return pivot

individual_dict = {}

for p in property_list:
    individual_dict.update({p:
        ({
        'Summary': summary_dict_to_df.loc[p, :],
        'Room Type': pivot_table_rent_lteqv(combined_df, p, ['Cluster or Studio']),
        'Long or Short Term': pivot_table_rent_lteqv(combined_df, p, ['LT or ST']),
        'Domestic or International': pivot_table_rent_lteqv(combined_df, p, ['Domestic EU or International']),
        'Nationality': pivot_table_rent_lteqv(combined_df, p, ['Nationality']),
        'University': pivot_table_rent_lteqv(combined_df, p, ['University']),
        'Year of Study': pivot_table_rent_lteqv(combined_df, p, ['Year of Study']),
        'Booking Source': pivot_table_rent_lteqv(combined_df, p, ['Booking Source']),
        'Arrers': pivot_table_rent_lteqv(combined_df, p, ['In Arrears']),
        'At Risk Pivot': pd.pivot_table(combined_df[combined_df['Property Name'] == p],
                                    values=['Contract Total Rent', 'Total Rent Paid', 'Total Rent Due', 'Total Rent Arrears', 'Contract ID', 'LT Equivalent'],
                                    index=['Checked In', 'Never Paid', 'Contract State', 'Domestic EU or International', 'Guarantor Signed Yes or No'],
                                    aggfunc={'Contract Total Rent': np.sum, 'Total Rent Paid': np.sum, 'Total Rent Due': np.sum, 'Total Rent Arrears': np.sum, 'Contract ID': 'count', 'LT Equivalent': np.sum}),
        'Not Checked In, Never Paid, No Guarantor': combined_df[(combined_df['Property Name'] == p) & (combined_df['Checked In'] == False) & (combined_df['Never Paid'] == True) & (combined_df['Guarantor Signed Yes or No'] == 'No Guarantor')]['Tenant']
        #'grouped_multiple_per_property': combined_df[(combined_df['Property Name'] == p)].groupby(['Domestic EU or International']).agg({'Contract Total Rent': ['mean', 'min', 'max']}),
        #'pivot_multiple1': pd.pivot_table(combined_df, index=['Property Name', 'Checked In', 'Domestic EU or International'],values=['Contract Total Rent', 'Total Rent Paid', 'Total Rent Due', 'Total Rent Arrears'], aggfunc=np.sum),
        #'pivot_multiple2': pd.pivot_table(combined_df, index=['Property Name', 'Checked In', 'Contract State','Domestic EU or International', 'Guarantor Signed Yes or No'],values=['Contract Total Rent', 'Total Rent Paid', 'Total Rent Due','Total Rent Arrears'], aggfunc=np.sum)
        })
    })

# Remove duplicates on list of high risk
# Add bookings taken in past week, and current month and previous month.


# Output to Excel
# ---------------

if run_export == True:

    output_path = 'C:\\Users\\daniel.radnor\\PycharmProjects\\Lettings\\Lettings Report Automation\\'+ str(date.today())+'output.xlsx'

    with pd.ExcelWriter(output_path) as writer:

        workbook = writer.book
        format1 = workbook.add_format({'num_format': '#,##0'})
        format2 = workbook.add_format({'num_format': '0%'})
        format3 = workbook.add_format({'num_format': '#,##0'})

        # Export the summary analysis on all properties
        summary_dict_to_df.to_excel(writer, sheet_name='Analysis Summary', index=True)

        worksheet_summary = writer.sheets['Analysis Summary']
        worksheet_summary.set_column('A:Z', 18)
        worksheet_summary.set_column('B:F', 18, format1)
        worksheet_summary.set_column('G:G', 18, format2)
        worksheet_summary.set_column('H:K', 18, format1)
        worksheet_summary.set_column('L:L', 18, format2)
        worksheet_summary.set_column('M:V', 18, format1)
        worksheet_summary.set_column('W:W', 18, format2)
        worksheet_summary.set_column('X:AA', 18, format1)

        # Export all the cleaned data
        combined_df.to_excel(writer, sheet_name='Clean Data', index=False)

        # Export all of the property level information to excel (single sheet epr property)
        for index, value in individual_dict.items():
            # This exports to Excel
            start_row_agg = 0
            for i, v in value.items():
                v.to_excel(writer, sheet_name=index, index=True, startrow=start_row_agg, startcol=0)
                #print("V shape: ", v.shape[0], ", start_row_agg: ", start_row_agg)
                start_row_agg += v.shape[0] + 5

            worksheet = writer.sheets[index]
            worksheet.set_column('A:K', 18)
            worksheet.set_column('B:K', 18, format1)
            worksheet.set_column('B:B', 18, format2)
            worksheet.set_column('D:D', 18, format2)
            for i in range(1,27):
                worksheet.set_row(i, None, format1)

            worksheet.set_row(6, None, format2)
            worksheet.set_row(11, None, format2)
            worksheet.set_row(22, None, format2)

        # Export remaining information
        original_combined_df.to_excel(writer, sheet_name='Original Data', index=False)
        twin_occupancy_df.to_excel(writer, sheet_name='Twin Occupancy', index=False)
        dups_by_room_df.to_excel(writer, sheet_name='Duplicates By Room', index=False)
        dups_by_tenant_df.to_excel(writer, sheet_name='Duplicates By Tenant', index=False)
        not_2020_21_df.to_excel(writer, sheet_name='Not 2020_21', index=False)
        expired_pending_df.to_excel(writer, sheet_name='Old Pending Signatures', index=False)
        cancelled_df.to_excel(writer, sheet_name='Cancelled Contracts', index=False)

end_time = time.perf_counter()

total_time = end_time - start_time

print("Time taken to run code is %s" % total_time)

#  End of main code.

