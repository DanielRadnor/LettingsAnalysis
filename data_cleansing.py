import variables as vb
import functions as fn
import data_import as dti
import pandas as pd
from datetime import date, datetime, timedelta

combined_df = dti.combined_df.copy() # Deep copy required.

# Add Property Name to use as part of merge as some properties have 2 spellings
combined_df['Property Name'] = combined_df.apply(lambda x: fn.property_name(x['Address']), axis=1)

# Bring Property Name to the first column position
prop_name = combined_df['Property Name']
combined_df.drop(labels=['Property Name'], axis=1,inplace = True)
combined_df.insert(0, 'Property Name', prop_name)

# Remove all records where the rent is 0
combined_df = combined_df[(combined_df['Contract Total Rent'] > 0)]

# Add column to denote academic year
combined_df['Academic Year'] = combined_df.apply(lambda x: fn.booking_academic_year(x['Contract Start'], x['Contract End']), axis = 1)

# Converts all room numbers to strings
combined_df['Room Number'] = combined_df['Room Number'].apply(str)

# Add property bed capacity for each row
combined_df['Capacity'] = combined_df['Property Name'].map(vb.capacity_dict)

# Add property total potential rent for each row
combined_df['Total Property Potential Rent'] = combined_df['Property Name'].map(vb.total_potential_rent)

# Add Daily Rent rate
combined_df['Contract Daily Rent'] = combined_df.apply(lambda x: (x['Contract Total Rent'] / (x['Contract Days'])), axis =1)

# Add Cluster or Studio column
combined_df['Cluster or Studio'] = combined_df.apply(lambda x: ("Studio" if 'Studio' in x['Room Type'] or 'Twodio' in x['Room Type'] or 'Threedio' in x['Room Type'] or 'Apartment' in x['Room Type'] else "Cluster"), axis = 1)

# Add ST/LT column
combined_df['LT or ST'] = combined_df.apply(lambda x: ("LT" if x['Contract Weeks'] >= vb.lt_weeks else "ST"), axis = 1)

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
combined_df['Domestic or International'] = combined_df.apply(lambda x: fn.domestic_or_intl(x['Nationality']), axis = 1)
combined_df['Domestic EU or International'] = combined_df.apply(lambda x: fn.eu_or_intl(x['Nationality']), axis = 1)

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
combined_df['Pending Expired'] = combined_df.apply(lambda x: (True if (x['Contract State'] == "Pending Signatures" and ((datetime.now() - x['Created']) > timedelta(days=vb.expire_pending))) else False), axis = 1)

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
cancelled_df['Cancellation Date'] = cancelled_df.apply(lambda x: fn.extract_date_from_str(x['Contract State']), axis = 1)

# This removes Pending Signature which are more than "vb.expire_pending" days old
expired_pending_df = combined_df[((combined_df['Contract State'] == "Pending Signatures") & ((datetime.now() - combined_df['Created']) > timedelta(days=vb.expire_pending)))]
combined_df.drop(expired_pending_df.index.tolist(), axis = 0, inplace = True)

# Add column to mark which contracts are 2020_21 and then remove from the main DataFrame
#combined_df['2020_21 Contracts'] = combined_df.apply(lambda x: fn.booking_in_2020_21(x['Contract Start'], x['Contract End']), axis=1)
#not_2020_21_df = combined_df[combined_df['2020_21 Contracts'] == False] # Create dataframe for  rows not for 2020_21 to new dataframe (these will be removed)
#list_to_delete_not_2020_21 = combined_df.index[combined_df['2020_21 Contracts'] == False].tolist()
#combined_df.drop(list_to_delete_not_2020_21, axis=0, inplace=True)

# Remove records for non-relevant academic years
combined_df = combined_df[(combined_df['Academic Year'] == vb.academic_year_selection)]

# MULTI OCCUPANCY
# --------------
# Duplicates by contract ID remain as the Rent fields are split across the tenants which share a bed.  Just need to identify

# Start by using GroupBy to count the number of occupants per room
multi_occupancy_groups = combined_df.groupby(['Contract ID','Property Name', 'Room Number', 'Contract Start', 'Contract End']) #.agg({'Contract ID': 'count'})
multi_occupancy_groups.apply(lambda x: x)
size = multi_occupancy_groups.size().reset_index()
multi_occupancy_list_df = size#.to_frame()
multi_occupancy_list_df = multi_occupancy_list_df.rename(columns={0: 'Number of Occupants'})
# MErge with original dataframe and remove the duplicated columns
combined_df = combined_df.merge(multi_occupancy_list_df, left_on=['Contract ID','Property Name', 'Room Number', 'Contract Start', 'Contract End'],
                                right_on=['Contract ID','Property Name', 'Room Number', 'Contract Start', 'Contract End'],
                                suffixes=('', '_y'))

# Removes columns which have been duplicated (i.e. which were in both excel worksheets)
combined_df.drop(combined_df.filter(regex='_y$').columns.tolist(),axis=1, inplace=True)

# Create list of rows which are Multi occupancy and mark them in new Multi Occupancy column
combined_df['Multi Occupancy'] = combined_df.apply(lambda x: ('Single' if x['Number of Occupants'] == 1 else 'Twin' if x['Number of Occupants'] == 2 else 'Triple' if x['Number of Occupants'] == 3 else 'More than 3'), axis = 1)
combined_df['Contract Days Adj For Multi Occupancy'] = combined_df.apply(lambda x: (x['Contract Days'] / x['Number of Occupants']), axis = 1)

multi_occupancy_df = combined_df[combined_df['Number of Occupants'] > 1]

# LT EQUIVALENT
# -------------

# If LT then 1 else fraction of 43.5 making sure multi occupancy are halved
combined_df['LT Equivalent'] = combined_df.apply(lambda x: (1 if x['LT or ST'] == 'LT' else x['Contract Weeks'] / vb.lt_weeks), axis = 1)
combined_df['LT Equivalent'] = combined_df.apply(lambda x: (x['LT Equivalent'] / x['Number of Occupants']), axis = 1)

# DUPLICATES BY ROOM
# ------------------
# All rows where there are duplicates by Room Number (i.e. 2 contracts for the same room)
# =========================================================================================================================
# If there is the odd case of a Multi Occupancy which is a duplicate with a normal occupancy then this will not be reported
# =========================================================================================================================
dups_by_room_df = combined_df[combined_df['Multi Occupancy'] == 'Single']
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
dups_by_tenant_df = combined_df[combined_df['Multi Occupancy'] == 'Single']
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
