import variables as vb
import data_cleansing as dcl
import pandas as pd
import numpy as np

combined_df = dcl.combined_df.copy() # Deep copy required.

# --------------------
# SUMMARY AND ANALYSIS
# --------------------

# New way of doing this (by iterating through properties and creating a dictionary with summary results)

summary_dict = {}

for p in vb.property_list:
    # By Beds
    site_capacity = vb.capacity_dict[p]
    total_lt_bookings = combined_df[(combined_df['Property Name'] == p) & (combined_df['LT or ST'] == 'LT')].nunique()["Contract ID"]
    total_st_bookings = combined_df[(combined_df['Property Name'] == p) & (combined_df['LT or ST'] == 'ST')].nunique()["Contract ID"]
    total_bookings = total_lt_bookings + total_st_bookings
    total_bookings_lt_eqv = combined_df[(combined_df['Property Name'] == p)]['LT Equivalent'].sum()
    occupancy_by_bed = total_bookings_lt_eqv / site_capacity
    # By Rent
    quoting_rent = vb.total_potential_rent[p]
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
    if total_records == 0:
        avg_wks = 'NA'
    else:
        avg_wks = total_weeks / total_records
    if total_weeks == 0:
        avg_weekly_rent = 'NA'
    else:
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

for p in vb.property_list:
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
