import variables as vb
import data_cleansing as dcl
import data_analysis as da
import pandas as pd
import numpy as np
from functools import reduce

#def custom_group_by_sum (df, filter_col, filter_val, group_col, agg_col):
#    return df[df[filter_col] == filter_val].groupby(group_col)[agg_col].sum()

# Groupby Results

grouped_dict = {}

for p in vb.property_list:

    grouped_by_room = dcl.combined_df[(dcl.combined_df['Property Name'] == p)].groupby(['Room Number', 'Room Type']).agg({'Contract ID': 'count', 'Contract Total Rent': 'sum','Contract Days Halved For Twin Occupancy': 'sum','Total Rent Arrears': 'sum'})
    grouped_by_room['Max Remaining Days To Sell'] = grouped_by_room.apply(lambda x: 51 * 7 - x['Contract Days Halved For Twin Occupancy'], axis =1)

    grouped_by_room.rename(columns={'Contract ID': 'Number of Contracts', 'Contract Total Rent': 'Total Contracted Rent for Room', 'Contract Days Halved For Twin Occupancy': 'Contracted Days'}, inplace=True)

    grouped_by_tenant = dcl.combined_df[(dcl.combined_df['Property Name'] == p)].groupby(['User Id', 'Tenant']).agg({'Contract ID': 'nunique', 'Room Number': 'count', 'Contract Days Halved For Twin Occupancy': 'sum', 'Contract Total Rent': 'sum','Total Rent Paid': 'sum','Total Rent Arrears': 'sum'})
    grouped_by_tenant.rename(columns={'Contract ID': 'Number of Contracts', 'Room Number': 'Number of Rooms Occupied', 'Contract Days Halved For Twin Occupancy': 'Contracted Days'}, inplace=True)

    #grouped_by_room_contract_id = dcl.combined_df[(dcl.combined_df['Property Name'] == p)].groupby(['Room Number', 'Contract ID']).agg({'Contract ID': 'count', 'Contract Total Rent': 'sum','Contract Days Halved For Twin Occupancy': 'sum','Total Rent Arrears': 'sum'})
    #grouped_by_room_tenant = dcl.combined_df[(dcl.combined_df['Property Name'] == p)].groupby(['Room Number', 'Contract ID', 'Tenant', 'Contract Start', 'Contract End']).agg({'Contract Total Rent': ['sum','count'],'Total Rent Paid': 'sum','Total Rent Arrears': 'sum'})

    #df_merged = reduce(lambda left, right: pd.merge(left, right, on=['Room Number'],how='outer'), [grouped_by_room,grouped_by_room_contract_id])

    grouped_dict.update({p: ({
        'Grouped By Room': grouped_by_room,
        'Grouped By Tenant': grouped_by_tenant})
    })

if vb.run_export == True:

    with pd.ExcelWriter(vb.output_test_path) as writer:

        for index, value in grouped_dict.items():
            # This exports to Excel
            start_row_agg = 0
            for i, v in value.items():
                v.to_excel(writer, sheet_name=index, index=True, startrow=start_row_agg, startcol=0)
                #print("V shape: ", v.shape[0], ", start_row_agg: ", start_row_agg)
                start_row_agg += v.shape[0] + 5