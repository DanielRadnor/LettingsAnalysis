import variables as vb
import data_cleansing as dcl
import data_analysis as da
import pandas as pd
import numpy as np

# Output to Excel
# ---------------

if vb.run_export == True:

    with pd.ExcelWriter(vb.output_path) as writer:

        workbook = writer.book
        format1 = workbook.add_format({'num_format': '#,##0'})
        format2 = workbook.add_format({'num_format': '0%'})
        format3 = workbook.add_format({'num_format': '#,##0'})

        # Export the summary analysis on all properties
        da.summary_dict_to_df.to_excel(writer, sheet_name='Analysis Summary', index=True)

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
        da.combined_df.to_excel(writer, sheet_name='Clean Data', index=False)

        # Export all of the property level information to excel (single sheet epr property)
        for index, value in da.individual_dict.items():
            # This exports to Excel
            start_row_agg = 0

            for i, v in value.items():
                if len(v) > 0:
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
        dcl.original_combined_df.to_excel(writer, sheet_name='Original Data', index=False)
        dcl.multi_occupancy_df.to_excel(writer, sheet_name='Multi Occupancy', index=False)
        dcl.dups_by_room_df.to_excel(writer, sheet_name='Duplicates By Room', index=False)
        dcl.dups_by_tenant_df.to_excel(writer, sheet_name='Duplicates By Tenant', index=False)
        #dcl.not_2020_21_df.to_excel(writer, sheet_name='Not 2020_21', index=False)
        dcl.expired_pending_df.to_excel(writer, sheet_name='Old Pending Signatures', index=False)
        dcl.cancelled_df.to_excel(writer, sheet_name='Cancelled Contracts', index=False)
        dcl.cancelled_checked_out_df.to_excel(writer, sheet_name='Checked Out Cancelled Contracts', index=False)
