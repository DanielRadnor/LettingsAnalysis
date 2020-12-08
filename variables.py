from datetime import date, datetime, timedelta

# INPUT VARIABLES
# ---------------

# Select Academic Year to Run Report
#academic_year_selection = '19_20'
academic_year_selection = '20_21'
#academic_year_selection = '21_22'

# Pathname for raw files where data is imported from
path = "C:\\Users\\daniel.radnor\\PycharmProjects\\Lettings\\Lettings Report Automation\\Raw Files"

# Pathname for exporting analysis
output_path = 'C:\\Users\\daniel.radnor\\PycharmProjects\\Lettings\\Lettings Report Automation\\' + str(date.today()) + ' -  Lettings Report - '  + academic_year_selection +  '.xlsx'
output_test_path = 'C:\\Users\\daniel.radnor\\PycharmProjects\\Lettings\\Lettings Report Automation\\' + str(date.today()) + ' - Room and Tenant Report - ' + academic_year_selection + '.xlsx'

# List of proeprties
property_list = ["DRM-Ernest Place", "NCL-Mansion Tyne", "BRS-Centre Gate", "BRS-The Colston", "MCR-Mansion Point", "NOT-Mansion Place", "SHF-Redvers Tower"]

# Minimum number of weeks considered as Long Term
lt_weeks = 43.5

# Max days a contract can be Pending Signatures before being removed
expire_pending = 14

# Switch to run excel export or not
run_export = True

# Dates for 20/21 Academic year
start_date_2020_21 = "2020-09-05"
end_date_2020_21 = '2021-09-03'

start_date_2021_22 = "2021-09-04"
end_date_2021_22 = '2022-09-03'

# ------------
# DICTIONARIES
# ------------

# Total capacity at each property
capacity_dict = {"LDS-Leodis Residence": 719,
            "DRM-Ernest Place": 362,
            "MCR-Mansion Point": 206,
            "NCL-Mansion Tyne": 416,
            "BRS-Centre Gate": 85,
            "BRS-The Colston": 77,
            "NOT-Mansion Place": 139,
            "SHF-Redvers Tower": 227}

# Dict for the total potential rent at each property
if academic_year_selection == '20_21':
    total_potential_rent = {"LDS-Leodis Residence": 3422562.23,
                 "DRM-Ernest Place": 2846820,
                 "MCR-Mansion Point": 1727370,
                 "NCL-Mansion Tyne": 2268684,
                 "BRS-Centre Gate": 779586,
                 "BRS-The Colston": 711552,
                 "NOT-Mansion Place": 1036422,
                 "SHF-Redvers Tower": 1994049}
elif academic_year_selection == '21_22':
    total_potential_rent = {"LDS-Leodis Residence": 3790362.49,
                 "DRM-Ernest Place": 2750175,
                 "MCR-Mansion Point": 1769241,
                 "NCL-Mansion Tyne": 2301630,
                 "BRS-Centre Gate": 829617,
                 "BRS-The Colston": 780045,
                 "NOT-Mansion Place": 1065033,
                 "SHF-Redvers Tower": 1748280}

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
