import variables as vb
import pandas as pd
import re
from datetime import datetime

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

# Returns the 'Capacity" field in the first row of the dataframe
def capacity(dataframe):
    cp = dataframe['Capacity'].loc[0]
    return cp

# Returns the 'Capacity" in the dataframe for a specified property name
def capacity2(dataframe, property_name):
    cp2 = dataframe[dataframe['Property Name'] == property_name].max()["Capacity"]
    return cp2

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
    tltr = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] >= vb.lt_weeks)].sum()["Contract Total Rent"]
    return tltr

# Returns the total Short Term rent achieved  for a property
def total_st_rent(dataframe, property_name):
    tstr = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < vb.lt_weeks)].sum()["Contract Total Rent"]
    return tstr

# Returns the total bookings achieved for a property by counting the number of times there is a Contract Total Rent
def total_bookings(dataframe, property_name):
    tb = dataframe[(dataframe['Property Name'] == property_name)
                   & (dataframe['Contract Total Rent'] > 0)].nunique()["Contract ID"]
    return tb

# Total achieved Long Term bookings for a property (counts twin occupancy as 1 bed)
def total_lt_bookings(dataframe, property_name):
    tltb = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] >= vb.lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].nunique()["Contract ID"]
    return tltb

# Total achieved Short Term bookings for a property (counts twin occupancy as 1 bed)
def total_st_bookings(dataframe, property_name):
    tstb = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < vb.lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].nunique()["Contract ID"]
    return tstb

# Total achieved Short Term bookings for a property converted to long term equivalents
def total_st_bookings_lteqv(dataframe, property_name):
    tstblteqv = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < vb.lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].sum()["Contract Weeks"] / vb.lt_weeks
    return tstblteqv

# Total achieved Short Term bookings for a property in terms of total weeks booked
def total_st_bookings_weeks(dataframe, property_name):
    tstbwk = dataframe[(dataframe['Property Name'] == property_name) & (dataframe['Contract Weeks'] < vb.lt_weeks)
                     & (dataframe['Contract Total Rent'] > 0)].sum()["Contract Weeks"]
    return tstbwk

# Check whether a booking is in 2020/21 based on 50% or more of the booking days after 5th September
def booking_in_2020_21(start_date,end_date):
    vb.start_date_2020_21 = '2020-09-05'
    vb.start_date_2020_21 = pd.to_datetime(vb.start_date_2020_21)
    if start_date > vb.start_date_2020_21:
        return True
    elif (end_date - vb.start_date_2020_21) >= (vb.start_date_2020_21 - start_date):
        return True
    else:
        return False

# Check whether a booking is in 2020/21 based on 50% or more of the booking days after 5th September
def booking_academic_year(start_date,end_date):

    start_date_2019_20 = pd.to_datetime('2019-09-05')
    end_date_2019_20 = pd.to_datetime('2020-09-04')

    start_date_2020_21 = pd.to_datetime('2020-09-05')
    end_date_2020_21 = pd.to_datetime('2021-09-04')

    start_date_2021_22 = pd.to_datetime('2021-09-05')
    end_date_2021_22 = pd.to_datetime('2022-09-04')

    proportion_dict = {
        'portion_19_20': min(end_date, end_date_2019_20) - max(start_date, start_date_2019_20),
        'portion_20_21': min(end_date, end_date_2020_21) - max(start_date, start_date_2020_21),
        'portion_21_22': min(end_date, end_date_2021_22) - max(start_date, start_date_2021_22)
    }

    max_days = proportion_dict['portion_19_20']

    for i, v in proportion_dict.items():
        if v >= max_days:
            max_days = v
            ay = i

    academic_year = ay[-5:]

    return academic_year

# Check whether a student is Domestic or International
def domestic_or_intl (country):
    if any(x in country for x in vb.british_country_list):
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
        if any(x in country for x in vb.eu_country_list):
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