# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
from collections import Counter
from fuzzywuzzy import fuzz
import numpy as np
import datetime
import jaro

starttime = datetime.datetime.now()
print("Start: ", starttime)

#%% ####################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = 'JARO_NAME'
# Put the filepath to the GL/Other data
df_filepath = 'Data/Loan_Name.xlsx'
# What column are we comparing?
dfc = {'LOAN_NAME': 'LOAN_NAME'}
# IF the column in the LASTNAME, FIRST NAME format, type 'YES'
df_reverse = 'YES'

# Put the filepath to the related parties
rp_filepath = 'Data/Related Parties Clean.xlsx'
# What column are we comparing?
rpc = {'RP_NAME': 'OTHER'} #NAME, OCC, BUS, OTHER
# Is this column in the LASTNAME, FIRST NAME format?
rp_reverse = 'NO'

########################################################################################################################
#%%
df = pd.read_excel(df_filepath)
rp = pd.read_excel(rp_filepath)

#%% Pull out the filenames for the info tab
df_filename = df_filepath.split("/")[-1]
rp_filename = rp_filepath.split("/")[-1]

# %% Clean the DF column
df = df.drop_duplicates(subset=[dfc['LOAN_NAME']])
# Extract the string from our comparison column and add "_BASE"
df_base = str(dfc['LOAN_NAME']) + "_BASE"
# Clean the loan column
df[dfc['LOAN_NAME']] = df[dfc['LOAN_NAME']].astype(str)
df[dfc['LOAN_NAME']] = df[dfc['LOAN_NAME']].str.upper()
df[dfc['LOAN_NAME']] = df[dfc['LOAN_NAME']].str.strip()
# Create the BASE column off the cleaned column
df[df_base] = df[dfc['LOAN_NAME']].str.replace(r'\.', ' ', regex=True) # remove periods
df[df_base] = df[df_base].str.replace(r"\'", "", regex=True) # remove apostrophes
# Remove slashes and numbers
df[df_base] = df[df_base].str.replace(r"\/", " ", regex=True) # remove slashes
df[df_base] = df[df_base].str.replace(r"\/", " ", regex=True) # remove slashes
# Remove anything within parenthesis # todo

df[df_base] = df[df_base].str.replace(r"\d+", " ", regex=True) # remove numbers
df[df_base] = df[df_base].str.replace(r"\-", " ", regex=True) # remove hyphens
# replace 2 or more spaces with one space
df[df_base] = df[df_base].apply(lambda x: " ".join(x.split()))

#%%
if df_reverse == 'YES': # if df_reverse selected as yes, reverse the order of the string at the comma
    df[df_base] = df[df_base].apply(lambda x: ' '.join(reversed(x.split(', '))))
    info_df_reverse = 'The ' + dfc['LOAN_NAME'] + ' column was in LAST, FIRST order which was changed to FIRST LAST'
else: # if df_reverse was not selected as yes, remove the comma
    df[df_base] = df[df_base].str.replace(r'\,', '', regex=True)
    info_df_reverse = 'The ' + dfc['LOAN_NAME'] + ' column was in FIRST LAST order which was not changed'

#%% ignore words #todo add more drop_words from cleanco package
# todo "401 (K)" not removed because two words and numbers removed above... do this before? Dealt with in () above
drop_words = ['FOUNDATION', 'HOLDINGS', 'MANAGEMENT', 'INVESTMENTS', 'PROPERTIES', 'INTERNATIONAL', 'THE', '401K',
             'PARTNERSHIP', 'LIMITED', 'ENTERPRISES', 'ASSOCIATES', 'PARTNERS', 'INVESTMENT', 'GROUP', 'COMPANY',
             'ASSOCIATION', '401 (K)', '401(K)', 'LLC', 'HOLDING', 'INVESTORS', 'INC', '-', 'AND', '&', 'PLLC']

# Words which you read nothing after when encountered in a string
stop_words = ['DTD', 'DATED']

# dodge # ford
car_brands = ['ACURA', 'AUDI', 'BMW', 'BUICK', 'CADILLAC', 'CHEVROLET', 'CHEVY', 'CHRYSLER', 'FIAT', 'HONDA', 'HYUNDAI',
              'JAGUAR', 'JEEP', 'KIA', 'LAND ROVER', 'LEXUS', 'MAZDA', 'MERCEDES BENZ', 'MITSUBISHI', 'NISSAN',
              'PONTIAC', 'PORSCHE', 'SATURN', 'SUBARU', 'SUZUKI', 'TESLA', 'TOYOTA', 'VOLKSWAGEN', 'VOLVO']

# Drop the drop_words
df[df_base] = df[df_base].apply(lambda x: ' '.join([word for word in x.split() if word not in drop_words]))
df[df_base] = df[df_base].apply(lambda x: ' '.join([word for word in x.split() if word not in car_brands]))



#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
df[df_base] = df[df_base].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)
#%% Stop reading when you encounter this string. Cannot take multiple arguments, must run twice
for word in stop_words:
    df[df_base] = df[df_base].str.partition(word)[0]

#%% Clean the RP column
rp = rp.dropna(subset=[rpc['RP_NAME']])
rp = rp.drop_duplicates(subset=[rpc['RP_NAME']])
# Extract the string from our comparison column and add "_BASE"
rp_base = str(rpc['RP_NAME']) + "_BASE"

# Clean the RP Columns
rp[rpc['RP_NAME']] = rp[rpc['RP_NAME']].astype(str)
rp[rpc['RP_NAME']] = rp[rpc['RP_NAME']].str.upper()
rp[rpc['RP_NAME']] = rp[rpc['RP_NAME']].str.strip()
# Create the base column off the clean RP NAME column
rp[rp_base] = rp[rpc['RP_NAME']].str.replace(r'\.', ' ', regex=True) # Remove periods
rp[rp_base] = rp[rp_base].str.replace(r"\'", "", regex=True) # Remove apostrophes
# Remove slashes and numbers
rp[rp_base] = rp[rp_base].str.replace(r"\/", " ", regex=True) # Remove slashes
rp[rp_base] = rp[rp_base].str.replace(r"\d+", " ", regex=True) # Remove numbers
rp[rp_base] = rp[rp_base].str.replace(r"\-", " ", regex=True) # Remove hyphens
# replace 2 or more spaces with one space
rp[rp_base] = rp[rp_base].apply(lambda x: " ".join(x.split()))

#%%
if rp_reverse == 'YES': # if rp_reverse selected as yes, reverse the order of the string at the comma
    rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join(reversed(x.split(', ')))) # reverse the order at comma
    info_rp_reverse = 'The ' + rpc['RP_NAME'] + ' column was in LAST, FIRST order which was changed to FIRST LAST'
else: #if rp_reverse was not selected as yes, remove the comma
    rp[rp_base] = rp[rp_base].str.replace(r'\,', '', regex=True) # Remove commas
    info_rp_reverse = 'The ' + rpc['RP_NAME'] + ' column was in FIRST LAST order which was not changed'

#%% stop words
# Drop the drop_words
rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join([word for word in x.split() if word not in drop_words]))
rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join([word for word in x.split() if word not in car_brands]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

#%% Stop reading when you encounter this string. Cannot take multiple arguments, must run twice
for word in stop_words:
    rp[rp_base] = rp[rp_base].str.partition(word)[0]

#%% do it again after cleaning to see what we've missed
common_words_2 = pd.DataFrame(Counter(" ".join(rp[rp_base]).split()).most_common(200))
common_words_2.columns=['Word','Count']

#%% keep only the columns of interest
rp_short = rp[[rpc['RP_NAME'], rp_base]]
df_short = df[[dfc['LOAN_NAME'], df_base]]
# Create a cross product
cross_df = df_short.merge(rp_short, how='cross')

#%% TODO can get rid of one here (copy)
matches = cross_df.copy()
matches['RATIO_BASE'] = matches.apply(lambda x: fuzz.ratio(x[df_base], x[rp_base]), axis=1)

#%%
matches = matches[matches['RATIO_BASE']>=66]
matches['RATIO_ORDER'] = matches.apply(lambda x: fuzz.token_sort_ratio(x[df_base], x[rp_base]), axis=1)
matches['RATIO_FULL'] = matches.apply(lambda x: fuzz.ratio(x[dfc['LOAN_NAME']], x[rpc['RP_NAME']]), axis=1)
matches['JARO_BASE'] = matches.apply(lambda x: round((jaro.jaro_metric(x[df_base], x[rp_base]))*100), axis=1)
matches['JARO_FULL'] = matches.apply(lambda x: round((jaro.jaro_metric(x[dfc['LOAN_NAME']], x[rpc['RP_NAME']]))*100), axis=1)
matches['JARO_WINKLER_BASE'] = matches.apply(lambda x: round((jaro.jaro_winkler_metric(x[df_base], x[rp_base]))*100), axis=1)

matches = matches.dropna()
#matches = matches[matches['JARO_WINKLER_BASE']>=80]
matches = matches[[rpc['RP_NAME'],dfc['LOAN_NAME'], df_base, rp_base, 'RATIO_BASE', 'RATIO_ORDER', 'RATIO_FULL', 'JARO_BASE', 'JARO_FULL', 'JARO_WINKLER_BASE']]
matches = matches.sort_values(by=['RATIO_BASE', 'RATIO_ORDER'], ascending=(False, False))
#%%
InfoDict = {'Sources': ['Names in the ' + dfc['LOAN_NAME'] + ' column of the ' + df_filename +
                        ' were compared against names in the  ' + rpc['RP_NAME'] + ' column of the ' + rp_filename],
    'Drop Words': [str(drop_words)],
    'Loan Order': [info_df_reverse],
    'Related Parties Order': [info_rp_reverse]}
Info = pd.DataFrame.from_dict(InfoDict, orient='index')

perfect_base_matches = matches[matches['RATIO_BASE'] == 100]
perfect_full_matches = matches[matches['RATIO_FULL'] == 100]
likely_matches = matches[(90 <= matches['RATIO_BASE']) & (matches['RATIO_BASE'] < 100)]
medium_confidence = matches[(80 <= matches['RATIO_BASE']) & (matches['RATIO_BASE'] < 90)]
low_confidence = matches[(60 <= matches['RATIO_BASE']) & (matches['RATIO_BASE'] < 80)]

#%%
# EXPORT
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(client_name + '_RelatedParties.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.

Info.to_excel(writer, sheet_name='Information', header=False)

perfect_base_matches.to_excel(writer, sheet_name='Perfect Base Matches', index=False)
perfect_full_matches.to_excel(writer, sheet_name='Perfect Full Matches', index=False)
likely_matches.to_excel(writer, sheet_name='Likely Matches', index=False)
medium_confidence.to_excel(writer, sheet_name='Medium Confidence Matches', index=False)
low_confidence.to_excel(writer, sheet_name='Low Confidence Matches', index=False)

matches.to_excel(writer, sheet_name='All Matches',index=False)

# Get the xlsxwriter objects from the dataframe writer object.
workbook = writer.book
info_worksheet = writer.sheets['Information']

# Add some cell formats.

formatbold = workbook.add_format({'bold': True})
format_wrap = workbook.add_format({'text_wrap': True})
under_border_format = workbook.add_format({'bottom': 1})
excel_format = workbook.add_format({'bg_color': '#D0E2C5', 'border': True})
color1_format = workbook.add_format({'bg_color': '#b0c4de'})
color1_light_format = workbook.add_format({'bg_color': '#dbe4f0'})
color2_format = workbook.add_format({'bg_color': '#b1cbbb'})
color2_light_format = workbook.add_format({'bg_color': 'e0ebe4'})
ratio_format = workbook.add_format({'bg_color': '#ebebeb'})

# Set column width for Information Sheet
info_worksheet.set_column('A:A', 26)
info_worksheet.set_column('B:B', 100, format_wrap)

# Set the column width and format
def set_match_worksheet_format(match_worksheet_name):
    match_worksheet = writer.sheets[match_worksheet_name]

    match_worksheet.freeze_panes(1, 1)
    match_worksheet.set_column('A:A', 30, color2_light_format)
    match_worksheet.set_column('B:B', 30, color1_light_format)
    match_worksheet.set_column('C:C', 25, color1_light_format)
    match_worksheet.set_column('D:D', 25, color2_light_format)
    match_worksheet.set_column('E:I', 13, ratio_format)
    match_worksheet.set_column('J:J', 19, ratio_format)

set_match_worksheet_format('Perfect Base Matches')
set_match_worksheet_format('Perfect Full Matches')
set_match_worksheet_format('Likely Matches')
set_match_worksheet_format('Medium Confidence Matches')
set_match_worksheet_format('Low Confidence Matches')
set_match_worksheet_format('All Matches')


# Close the Pandas Excel writer and output the Excel file.
writer.save()

#%%
def convert(seconds):
    seconds = seconds % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60

    return "%d:%02d:%02d" % (hour, minutes, seconds)


endtime = datetime.datetime.now()
print("End: ", endtime)
run_time = (endtime-starttime)
run_seconds = run_time.total_seconds()
print("Runtime (HH:MM:SS): ", convert(run_seconds))
