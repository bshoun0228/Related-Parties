# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
from collections import Counter
from fuzzywuzzy import fuzz
import numpy as np
import datetime
import jaro
#%%
starttime = datetime.datetime.now()
print("Start: ", starttime)


#%% ####################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = 'RP_OTHER_COL_MANY_PACKAGES'
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

# TODO
# Remove middle initials?
# Drop all single letters (middle initials)
# dhg['Base_Name']= dhg['Base_Name'].str.replace(r'\b\w\b', '').str.replace(r'\s+', ' ')
# bkd['Base_Name']= bkd['Base_Name'].str.replace(r'\b\w\b', '').str.replace(r'\s+', ' ')


########################################################################################################################
#%%
df = pd.read_excel(df_filepath)
rp = pd.read_excel(rp_filepath)

#%% Pull out the filenames for the info tab
df_filename = df_filepath.split("/")[-1]
rp_filename = rp_filepath.split("/")[-1]


# %%

df = df.drop_duplicates(subset=[dfc['LOAN_NAME']])
# Extract the string from our comparison column and add "_BASE"
df_base = str(dfc['LOAN_NAME']) + "_BASE"
# Clean the loan column
df[dfc['LOAN_NAME']] = df[dfc['LOAN_NAME']].astype(str)
df[dfc['LOAN_NAME']] = df[dfc['LOAN_NAME']].str.upper()
df[dfc['LOAN_NAME']] = df[dfc['LOAN_NAME']].str.strip()
# Create the BASE column off the cleaned column
df[df_base] = df[dfc['LOAN_NAME']].str.replace(r'\.', '', regex=True) # remove periods
df[df_base] = df[df_base].str.replace(r"\'S", "", regex=True) # remove 'S
# Remove slashes and numbers
df[df_base] = df[df_base].str.replace(r"\/", "", regex=True) # remove slashes
df[df_base] = df[df_base].str.replace(r"\d+", "", regex=True) # remove numbers
# TODO we want to remove the LLC/STOPwords before moving this around? Does it matter?
if df_reverse == 'YES':
    df[df_base] = df[df_base].apply(lambda x: ' '.join(reversed(x.split(', '))))
    info_df_reverse = 'The ' + dfc['LOAN_NAME'] + ' column was in LAST, FIRST order which was changed to FIRST LAST'
else:
    df[df_base] = df[df_base].str.replace(r'\,', '', regex=True)
    info_df_reverse = 'The ' + dfc['LOAN_NAME'] + ' column was in FIRST LAST order which was not changed'
#%% stop words #todo add more stopwords from cleanco package
stopwords = ['FOUNDATION', 'HOLDINGS', 'MANAGEMENT', 'INVESTMENTS', 'PROPERTIES', 'INTERNATIONAL', 'THE', '401K',
             'PARTNERSHIP', 'LIMITED', 'ENTERPRISES', 'ASSOCIATES', 'PARTNERS', 'INVESTMENT', 'GROUP', 'COMPANY',
             'ASSOCIATION', '401 (K)', '401(K)', 'LLC', 'HOLDING', 'INVESTORS', 'INC', '-', 'AND', '&', 'PLLC', 'DTD']

# Drop the stopwords
df[df_base] = df[df_base].apply(lambda x: ' '.join([word for word in x.split() if word not in stopwords]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
df[df_base] = df[df_base].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

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
rp[rp_base] = rp[rpc['RP_NAME']].str.replace(r'\.', '', regex=True) # Remove periods
rp[rp_base] = rp[rp_base].str.replace(r"\'S", "", regex=True) # Remove 'S
# Remove slashes and numbers
rp[rp_base] = rp[rp_base].str.replace(r"\/", "", regex=True) # Remove slashes
rp[rp_base] = rp[rp_base].str.replace(r"\d+", "", regex=True) # Remove numbers
if rp_reverse == 'YES':
    rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join(reversed(x.split(', ')))) # reverse the order at comma
    info_rp_reverse = 'The ' + rpc['RP_NAME'] + ' column was in LAST, FIRST order which was changed to FIRST LAST'
else:
    rp[rp_base] = rp[rp_base].str.replace(r'\,', '', regex=True) # Remove commas
    info_rp_reverse = 'The ' + rpc['RP_NAME'] + ' column was in FIRST LAST order which was not changed'

#%% stop words
# Drop the stopwords
rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join([word for word in x.split() if word not in stopwords]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
rp[rp_base] = rp[rp_base].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

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
matches['LEV_RATIO'] = matches.apply(lambda x: fuzz.ratio(x[df_base], x[rp_base]), axis=1)

#%%
matches = matches[matches['LEV_RATIO']>=66]
# Irrespective of order
matches['LEV_TOKEN_SORT'] = matches.apply(lambda x: fuzz.token_sort_ratio(x[df_base], x[rp_base]), axis=1)
# takes out common tokens
matches['LEV_TOKEN_SET'] = matches.apply(lambda x: fuzz.token_set_ratio(x[df_base], x[rp_base]), axis=1)
matches['LEV_PARTIAL'] = matches.apply(lambda x: fuzz.partial_ratio(x[df_base], x[rp_base]), axis=1)
matches['LEV_PARTIAL_SORT'] = matches.apply(lambda x: fuzz.partial_token_sort_ratio(x[df_base], x[rp_base]), axis=1)
matches['LEV_PARTIAL_SET'] = matches.apply(lambda x: fuzz.partial_token_set_ratio(x[df_base], x[rp_base]), axis=1)
matches['LEV_Q_RATIO'] = matches.apply(lambda x: fuzz.QRatio(x[df_base], x[rp_base]), axis=1)
matches['LEV_UQ_RATIO'] = matches.apply(lambda x: fuzz.UQRatio(x[df_base], x[rp_base]), axis=1)
matches['LEV_UW_RATIO'] = matches.apply(lambda x: fuzz.UWRatio(x[df_base], x[rp_base]), axis=1)
matches['LEV_W_RATIO'] = matches.apply(lambda x: fuzz.WRatio(x[df_base], x[rp_base]), axis=1)

matches['JARO_METRIC'] = matches.apply(lambda x: round((jaro.jaro_metric(x[df_base], x[rp_base]))*100), axis=1)
matches['JARO_WINKLER'] = matches.apply(lambda x: round((jaro.jaro_winkler_metric(x[df_base], x[rp_base]))*100), axis=1)
matches['JARO_ORIGINAL_METRIC'] = matches.apply(lambda x: round((jaro.original_metric(x[df_base], x[rp_base]))*100), axis=1)
#matches['CUSTOM_METRIC'] = matches.apply(lambda x: round((jaro.custom_metric(x[df_base], x[rp_base]))*100), axis=1)
#%%
matches = matches.dropna()
#matches = matches[matches['JARO_WINKLER_BASE']>=80]
matches = matches[[rpc['RP_NAME'],dfc['LOAN_NAME'], df_base, rp_base, 'LEV_RATIO', 'LEV_TOKEN_SORT', 'LEV_TOKEN_SET',
                   'LEV_PARTIAL', 'LEV_PARTIAL_SORT', 'LEV_PARTIAL_SET', 'LEV_Q_RATIO', 'LEV_UQ_RATIO', 'LEV_UW_RATIO',
                   'LEV_W_RATIO', 'JARO_METRIC', 'JARO_WINKLER', 'JARO_ORIGINAL_METRIC']]
matches = matches.sort_values(by=['LEV_RATIO', 'LEV_TOKEN_SORT'], ascending=(False, False))
#%%
InfoDict = {'Sources': ['Names in the ' + dfc['LOAN_NAME'] + ' column of the ' + df_filename +
                        ' were compared against names in the  ' + rpc['RP_NAME'] + ' column of the ' + rp_filename],
    'Drop Words': [str(stopwords)],
    'Loan Order': [info_df_reverse],
    'Related Parties Order': [info_rp_reverse]}
Info = pd.DataFrame.from_dict(InfoDict, orient='index')

#%%
# EXPORT
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(client_name + '_RelatedParties.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.

Info.to_excel(writer, sheet_name='Information', header=False)
matches.to_excel(writer, sheet_name='Matches',index=False)

# Get the xlsxwriter objects from the dataframe writer object.
workbook = writer.book
info_worksheet = writer.sheets['Information']
match_worksheet = writer.sheets['Matches']

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
match_worksheet.freeze_panes(1, 1)
match_worksheet.set_column('A:A', 30, color2_light_format)
match_worksheet.set_column('B:B', 30, color1_light_format)
match_worksheet.set_column('C:C', 25, color1_light_format)
match_worksheet.set_column('D:D', 25, color2_light_format)
match_worksheet.set_column('E:I', 13, ratio_format)
match_worksheet.set_column('J:J', 19, ratio_format)


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
