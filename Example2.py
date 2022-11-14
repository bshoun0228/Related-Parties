# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
from collections import Counter
from fuzzywuzzy import fuzz
import numpy as np
import datetime
#%%
starttime = datetime.datetime.now()
print("Start: ", starttime)
#%% READ IN THE DATA
df = pd.read_excel('Data/Loan_Name.xlsx')
rp = pd.read_excel('Data/Related Parties Clean.xlsx')

# %% Clean the loan column
df['LOAN_BASE'] = df['LOAN_NAME'].astype(str)

df['LOAN_BASE'] = df['LOAN_BASE'].str.upper()
df['LOAN_BASE'] = df['LOAN_BASE'].str.strip()
df['LOAN_BASE'] = df['LOAN_BASE'].str.replace(r'\,', '', regex=True)
df['LOAN_BASE'] = df['LOAN_BASE'].str.replace(r'\.', '', regex=True)
df['LOAN_BASE'] = df['LOAN_BASE'].str.replace(r"\'S", "", regex=True)

#%% stop words
stopwords = ['FOUNDATION', 'HOLDINGS', 'MANAGEMENT', 'INVESTMENTS', 'PROPERTIES', 'INTERNATIONAL', 'THE', '401K',
             'PARTNERSHIP', 'LIMITED', 'ENTERPRISES', 'ASSOCIATES', 'PARTNERS', 'INVESTMENT', 'GROUP', 'COMPANY',
             'ASSOCIATION', '401 (K)', '401(K)', 'LLC', 'HOLDING', 'INVESTORS', 'INC', '-', 'AND', '&', 'PLLC']

# Drop the drop_words
df['LOAN_BASE'] = df['LOAN_BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in stopwords]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
df['LOAN_BASE'] = df['LOAN_BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

#%% do it again after cleaning to see what we've missed
common_words = pd.DataFrame(Counter(" ".join(df['LOAN_BASE']).split()).most_common(200))
common_words.columns=['Word','Count']

#%% Clean the RP column
rp = rp.dropna(subset=['BUS'])
rp['BUS_BASE'] = rp['BUS'].astype(str)
rp['BUS_BASE'] = rp['BUS_BASE'].str.upper()
rp['BUS_BASE'] = rp['BUS_BASE'].str.strip()
rp['BUS_BASE'] = rp['BUS_BASE'].str.replace(r'\,', '', regex=True)
rp['BUS_BASE'] = rp['BUS_BASE'].str.replace(r'\.', '', regex=True)
rp['BUS_BASE'] = rp['BUS_BASE'].str.replace(r"\'S", "", regex=True)

#%% stop words
# Drop the drop_words
rp['BUS_BASE'] = rp['BUS_BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in stopwords]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
rp['BUS_BASE'] = rp['BUS_BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

#%% do it again after cleaning to see what we've missed
common_words_2 = pd.DataFrame(Counter(" ".join(rp['BUS_BASE']).split()).most_common(200))
common_words_2.columns=['Word','Count']

#%%
rp_bus = rp[['BUS', 'BUS_BASE']]
rp_bus = rp_bus.rename(columns={'BUS': 'RP_BUS'})
cross_df = df.merge(rp_bus, how='cross')

#cross_df = bkd.merge(dhg_AU, how='cross')
#%%

matches = cross_df.copy()
matches['RATIO_BASE'] = matches.apply(lambda x: fuzz.ratio(x['LOAN_BASE'], x['BUS_BASE']), axis=1)

#%%
matches = matches[matches['RATIO_BASE']>=60]
matches['RATIO_ORDER'] = matches.apply(lambda x: fuzz.token_sort_ratio(x['LOAN_BASE'], x['BUS_BASE']), axis=1)
matches['RATIO_FULL'] = matches.apply(lambda x: fuzz.ratio(x['LOAN_BASE'], x['RP_BUS']), axis=1)

matches = matches[['RP_BUS','LOAN_NAME', 'LOAN_BASE', 'BUS_BASE', 'RATIO_BASE', 'RATIO_ORDER', 'RATIO_FULL']]
matches = matches.sort_values(by=['RATIO_BASE', 'RATIO_ORDER'], ascending=(False, False))
#%%
InfoDict = {'Description': ['Describe the columns here'],
    'Drop Words': [str(stopwords)]}
Info = pd.DataFrame.from_dict(InfoDict, orient='index')

#%%
# EXPORT
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('RelatedPartiesMatchingExample2.xlsx', engine='xlsxwriter')

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
match_worksheet.set_column('E:G', 13, ratio_format)


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
