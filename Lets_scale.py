# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
from collections import Counter
from fuzzywuzzy import fuzz
import numpy as np
import datetime

starttime = datetime.datetime.now()
print("Start: ", starttime)

#%% ####################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = 'NAME_Demo'
# Put the filepath to the GL/Other data
ln_df_filepath = 'Data/Loan_Name.xlsx'

# What column are we comparing?
lnc = {'LOAN_NAME': 'LOAN_NAME'}
# IF the column in the LASTNAME, FIRST NAME format, type 'YES'
ln_reverse = 'YES'

# Put the filepath to the related parties
rp_filepath = 'Data/Related Parties Clean.xlsx'

# What column are we comparing?
rpc = {'RP_NAME': 'NAME'}
# Is this column in the LASTNAME, FIRST NAME format?
rp_reverse = 'NO'

########################################################################################################################
#%%
ln_df = pd.read_excel(ln_df_filepath)
rp_df = pd.read_excel(rp_filepath)

#%% Pull out the filenames for the info tab
ln_df_filename = ln_df_filepath.split("/")[-1]
rp_filename = rp_filepath.split("/")[-1]

# %% ignore words #todo add more drop_words from cleanco package
# todo "401 (K)" not removed because two words and numbers removed above... do this before? Dealt with in () above
drop_words = ['FOUNDATION', 'HOLDINGS', 'MANAGEMENT', 'INVESTMENTS', 'PROPERTIES', 'INTERNATIONAL', 'THE', '401K',
              'PARTNERSHIP', 'LIMITED', 'ENTERPRISES', 'ASSOCIATES', 'PARTNERS', 'INVESTMENT', 'GROUP', 'COMPANY',
              'ASSOCIATION', '401 (K)', '401(K)', 'LLC', 'HOLDING', 'INVESTORS', '-', 'AND', '&',
              'IRREVOCABLE', 'REVOCABLE', 'DESCENDANTS', 'TRUST', 'COMPANY', 'INCORPORATED', 'CORPORATION', 'CORP',
              'INC', 'LTDA', 'UNLTD', 'LTD', 'PLLC', 'LLC', 'LLP', 'LLLP', 'LP']


# Words which you read nothing after when encountered in a string
stop_words = ['DTD', 'DATED']

# dodge # ford
car_brands = ['ACURA', 'AUDI', 'BMW', 'BUICK', 'CADILLAC', 'CHEVROLET', 'CHEVY', 'CHRYSLER', 'FIAT', 'GMC', 'HONDA',
              'HYUNDAI', 'JAGUAR', 'JEEP', 'KIA', 'LAND ROVER', 'LEXUS', 'MAZDA', 'MERCEDES BENZ', 'MITSUBISHI',
              'NISSAN', 'PONTIAC', 'PORSCHE', 'SATURN', 'SUBARU', 'SUZUKI', 'TESLA', 'TOYOTA', 'VOLKSWAGEN', 'VOLVO']

# %% Clean the DF column
def make_base_names(df, original_col, base_col, reverse):
    # Clean the loan column
    df[original_col] = df[original_col].astype(str)
    df[original_col] = df[original_col].str.upper()
    df[original_col] = df[original_col].str.strip()
    # Create the BASE column off the cleaned column
    df[base_col] = df[original_col].str.replace(r'\.', ' ', regex=True) # remove periods
    df[base_col] = df[base_col].str.replace(r"\'", "", regex=True) # remove apostrophes
    # Remove anything within parenthesis # todo
    df[base_col] = df[base_col].str.replace(r"\([^)]*\)", " ", regex=True) # remove slashes
    # Remove slashes, hyphens, and numbers
    df[base_col] = df[base_col].str.replace(r"\/", " ", regex=True) # remove slashes
    df[base_col] = df[base_col].str.replace(r"\-", " ", regex=True) # remove hyphens
    df[base_col] = df[base_col].str.replace(r"\d+", " ", regex=True) # remove numbers
    # replace 2 or more spaces with one space
    df[base_col] = df[base_col].apply(lambda x: " ".join(x.split()))

    if reverse == 'YES': # if df_reverse selected as yes, reverse the order of the string at the comma
        df[base_col] = df[base_col].apply(lambda x: ' '.join(reversed(x.split(', '))))
    else: # if df_reverse was not selected as yes, remove the comma
        df[base_col] = df[base_col].str.replace(r'\,', '', regex=True)

    # Drop the drop_words
    df[base_col] = df[base_col].apply(lambda x: ' '.join([word for word in x.split() if word not in drop_words]))
    df[base_col] = df[base_col].apply(lambda x: ' '.join([word for word in x.split() if word not in car_brands]))

    # Drop only ESTATE if it is not part of REAL ESTATE
    reallist = ['ESTATE']
    df[base_col] = df[base_col].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

    # Stop reading when you encounter this string. Cannot take multiple arguments, so must be looped
    for word in stop_words:
        df[base_col] = df[base_col].str.partition(word)[0]

    return df

#%% Clean the DF column
ln_df = ln_df.drop_duplicates(subset=[lnc['LOAN_NAME']])
# Extract the string from our comparison column and add "_BASE"
ln_base = str(lnc['LOAN_NAME']) + "_BASE"
ln_df = make_base_names(ln_df, lnc['LOAN_NAME'], ln_base, ln_reverse)

#%% Clean the RP column
rp_df = rp_df.dropna(subset=[rpc['RP_NAME']])
rp_df = rp_df.drop_duplicates(subset=[rpc['RP_NAME']])
# Extract the string from our comparison column and add "_BASE"
rp_base = str(rpc['RP_NAME']) + "_BASE"
rp_df = make_base_names(rp_df, rpc['RP_NAME'], rp_base, rp_reverse)

#%% do it again after cleaning to see what we've missed
common_words_2 = pd.DataFrame(Counter(" ".join(rp_df[rp_base]).split()).most_common(200))
common_words_2.columns=['Word','Count']

#%% keep only the columns of interest
rp_short = rp_df[[rpc['RP_NAME'], rp_base]]
ln_df_short = ln_df[[lnc['LOAN_NAME'], ln_base]]
# Create a cross product
cross_df = ln_df_short.merge(rp_short, how='cross')

#%% TODO can get rid of one here (copy)
matches = cross_df.copy()
matches['RATIO_BASE'] = matches.apply(lambda x: fuzz.ratio(x[ln_base], x[rp_base]), axis=1)

#%%
matches = matches[matches['RATIO_BASE']>=66]
#%%
matches['RATIO_ORDER'] = matches.apply(lambda x: fuzz.token_sort_ratio(x[ln_base], x[rp_base]), axis=1)
matches['RATIO_FULL'] = matches.apply(lambda x: fuzz.ratio(x[lnc['LOAN_NAME']], x[rpc['RP_NAME']]), axis=1)

matches = matches.dropna()

matches = matches[[rpc['RP_NAME'],lnc['LOAN_NAME'], ln_base, rp_base, 'RATIO_BASE', 'RATIO_ORDER', 'RATIO_FULL']]
matches = matches.sort_values(by=['RATIO_BASE', 'RATIO_ORDER'], ascending=(False, False))

#%%
if ln_reverse == 'YES': # if df_reverse selected as yes, reverse the order of the string at the comma
    info_df_reverse = 'The ' + lnc['LOAN_NAME'] + ' column was in LAST, FIRST order which was changed to FIRST LAST'
else: # if df_reverse was not selected as yes, remove the comma
    info_df_reverse = 'The ' + lnc['LOAN_NAME'] + ' column was in FIRST LAST order which was not changed'

if rp_reverse == 'YES':
    info_rp_reverse = 'The ' + rpc['RP_NAME'] + ' column was in LAST, FIRST order which was changed to FIRST LAST'
else:
    info_rp_reverse = 'The ' + rpc['RP_NAME'] + ' column was in FIRST LAST order which was not changed'

#%%
low_score_threshold = 60
medium_score_threshold = 80
high_score_threshold = 90
perfect_score_threshold = 100

InfoDict = [
    ['Alterations Made to the Original Name: (fields end in “_FULL”)', ''],
        ['Removed punctuation',
            'Periods, apostrophes, slashes, hyphens, numbers, and extra spaces potentially caused by punctuation removal'],
        ['Format changes',
            'Forced to all caps'],
    ['Modified Name Used For Name Matching: (fields end in “_BASE”)', ''],
        ['Removed punctuation',
            'Periods, apostrophes, slashes, hyphens, numbers, commas*, and extra spaces potentially caused by punctuation removal'],
        ['Format changes',
            'Forced to all caps; if manual review of raw data shows names are in last, first name order, names were switched to first name last name form (*commas were removed after this)'],
        ['Name changes',
            'Drop Words: Words that were stripped from the name to get a more true name assessment score not influenced by these common words'],
            ['', '     Common Words: '+ str(drop_words)],
            ['', '     Car Names: ' + str(car_brands)],
            ['', 'Stop Words: Only the portion of the name that appeared before these words was kept (E.g. “Gordon Foods Dated: 12/2/2018” returns as “Gordon Foods”)'],
            ['', '     ' + str(stop_words)],
            ['', "Special Instances: Estate (E.g. Caroline Real Estate vs Caroline Estate)"],
            ['', "     ‘Estate’ is removed unless it is part of ‘Real Estate’"],
    ['Output fields',
        'RATIO_BASE: Levenshtein score of modified loan name compared to modified related party name'],
        ['', 'RATIO_ORDER: all letters in a name are rearranged into alphabetical order and then compared to the rearranged letters of the other column (Loan name compared to Related Party); a score is then calculated based on how many changes need to be made for them to match'],
        ['', 'RATIO_FULL: Lev score of the original names with minor modifications (all caps, removed extra spaces, and all punctuation except for commas)'],
    ['Process Summary',
        'Take original names and modify slightly'],
        ['', 'Get the "base names" from the original names, as described above, in order to more accurately represent each name'],
        ['', 'Calculate scores based on the base names using the levenshtein ratio'],
        ['', 'Ignore records that receive a RATIO_BASE lower than ' + str(low_score_threshold)],
        ['', 'Calculate FULL scores using the only slightly modified original names'],
        ['', 'Sort the data by ascending RATIO_BASE and RATIO_ORDER scores'],
    ['Sheets',
        'Perfect Base Matches: RATIO_BASE = ' + str(perfect_score_threshold)],
        ['', 'Perfect Full Matches: RATIO_FULL = ' + str(perfect_score_threshold)],
        ['', 'Likely Matches: RATIO_BASE >= ' + str(high_score_threshold) + ' and < ' + str(perfect_score_threshold)],
        ['', 'Medium Confidence Matches: RATIO_BASE >= ' + str(medium_score_threshold) + ' and < ' + str(high_score_threshold)],
        ['', 'Low Confidence Matches: RATIO_BASE >= ' + str(low_score_threshold) + ' and < ' + str(medium_score_threshold)]
]

#%%
Info = pd.DataFrame(InfoDict, columns=[' ', 'Details'])
#%%

perfect_base_matches = matches[matches['RATIO_BASE'] == perfect_score_threshold]
perfect_full_matches = matches[matches['RATIO_FULL'] == perfect_score_threshold]
likely_matches = matches[(high_score_threshold <= matches['RATIO_BASE'])
                         & (matches['RATIO_BASE'] < perfect_score_threshold)]
medium_confidence = matches[(medium_score_threshold <= matches['RATIO_BASE'])
                            & (matches['RATIO_BASE'] < high_score_threshold)]
low_confidence = matches[(low_score_threshold <= matches['RATIO_BASE'])
                         & (matches['RATIO_BASE'] < medium_score_threshold)]

#%%
# EXPORT
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(client_name + '_RelatedParties.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.

Info.to_excel(writer, sheet_name='Information', index=False)

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
bottom_right_format = workbook.add_format({'bottom': 1, 'right': 1})
excel_format = workbook.add_format({'bg_color': '#D0E2C5', 'border': True})
right_format = workbook.add_format({'right': 1})
color1_format = workbook.add_format({'bg_color': '#b0c4de'})
color1_light_format = workbook.add_format({'bg_color': '#dbe4f0'})
color2_format = workbook.add_format({'bg_color': '#b1cbbb'})
color2_light_format = workbook.add_format({'bg_color': 'e0ebe4'})
ratio_format = workbook.add_format({'bg_color': '#ebebeb'})

# Set column width for Information Sheet
info_worksheet.set_column('A:A', 59)
info_worksheet.set_column('B:B', 100, format_wrap)
# With Row/Column notation you must specify all four cells in the range: (first_row, first_col, last_row, last_col)

info_worksheet.conditional_format(1, 0, 1, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(4, 0, 4, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(14, 0, 14, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(17, 0, 17, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(23, 0, 23, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})

info_worksheet.conditional_format(0, 0, 0, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(3, 0, 3, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(13, 0, 13, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(16, 0, 16, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(22, 0, 22, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(27, 0, 27, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})

info_worksheet.conditional_format(0, 1, 0, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(3, 1, 3, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(13, 1, 13, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(16, 1, 16, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(22, 1, 22, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(27, 1, 27, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})

info_worksheet.conditional_format(1, 1, 27, 1, {'type': 'formula', 'criteria': 'True', 'format': right_format})

# Set the column width and format
def set_match_worksheet_format(match_worksheet_name):
    match_worksheet = writer.sheets[match_worksheet_name]

    match_worksheet.freeze_panes(1, 1)
    match_worksheet.set_column('A:A', 30, color2_light_format)
    match_worksheet.set_column('B:B', 30, color1_light_format)
    match_worksheet.set_column('C:C', 25, color1_light_format)
    match_worksheet.set_column('D:D', 25, color2_light_format)
    match_worksheet.set_column('E:G', 13, ratio_format)


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

#%%
