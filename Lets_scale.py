# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
from collections import Counter
from fuzzywuzzy import fuzz
import numpy as np
import datetime
import os

starttime = datetime.datetime.now()
print("Start: ", starttime)

#%% ####################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = 'Newtab'  # what do you want files named?
export_path = os.path.expanduser(r"~\OneDrive - FORVIS, LLP\Related Parties Examples\Data")  # where you want results

# Put the filepath to the GL/Other data
ln_df_filepath = os.path.expanduser(r"~\OneDrive - FORVIS, LLP\Related Parties Examples\Data\Loan Flat File.xlsx")

# What column are we comparing? Enter the column headers for the Customer Name and Account columns
## If there is no Account column provided, type None
lnc = {'LOAN_NAME': 'Customer Name', 'ACCOUNTS': 'Account Number'}
# IF the column in the LASTNAME, FIRST NAME format, type 'YES' (CAPITAL)
ln_reverse = 'YES'

# Put the filepath to the related parties
rp_filepath = os.path.expanduser(r"~\OneDrive - FORVIS, LLP\Related Parties Examples\Data\Related Parties Clean.xlsx")

# What column are we comparing?
rpc={'RP_NAME': 'OTHER'}
# Is this column in the LASTNAME, FIRST NAME format?
rp_reverse = 'NO'

########################################################################################################################
#%%
logname = export_path + '\\' + client_name + " Related Parties Log.txt"

# Have to initiate logger after getting client name so that can export log with including client name
log = open(logname, "a+")
log.write("-----RELATED PARTIES ANALYSIS RUN: " + datetime.datetime.now().strftime('%d-%b-%y %H:%M:%S') + " -----\n\n")

#%% Read in the data
ln_df = pd.read_excel(ln_df_filepath, dtype=str)  # Read in the loan file
ln_df_raw = ln_df.copy()
ln_count = len(ln_df)
log.write("Loan number of rows read in from file: " + str(ln_count) + '\n')

rp_df = pd.read_excel(rp_filepath)  # Read in the related party file
rp_df_raw = rp_df.copy()
rp_count = len(rp_df)
log.write("Related Parties number of rows read in from file: " + str(rp_count) + '\n\n')
#%%
if lnc['ACCOUNTS'] == None:
    ln_df['ACCOUNTS'] = np.nan
    #lnc.update({'ACCOUNTS': 'ACCOUNTS'})

ln_df = ln_df.rename(columns={lnc['LOAN_NAME']: 'LOAN_NAME', lnc['ACCOUNTS']: 'ACCOUNTS'})  # Rename the column
ln_df = ln_df[['LOAN_NAME', 'ACCOUNTS']]  # only keep that colum
ln_df['ACCOUNTS'] = ln_df['ACCOUNTS'].astype(str) # defensive - make sure it is a string for comma separated list

rp_df = rp_df.rename(columns={rpc['RP_NAME']:'RELATED_PARTY_NAME'})  # Rename the column
rp_df = rp_df[['RELATED_PARTY_NAME']]  # Only keep that column

#%% Pull out the filenames for the info tab
ln_df_filename = ln_df_filepath.split("\\")[-1]  # Take everything after the last slash - the filename
rp_filename = rp_filepath.split("\\")[-1]

# %% ignore words #NOTE: Real estate is added here for the info tab but is not removed when this list is applied due to space
drop_words = ['FOUNDATION', 'HOLDINGS', 'MANAGEMENT', 'INVESTMENTS', 'PROPERTIES', 'INTERNATIONAL', 'THE', '401K',
              'PARTNERSHIP', 'LIMITED', 'ENTERPRISES', 'ASSOCIATES', 'PARTNERS', 'INVESTMENT', 'GROUP', 'COMPANY',
              'ASSOCIATION', '401 (K)', '401(K)', 'LLC', 'HOLDING', 'INVESTORS', '-', 'AND', '&',
              'IRREVOCABLE', 'REVOCABLE', 'DESCENDANTS', 'TRUST', 'COMPANY', 'INCORPORATED', 'CORPORATION', 'CORP',
              'INC', 'LTDA', 'UNLTD', 'LTD', 'PLLC', 'LLC', 'LLP', 'LLLP', 'LP', 'REAL ESTATE']

# Words which you read nothing after when encountered in a string
stop_words = ['DTD', 'DATED']

car_brands = ['ACURA', 'AUDI', 'BMW', 'BUICK', 'CADILLAC', 'CHEVROLET', 'CHEVY', 'CHRYSLER', 'FIAT', 'GMC', 'HONDA',
              'HYUNDAI', 'JAGUAR', 'JEEP', 'KIA', 'LAND ROVER', 'LEXUS', 'MAZDA', 'MERCEDES BENZ', 'MITSUBISHI',
              'NISSAN', 'PONTIAC', 'PORSCHE', 'SATURN', 'SUBARU', 'SUZUKI', 'TESLA', 'TOYOTA', 'VOLKSWAGEN', 'VOLVO']

# %% Clean the DF column
def make_base_names(df, original_col, base_col, reverse):
    # Clean the loan column
    df[original_col] = df[original_col].astype(str)
    df[original_col] = df[original_col].str.upper()
    df[original_col] = df[original_col].str.strip()
    # have to drop duplicates again after cleaning
    df = df.dropna(subset=[original_col])  # Drop blank rows
    #df = df.drop_duplicates(subset=[original_col, 'ACCOUNTS'])  # drop duplicates
    # Create the BASE column off the cleaned column
    df[base_col] = df[original_col].str.replace(r'\.', ' ', regex=True) # remove periods
    df[base_col] = df[base_col].str.replace(r"\'", "", regex=True) # remove apostrophes
    # Remove anything within parenthesis
    df[base_col] = df[base_col].str.replace(r"\([^)]*\)", " ", regex=True) # remove slashes
    # Remove slashes, hyphens, and numbers
    df[base_col] = df[base_col].str.replace(r"\/", " ", regex=True) # remove slashes
    df[base_col] = df[base_col].str.replace(r"\-", " ", regex=True) # remove hyphens
    df[base_col] = df[base_col].str.replace(r"\d+", " ", regex=True) # remove numbers
    df[base_col] = df[base_col].str.replace(r"REAL ESTATE", " ", regex=True)  # remove REAL ESTATE
    # replace 2 or more spaces with one space
    df[base_col] = df[base_col].apply(lambda x: " ".join(x.split()))

    if reverse == 'YES': # if df_reverse selected as yes, reverse the order of the string at the comma
        df[base_col] = df[base_col].apply(lambda x: ' '.join(reversed(x.split(', '))))
    else: # if df_reverse was not selected as yes, remove the comma
        df[base_col] = df[base_col].str.replace(r'\,', '', regex=True)

    # Drop the drop_words
    df[base_col] = df[base_col].apply(lambda x: ' '.join([word for word in x.split() if word not in drop_words]))
    df[base_col] = df[base_col].apply(lambda x: ' '.join([word for word in x.split() if word not in car_brands]))

    # Stop reading when you encounter this string. Cannot take multiple arguments, so must be looped
    for word in stop_words:
        df[base_col] = df[base_col].str.partition(word)[0]

    return df

#%%
ln_count_before = ln_count
# Apply the cleaning function BEFORE aggregation/drop duplicates (cleaning may make new duplicates
ln_df = make_base_names(ln_df, 'LOAN_NAME', 'LOAN_BASE', ln_reverse)  # apply the cleaning function

if lnc['ACCOUNTS'] == None:
    ln_df = ln_df.dropna(subset=['LOAN_NAME'])  # Drop blank rows
    ln_df = ln_df.drop_duplicates(subset=['LOAN_NAME'])  # drop duplicates
    log.write('No ACCOUNTS given, duplicates evaluated on LOAN_NAME \n\n')
else:
    # Defensive - strip in case of spaces
    ln_df['ACCOUNTS'] = ln_df['ACCOUNTS'].str.strip()
    # Get a list of all account numbers for each name
    ln_df = ln_df.groupby(['LOAN_NAME', 'LOAN_BASE']).agg({'ACCOUNTS': ', '.join}).reset_index()
    log.write('ACCOUNTS provided, accounts for each duplicative name aggregated \n\n')

#%% Clean the DF column


ln_count = len(ln_df)  # have to do this AFTER make_base_names because new duplicates can be created with cleaning

ln_diff = ln_count_before-ln_count
log.write(str(ln_diff) + " empty or duplicative LOAN_NAME instances found \n")
log.write(str(ln_count) + " unique non-null LOAN_NAMES for analysis\n")
if ln_count_before-ln_diff == ln_count:
    log.write(str(ln_count_before) + " correctly equals " + str(ln_count) + " + " + str(ln_diff) + "\n\n")
else: # TODO test this
    log.write(str(ln_count_before) + " DOES NOT EQUAL " + str(ln_count) + " + " + str(ln_diff) + "\n\n")

#%% Clean the RP column
rp_count_before = rp_count
# apply function BEFORE drop duplicates
rp_df = make_base_names(rp_df, 'RELATED_PARTY_NAME', 'RELATED_PARTY_BASE', rp_reverse)
rp_df = rp_df.dropna(subset=['RELATED_PARTY_NAME'])
rp_df = rp_df.drop_duplicates(subset=['RELATED_PARTY_NAME'])
rp_count = len(rp_df)
rp_diff = rp_count_before-rp_count
log.write(str(rp_diff) + " empty or duplicative RELATED_PARTY_NAME instances found \n")
log.write(str(rp_count) + " unique non-null RELATED_PARTY_NAMES for analysis\n")
if rp_count_before-rp_diff == rp_count:
    log.write(str(rp_count_before) + " correctly equals " + str(rp_count) + " + " + str(rp_diff) + "\n\n")
else: # todo make sure else works - alter data
    log.write(str(rp_count_before) + " DOES NOT EQUAL "+ str(rp_count) + " + " + str(rp_diff) + "\n\n")


#%% do it again after cleaning to see what we've missed
#common_words_2 = pd.DataFrame(Counter(" ".join(rp_df['RELATED_PARTY_BASE']).split()).most_common(200))
#common_words_2.columns=['Word','Count']

#%% keep only the columns of interest
cross_df = ln_df.merge(rp_df, how='cross')
# TODO can get rid of one here (copy)
matches = cross_df.copy()
matches['RATIO_BASE'] = matches.apply(lambda x: fuzz.ratio(x['LOAN_BASE'], x['RELATED_PARTY_BASE']), axis=1)

#%%
low_score_threshold = 80
medium_score_threshold = 85
high_score_threshold = 90
perfect_score_threshold = 100
#%%
ln_count_before = len(matches['LOAN_NAME'].unique())
rp_count_before = len(matches['RELATED_PARTY_NAME'].unique())

#%%
non_matches = matches[matches['RATIO_BASE'] < low_score_threshold]
matches = matches[matches['RATIO_BASE'] >= low_score_threshold]

#%%
ln_match_series = pd.Series(matches['LOAN_NAME'].unique(), name='LOAN_NAME')
ln_nonmatch_series = pd.Series([i for i in non_matches['LOAN_NAME'].unique() if i not in matches['LOAN_NAME'].unique()], name='LOAN_NAME') # Drop any that are in matches
ln_nonmatch_count = len(ln_nonmatch_series)
ln_count = len(matches['LOAN_NAME'].unique())

rp_match_series = pd.Series(matches['RELATED_PARTY_NAME'].unique(), name='RELATED_PARTY_NAME')
rp_nonmatch_series = pd.Series([i for i in non_matches['RELATED_PARTY_NAME'].unique() if i not in matches['RELATED_PARTY_NAME'].unique()], name='RELATED_PARTY_NAME')
rp_nonmatch_count = len(rp_nonmatch_series)
rp_count = len(matches['RELATED_PARTY_NAME'].unique())

#non_matches = pd.concat([rp_nonmatch_series, ln_nonmatch_series], axis=1)
#matches_unique = pd.concat([rp_match_series, ln_match_series], axis=1)

log.write(str(ln_count_before) + ' unique LOAN_NAMES were returned from matching algorithm with no threshold applied' + '\n')
log.write(str(ln_count) + ' unique LOAN_NAMES were above matching threshold' + '\n')
log.write(str(ln_nonmatch_count) + ' unique LOAN_NAMES were under matching threshold \n')
if ln_count_before-ln_nonmatch_count == ln_count:
    log.write(str(ln_count_before) + " correctly equals " + str(ln_count) + " + " + str(ln_nonmatch_count) + "\n\n")
else:  # TODO test this
    log.write(str(ln_count_before) + " DOES NOT EQUAL " + str(ln_count) + " + " + str(ln_nonmatch_count) + "\n\n")


log.write(str(rp_count_before) + ' unique RELATED_PARTY_NAMES were returned from matching algorithm with no threshold applied' + '\n')
log.write(str(rp_count) + ' unique RELATED_PARTY_NAMES were above matching threshold \n')
log.write(str(rp_nonmatch_count) + ' unique RELATED_PARTY_NAMES were under matching threshold \n')
if rp_count_before-rp_nonmatch_count == rp_count:
    log.write(str(rp_count_before) + " correctly equals " + str(rp_count) + " + " + str(rp_nonmatch_count) + "\n\n")
else: # todo test this
    log.write(str(rp_count_before) + " DOES NOT EQUAL " + str(rp_count) + " + " + str(rp_nonmatch_count) + "\n\n")

#%%
matches['RATIO_ORDER'] = matches.apply(lambda x: fuzz.token_sort_ratio(x['LOAN_BASE'], x['RELATED_PARTY_BASE']), axis=1)
matches['RATIO_FULL'] = matches.apply(lambda x: fuzz.ratio(x['LOAN_NAME'], x['RELATED_PARTY_NAME']), axis=1)

matches = matches.dropna(how='all', axis=0)

matches = matches[['RELATED_PARTY_NAME','LOAN_NAME', 'LOAN_BASE', 'RELATED_PARTY_BASE', 'RATIO_BASE', 'RATIO_ORDER', 'RATIO_FULL', 'ACCOUNTS']]
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
if lnc['ACCOUNTS']==None:
    matches = matches.drop(columns=['ACCOUNTS'])

#%% Last count for log
ln_count = len(matches['LOAN_NAME'].unique())
rp_count = len(matches['RELATED_PARTY_NAME'].unique())

log.write(str(ln_count) + ' unique LOAN_NAMES were export to Above Threshold worksheet' + '\n')
log.write(str(ln_nonmatch_count) + ' unique LOAN_NAMES were export to the Below Threshold worksheet' + '\n\n')

log.write(str(rp_count) + ' unique RELATED_PARTY_NAMES were export to Above Threshold worksheet' + '\n')
log.write(str(rp_nonmatch_count) + ' unique RELATED_PARTY_NAMES were export to the Below Threshold worksheet' + '\n\n')
#%%

InfoDict = [
    ['Alterations Made to the Original Name:', ''],
        ['Removed leading and trailing spaces'],
        ['Format changes', 'Capitalized all characters'],
    ['Alterations Made to Modified Name: (fields end in “_BASE”)', ''],
        ['Removed punctuation',
            'Periods, apostrophes, slashes, hyphens, numbers, commas*, and extra spaces potentially caused by punctuation removal'],
        ['Format changes',
            'Capitalized all characters; if manual review of raw data shows names are in last, first name order, names were switched to first name last name form (*commas were removed after this)'],
        ['Name changes',
            'Drop Words: Words that were stripped from the name to get a more true name assessment score not influenced by these common words'],
            ['', '     Common Words: '+ str(drop_words)],
            ['', '     Car Names: ' + str(car_brands)],
            ['', 'Stop Words: Only the portion of the name that appeared before these words was kept (E.g. “Gordon Foods Dated: 12/2/2018” returns as “Gordon Foods”)'],
            ['', '     ' + str(stop_words)],
    ['Output fields',
        'RATIO_BASE: Levenshtein score of modified loan name (_BASE) compared to modified related party name (_BASE)'],
        ['', 'RATIO_ORDER: all letters in both modified names (_BASE) are rearranged into alphabetical order and then '
             'compared (Loan_name_BASE compared to Related_Party_BASE)'],
        ['', 'RATIO_FULL: Levenshtein score of the original names with minor modifications (all caps, removed leading and trailing spaces)'],
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
        ['', 'Low Confidence Matches: RATIO_BASE >= ' + str(low_score_threshold) + ' and < ' + str(medium_score_threshold)],
    ['Sources',
     'Names in the ' + lnc['LOAN_NAME'] + ' column of the ' + ln_df_filename + ' were compared against names in the  '
     + rpc['RP_NAME'] + ' column of the ' + rp_filename],
    ['Levenshtein',
        'For each comparison, a ‘ratio’ was computed which applied python’s FuzzyWuzzy package fuzz.ratio or '
        'fuzz.token_sort module to each combination. https://pypi.org/project/fuzzywuzzy/ '
        'Levenshtein distance is the number of changes required to move from the example to target. '
        'The Fuzz.ratio provides a ratio of 100 or less (100 indicating a perfect match) '
        'between each ‘Base Name’ combination. For example, when comparing 2 client names “A Razorback Land and Sea '
        'Holding” against “A Razorback Land and Sea Holdings”, Levenshtein will result in 98 distance ratio. '
        'A shorter name combination compared to one another will result in a different ratio. '
        'For example, Razorback vs Razerback will result in 95. A perfect match is ratio of 100.'],
    ['Modified Name Reasoning (fields end in "_BASE")',
        'This was performed to reduce common words which would decrease matching efficacy. For example, both '
        '“A Razorback Land and Sea Company Foundation” and “Razorback Land & Sea Co., LLC” would be reduced to '
        '“RAZORBACK LAND SEA”.'],
    ['Accounts (optional)',
        'If an account column is provided in the loan file, a list of all account numbers associated with a given '
        'LOAN_NAME is output']
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
writer = pd.ExcelWriter(export_path + "\\" + client_name + '_RelatedParties.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.

Info.to_excel(writer, sheet_name='Information', index=False)

perfect_base_matches.to_excel(writer, sheet_name='Perfect Base Matches', index=False)
perfect_full_matches.to_excel(writer, sheet_name='Perfect Full Matches', index=False)
likely_matches.to_excel(writer, sheet_name='Likely Matches', index=False)
medium_confidence.to_excel(writer, sheet_name='Medium Confidence Matches', index=False)
low_confidence.to_excel(writer, sheet_name='Low Confidence Matches', index=False)
matches.to_excel(writer, sheet_name='All Matches',index=False)
#matches_unique.to_excel(writer, sheet_name='Above Threshold', index=False)
#non_matches.to_excel(writer, sheet_name='Below Threshold', index=False)
ln_df_raw.to_excel(writer, sheet_name='Loans Evaluated', index=False)
rp_df_raw.to_excel(writer, sheet_name='Related Parties Evaluated', index=False)

# Get the xlsxwriter objects from the dataframe writer object.
workbook = writer.book
info_worksheet = writer.sheets['Information']
ln_df_raw = writer.sheets['Loans Evaluated']
rp_raw_worksheet = writer.sheets['Related Parties Evaluated']
#orksheet = writer.sheets['Below Threshold']
#above_worksheet = writer.sheets['Above Threshold']

# Add some cell formats.

formatbold = workbook.add_format({'bold': True})
format_wrap = workbook.add_format({'text_wrap': True})
under_border_format = workbook.add_format({'bottom': 1})
bottom_right_format = workbook.add_format({'bottom': 1, 'right': 1})
excel_format = workbook.add_format({'bg_color': '#D0E2C5', 'border': True})
right_format = workbook.add_format({'right': 1})
ln_color_format = workbook.add_format({'bg_color': '#dbe4f0'})
rp_color_format = workbook.add_format({'bg_color': 'e0ebe4'})
ratio_format = workbook.add_format({'bg_color': '#ebebeb'})

# Set column width for Information Sheet
info_worksheet.set_column('A:A', 59)
info_worksheet.set_column('B:B', 100, format_wrap)
# With Row/Column notation you must specify all four cells in the range: (first_row, first_col, last_row, last_col)

info_worksheet.conditional_format(1, 0, 1, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(4, 0, 4, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(12, 0, 12, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(15, 0, 15, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(21, 0, 21, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})
info_worksheet.conditional_format(26, 0, 29, 0, {'type': 'formula', 'criteria': 'True',  'format': formatbold})

info_worksheet.conditional_format(0, 0, 0, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(3, 0, 3, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(11, 0, 11, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(14, 0, 14, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(20, 0, 20, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})
info_worksheet.conditional_format(25, 0, 29, 0, {'type': 'formula', 'criteria': 'True',  'format': under_border_format})

info_worksheet.conditional_format(0, 1, 0, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(3, 1, 3, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(11, 1, 11, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(14, 1, 14, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(20, 1, 20, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})
info_worksheet.conditional_format(25, 1, 29, 1, {'type': 'formula', 'criteria': 'True',  'format': bottom_right_format})

info_worksheet.conditional_format(1, 1, 25, 1, {'type': 'formula', 'criteria': 'True', 'format': right_format})

# Set the column width and format
def set_match_worksheet_format(match_worksheet_name):
    match_worksheet = writer.sheets[match_worksheet_name]

    match_worksheet.freeze_panes(1, 1)
    match_worksheet.set_column('A:A', 30, rp_color_format)
    match_worksheet.set_column('B:B', 30, ln_color_format)
    match_worksheet.set_column('C:C', 25, ln_color_format)
    match_worksheet.set_column('D:D', 25, rp_color_format)
    match_worksheet.set_column('E:G', 13, ratio_format)
    if lnc['ACCOUNTS'] is not None:
        match_worksheet.set_column('H:H', 30, ln_color_format)


set_match_worksheet_format('Perfect Base Matches')
set_match_worksheet_format('Perfect Full Matches')
set_match_worksheet_format('Likely Matches')
set_match_worksheet_format('Medium Confidence Matches')
set_match_worksheet_format('Low Confidence Matches')
set_match_worksheet_format('All Matches')

# format non_matches
#below_worksheet.freeze_panes(1, 1)
#below_worksheet.set_column('A:A', 30, rp_color_format)
#below_worksheet.set_column('B:B', 30, ln_color_format)

#above_worksheet.freeze_panes(1, 1)
#above_worksheet.set_column('A:A', 30, rp_color_format)
#above_worksheet.set_column('B:B', 30, ln_color_format)

# Close the Pandas Excel writer and output the Excel file.
writer.close()


log.write("\n")
log.write("\n")
log.close()
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
