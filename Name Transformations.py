# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
import numpy as np
import datetime

#%%
starttime = datetime.datetime.now()
print("Start: ", starttime)


#%% ####################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = 'RP_OTHER_COL'
# Put the filepath to the GL/Other data
df_filepath = 'Data/Related Parties Clean.xlsx'
# What column are we comparing?
dfc = {'LOAN_NAME': 'OTHER'}
# IF the column in the LASTNAME, FIRST NAME format, type 'YES'
df_reverse = 'NO'

########################################################################################################################
#%%
df = pd.read_excel(df_filepath)

#%% Pull out the filenames for the info tab
df_filename = df_filepath.split("/")[-1]


# %%
df = df.drop_duplicates(subset=[dfc['LOAN_NAME']])
# Extract the string from our comparison column and add "_BASE"
df_base = str(dfc['LOAN_NAME']) + "_BASE"

# Create an OG column
df_og = str(dfc['LOAN_NAME']) + "_OG"
df[df_og] = df[dfc['LOAN_NAME']]

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

#%% keep only the columns of interest
df_short = df[[df_og, dfc['LOAN_NAME'], df_base]]

#%%
# EXPORT
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(client_name + '_Names.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.

df_short.to_excel(writer, sheet_name='Names',index=False)

# Get the xlsxwriter objects from the dataframe writer object.
workbook = writer.book
names_worksheet = writer.sheets['Names']

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


# Set the column width and format
names_worksheet.freeze_panes(1, 1)
names_worksheet.set_column('A:A', 30)
names_worksheet.set_column('B:B', 30)
names_worksheet.set_column('C:C', 25)


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
