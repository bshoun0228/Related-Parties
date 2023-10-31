from Related_Parties_Function import *
import datetime

#%%

starttime = datetime.datetime.now()
print("Start: ", starttime)

# %% ###################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = '2023 Example'  # what do you want files named?
export_path = r"C:\Users\kl8475\OneDrive - FORVIS, LLP\Related Parties Examples\2023 Example"  # where you want results

# Put the filepath to the GL/Other data
ln_df_filepath = r"C:\Users\kl8475\OneDrive - FORVIS, LLP\Related Parties Examples\2022 Uwharrie\Loan Trial Balance 12.31.2022.xlsx"

# What column are we comparing? Enter the column headers for the Customer Name and Account columns
## If there is no Account column provided, type None
lnc = {'LOAN_NAME': 'Customer Name', 'ACCOUNTS': 'Account Number'}

# IF the column in the LASTNAME, FIRST NAME format, type 'YES' (CAPITAL)
ln_reverse = 'YES'

# Put the filepath to the related parties
rp_filepath = r"C:\Users\kl8475\OneDrive - FORVIS, LLP\Related Parties Examples\2022 Uwharrie\RP 4 - Related Party Related Interest Log - Updated.xlsx"

# What column are we comparing?
rpc = {'RP_NAME': 'Related Interests'}
# Is this column in the LASTNAME, FIRST NAME format? # 'YES' or 'NO'
rp_reverse = 'NO'

#%%
run_related_parties(client_name=client_name, export_path=export_path, ln_df_filepath=ln_df_filepath, lnc=lnc,
                    ln_reverse=ln_reverse, rp_filepath=rp_filepath, rpc=rpc, rp_reverse=rp_reverse)

# %% Calculate how long it took
def convert(seconds):
    seconds = seconds % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60

    return "%d:%02d:%02d" % (hour, minutes, seconds)

endtime = datetime.datetime.now()
print("End: ", endtime)
run_time = (endtime - starttime)
run_seconds = run_time.total_seconds()
print("Runtime (HH:MM:SS): ", convert(run_seconds))



