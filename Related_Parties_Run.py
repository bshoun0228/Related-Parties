from Related_Parties_Function import *

# %% ###################################################################################################################
#############################################    FILL THIS OUT  ########################################################
# Client Name (for export Name)
client_name = 'Related Party PRO Demo'  # what do you want files named?
export_path = r"C:\Users\kl8475\OneDrive - FORVIS, LLP\Related Parties Examples\Anon_Data"  # where you want results

# Put the filepath to the GL/Other data
ln_df_filepath = r"C:\Users\kl8475\OneDrive - FORVIS, LLP\Related Parties Examples\Anon_Data\Loan Flat File Anon.csv"

# What column are we comparing? Enter the column headers for the Customer Name and Account columns
## If there is no Account column provided, type None
lnc = {'LOAN_NAME': 'Customer Name', 'ACCOUNTS': 'Account Number'}

# IF the column in the LASTNAME, FIRST NAME format, type 'YES' (CAPITAL)
ln_reverse = 'NO'

# Put the filepath to the related parties
rp_filepath = r"C:\Users\kl8475\OneDrive - FORVIS, LLP\Related Parties Examples\Anon_Data\Related Parties Anon.csv"

# What column are we comparing?
rpc = {'RP_NAME': 'RP NAME'}
# Is this column in the LASTNAME, FIRST NAME format?
rp_reverse = 'NO'

#%%
run_related_parties(client_name=client_name, export_path=export_path, ln_df_filepath=ln_df_filepath, lnc=lnc,
                    ln_reverse=ln_reverse, rp_filepath=rp_filepath, rpc=rpc, rp_reverse=rp_reverse)

