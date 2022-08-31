# IMPORTS
import pandas as pd
pd.set_option('mode.chained_assignment', None)
from collections import Counter
from fuzzywuzzy import fuzz
import numpy as np

#%% READ IN THE DATA
df = pd.read_excel('Data/Loan_Name.xlsx')
rp = pd.read_excel('Data/Related Parties Clean.xlsx')

#%% CLEAN THEM UP
# %% Clean the loan column
df['BASE'] = df['LOAN_NAME'].astype(str)

df['BASE'] = df['BASE'].str.upper()
df['BASE'] = df['BASE'].str.strip()
df['BASE'] = df['BASE'].str.replace(',', '')
df['BASE'] = df['BASE'].str.replace('.', '')
df['BASE'] = df['BASE'].str.replace("'S", "")

#%% stop words
stopwords = ['FOUNDATION', 'HOLDINGS', 'MANAGEMENT', 'INVESTMENTS', 'PROPERTIES', 'INTERNATIONAL', 'THE', '401K',
             'PARTNERSHIP', 'LIMITED', 'ENTERPRISES', 'ASSOCIATES', 'PARTNERS', 'INVESTMENT', 'GROUP', 'COMPANY',
             'ASSOCIATION', '401 (K)', '401(K)', 'LLC', 'HOLDING', 'INVESTORS', 'INC', '-', 'AND', '&']

# Drop the stopwords
df['BASE'] = df['BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in stopwords]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
df['BASE'] = df['BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

#%% do it again after cleaning to see what we've missed
common_words_2 = pd.DataFrame(Counter(" ".join(df['BASE']).split()).most_common(200))
common_words_2.columns=['Word','Count']

#%% Clean the RP column
rp = rp.dropna(subset=['OCC'])
rp['BASE'] = rp['OCC'].astype(str)
rp['BASE'] = rp['BASE'].str.upper()
rp['BASE'] = rp['BASE'].str.strip()
rp['BASE'] = rp['BASE'].str.replace(',', '')
rp['BASE'] = rp['BASE'].str.replace('.', '')
rp['BASE'] = rp['BASE'].str.replace("'S", "")

#%% stop words
# Drop the stopwords
rp['BASE'] = rp['BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in stopwords]))

#%% Drop only ESTATE if it is not part of REAL ESTATE
reallist = ['ESTATE']
rp['BASE'] = rp['BASE'].apply(lambda x: ' '.join([word for word in x.split() if word not in reallist]) if (('ESTATE' in x) & ('REAL ESTATE' not in x)) else x)

#%% do it again after cleaning to see what we've missed
common_words_2 = pd.DataFrame(Counter(" ".join(rp['BASE']).split()).most_common(200))
common_words_2.columns=['Word','Count']

#%%

#%%

mm = pd.DataFrame(columns=['LOAN_NAME', 'NAME', 'RATIO'])

for rp_base in rp['BASE']:
    matches = df.copy()
    matches['RATIO'] = matches.apply(lambda x: fuzz.ratio(x['BASE'], rp_base), axis=1)
    matches['RP_BASE'] = rp_base
