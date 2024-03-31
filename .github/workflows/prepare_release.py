# This file runs every time a new release is created

# =====================
# =  Import packages  =
# =====================

import os
import sys

# ======================
# =  Define constants  =
# ======================

# The file describing the windows installer
ISS_FILE     = '.github/installer.iss'

# The template for the mac installer
MAC_TEMPLATE = '.github/mac_installer_template.sh'

# The name of the Python requirements
REQ_FILE     = 'requirements.txt'

# The name of the readme file
README_FILE  = 'README.md'

# =================
# =  Main Script  =
# =================

if __name__ == '__main__':
    
    # Load the environment
    env = os.environ
    
    # --------------------------------
    # -  General Prep and Debugging  -
    # --------------------------------
    
    # Ensure the last commit was made by the VBA Robot (i.e., that the checks
    # have processed)
    # TODO
    #assert env['LAST_COMMIT_AUTHOR'] == 'VBA Robot', 'VBA Robot did not complete'
    
    # Ensure that the readme had no errors
    with open(README_FILE, 'r') as f:
        readme_file = f.read()
    
    assert '## ERROR REPORT' not in readme_file, 'VBA Robot found some errors'
    
    # Get the version number from the readme file
    version = [i for i in readme_file.split('\n') if i.startswith('## Version: ')]
    assert len(version) == 1, 'Could not find version number'
    version = version[0].split()[-1]
    
    # Ensure we do not have a dev version
    assert 'dev' not in version, 'Please run prepare_for_prod in VBA before trying to create a release'
    
    # Ensure the tag and release name are equal to the version
    assert version == env['RELEASE_TAG'], 'Release tag is not equal to the #version'
    assert version == env['RELEASE_NAME'], 'Release name is to equal to the version'
        
    # ---------------------------------------
    # -   Create a requirement run string   -
    # ---------------------------------------
    
    '''
    with open(REQ_FILE, 'r') as f:
        req_file = f.read()
        
    req_file = [i for i in req_file.split('\n') if len(i) > 0 and i[0] != '#']
    req_pc = '; '.join(['import ' + i.split('==')[0] for i in req_file if 'Darwin' not in i])
    req_mac = '; '.join(['import ' + i.split('==')[0] for i in req_file if 'Windows' not in i])
    '''
    
    req_run_string = "import xlwings; import pandas; import numpy; import scipy; import itertools; import time; import warnings; import keyword; import os; import signal; import sys; import patsy; import sklearn.feature_extraction.text; import nltk; warnings.filterwarnings('ignore', category = DeprecationWarning); from sklearn.linear_model import LinearRegression, LogisticRegression, Lasso; from sklearn.neighbors import KNeighborsRegressor, KNeighborsClassifier; from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor ;from sklearn.decomposition import LatentDirichletAllocation; from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor; from sklearn.ensemble import GradientBoostingClassifier, GradientBoostingRegressor; import statsmodels.api as sm; from sklearn import model_selection as sk_ms; from sklearn.metrics import r2_score, roc_auc_score, roc_curve, make_scorer; import sklearn.inspection as sk_i; import matplotlib as mpl; import matplotlib.pyplot as plt; import seaborn as sns; import json; from datetime import datetime, timedelta; import hashlib; import requests"
    
    # ----------------------------
    # -  Prep Windows Installer  -
    # ----------------------------
    
    # Create a file with a version stamp for the windows installer
    os.mkdir(os.path.join(env['pythonLocation'], 'data'))
    with open(os.path.join(env['pythonLocation'], 'data', 'version'), 'w') as f:
        f.write(env['RELEASE_TAG'] + '\n')
    
    # Replace the version number in installer.iss
    with open(ISS_FILE, 'r') as f:
        iss_file = f.read()
        
    with open(ISS_FILE, 'w') as f:
        f.write(iss_file.replace('version_placeholder', version)
                        .replace('python_command_placeholder', req_run_string))

    # -------------------------------------
    # -   Create the mini Word2Vec file   -
    # -------------------------------------
        
    # Create a reduced version of Word2Vec for XLKit Learn

    import gensim.downloader as downloader
    import pandas as pd
    import pickle
    import unicodedata

    # Get the model
    original_stdout = sys.stdout
    sys.stdout = open(os.devnull, 'w')

    w2v = downloader.load('word2vec-google-news-300')

    sys.stdout.close()
    sys.stdout = original_stdout
    
    # Convert it to a Pandas DataFrame
    v_len = len(w2v.index_to_key)
    df = pd.DataFrame({'w_id' : list(range(v_len))                                  ,
                    'word' : list(w2v.index_to_key)                              ,
                    'freq' : [w2v.get_vecattr(i, 'count') for i in range(v_len)]  })

    # Create a lower case, non-accented verison of the word
    df['word_lower'] = df.word.str.lower().apply(lambda x : ''.join([i for i in unicodedata.normalize('NFD', x) if not unicodedata.combining(i)]))
    df_ag = df.groupby('word_lower').agg(ids     = ('w_id', list),
                                        freqs   = ('freq', list),
                                        av_freq = ('freq', 'max')).reset_index()
    df_ag = df_ag.sort_values('av_freq', ascending=False)

    # Only keep words
    df_ag = df_ag[df_ag.word_lower.apply(lambda x : all(i in 'abcdefghijklmnopqrstuvwxyz0123456789' for i in x))]

    # Only keep the top entries
    df_ag_small = df_ag.head(100000)

    # Output
    out = {}
    for i, row in df_ag_small.iterrows():
        out[row.word_lower] = sum([w2v[ind]*freq for ind, freq in zip(row.ids, row.freqs)])/sum(row.freqs)

    pickle.dump(out, open(os.path.join(env['pythonLocation'], 'data', 'w2v_small.bin'), 'wb'))

    # -----------------------
    # -  Prep Mac installer  -
    # ------------------------
    
    # Create a file with the version number for the release
    with open('version', 'w') as f:
        f.write(version)
    
    pickle.dump(out, open('w2v_small.bin', 'wb'))

    # Load the template installer
    with open(MAC_TEMPLATE, 'r') as f:
        mac_installer = f.read()
    
    # Load the requirements
    with open(REQ_FILE, 'r') as f:
        req_file = f.read().strip()
    
    # Replace the placeholders
    mac_installer = ( mac_installer.replace('{{version_placeholder}}', version)
                                   .replace('{{requirements_placeholder}}', req_file)
                                   .replace('python_command_placeholder', req_run_string) )
    
    # Output the installer
    with open('XLKitLearn_installer.sh', 'w') as f:
        f.write(mac_installer)
    
    # Change windows newlines to mac
    with open('XLKitLearn_installer.sh', 'rb') as f:
        mac_installer = f.read()
    
    mac_installer = mac_installer.replace(b'\r\n', b'\n')
    
    with open('XLKitLearn_installer.sh', 'wb') as f:
        f.write(mac_installer)