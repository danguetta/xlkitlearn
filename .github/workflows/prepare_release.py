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

# The code file
CODE_FILE    = 'XLKitLearn.py'

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
        
    # Write the version file - this will be used for the mac installer, and also will be pushed to the
    # release for the server to check
    with open('version', 'w') as f:
        f.write(version)
    
    # Find the EARLIEST_ALLOWABLE_VERSION and write it to a file
    with open(CODE_FILE, 'r') as f:
        lines = f.readlines()
        assert len(lines) > 8, 'Code file is too short'
        this_line = lines[7]
        assert this_line.startswith('EARLIEST_ALLOWABLE_VERSION'), 'Eight line of the code file does not have the earliest allowable version'
        
        earliest_allowable_version = this_line.split("'")[1].strip()
    
    with open('earliest_allowable_version', 'w') as f:
        f.write(earliest_allowable_version)
        
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
    
    req_run_string = "import xlwings; import pandas; import numpy; import scipy; import itertools; import time; import warnings; import keyword; import os; import signal; import sys; import patsy; import sklearn.feature_extraction.text; import nltk; warnings.filterwarnings('ignore', category = DeprecationWarning); from sklearn.linear_model import LinearRegression, LogisticRegression, Lasso; from sklearn.neighbors import KNeighborsRegressor, KNeighborsClassifier; from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor ;from sklearn.decomposition import LatentDirichletAllocation; from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor; from sklearn.ensemble import GradientBoostingClassifier, GradientBoostingRegressor; import statsmodels.api as sm; from sklearn import model_selection as sk_ms; from sklearn.metrics import r2_score, roc_auc_score, roc_curve, make_scorer; import sklearn.inspection as sk_i; import matplotlib as mpl; import matplotlib.pyplot as plt; import seaborn as sns; import json; from datetime import datetime, timedelta; import hashlib; import requests; import tiktoken; import openai"
    
    # ----------------------------
    # -  Prep Windows Installer  -
    # ----------------------------
    
    # Create a file with a version stamp for the windows installer
    os.mkdir(os.path.join(env['pythonLocation'], 'data'))
    with open(os.path.join(env['pythonLocation'], 'data', 'version'), 'w') as f:
        f.write(env['RELEASE_TAG'] + '\n')
    
    # Prepare an empty file for the offline runs
    with open(os.path.join(env['pythonLocation'], 'data', 'offline_runs'), 'w') as f:
        f.write('0\n')

    # Replace the version number in installer.iss
    with open(ISS_FILE, 'r') as f:
        iss_file = f.read()
        
    with open(ISS_FILE, 'w') as f:
        f.write(iss_file.replace('version_placeholder', version)
                        .replace('python_command_placeholder', req_run_string))

    # -----------------------
    # -  Prep Mac installer  -
    # ------------------------
    
    # Version number file was saved above 

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