with open('data/initial_compile.log', 'w') as f:
    f.write('Started')

import traceback

import_string = '''
import xlwings as xw
# Basic packages
import pandas as pd
import numpy as np
import scipy as sp
import itertools
import time
import warnings
import keyword
import os
import signal
from collections import OrderedDict
import traceback
import sys

# Data preparation utilities
import patsy as pt
import sklearn.feature_extraction.text as f_e
import nltk

# Sparse matrices
from scipy import sparse

# Estimators, learners, etc...
warnings.filterwarnings("ignore", category = DeprecationWarning)
from sklearn.linear_model import LinearRegression, LogisticRegression, Lasso
from sklearn.neighbors import KNeighborsRegressor, KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor
from sklearn.decomposition import LatentDirichletAllocation
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.ensemble import GradientBoostingClassifier, GradientBoostingRegressor

# Statsmodels learners for p-values
import statsmodels.api as sm

# Validation utilities and metrics
from sklearn import model_selection as sk_ms
from sklearn.metrics import r2_score, roc_auc_score, roc_curve, make_scorer
import sklearn.inspection as sk_i

# Plotting (ensure we don't use qt)
import matplotlib as mpl
import matplotlib.pyplot as plt
import seaborn as sns
import json

# Verification
from datetime import datetime, timedelta
import hashlib
import requests'''.split('\n')
    
print('=====================================')
print('=  Preparing Python for XLKitLearn  =')
print('=====================================')
print()
print('Please wait - this might take a while...')
print()

try:
    for l_n, l in enumerate(import_string):
        exec(l)
        
        perc_done = (l_n+1)/len(import_string)
        total_dashes = 40
        dashes_done = int(perc_done*total_dashes)
        perc_done = round(perc_done*100,)
        
        print('Progress [' + '-'*dashes_done + ' '*(total_dashes - dashes_done) + '] ' + str(perc_done) + '%', end='\r')
    
    with open('data/initial_compile.log', 'w') as f:
        f.write('Success')
    
except Exception as e:
    with open('data/initial_compile.log', 'w') as f:
        f.write('ERROR')
        f.write('')
        f.write(traceback.format_exc())