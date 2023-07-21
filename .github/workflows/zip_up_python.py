# This file will zip up the Python distro for people who can't use the
# installer

# =====================
# =  Import packages  =
# =====================

import os
import shutil

# =================
# =  Main Script  =
# =================

if __name__ == '__main__':
    
    # Load the Python location
    python_location = os.environ['pythonLocation']
    
    # Zip up that entire folder
    shutil.make_archive('python-distro.zip' ,' zip', python_location)