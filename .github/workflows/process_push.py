# This file runs every time a commit is pushed to the repo. It carries out
# the following steps
#   1. Ensure all expected files are present
#   2. Ensure the Python script matches the Python in the Excel
#   3. Extract the VBA code from the Excel
#   4. Update the README file with any errors and the version number
#   5. Create a new commit with all these goodies

# =====================
# =  Import packages  =
# =====================

import os
import shutil
import sys
from oletools.olevba3 import VBA_Parser

import tempfile
import pathlib
import shutil

import pandas as pd

# ======================
# =  Define constants  =
# ======================

# The extensions we should look for when extracting VBA code; the script will
# ensure there is only *one* file with these extensions, and extract VBA from
# that file
ADDIN_FILE       = 'XLKitLearn.xltm'

# The name of the file containing the Python code
PYTHON_FILE      = 'XLKitLearn.py'

# The name of the readme file
README_FILE      = 'README.md'

# The name/location of the HTML page
HTML_PAGE        = 'docs/index.html'

# The comment in the readme file; all errors will be output ABOVE that line.
# Anything below the line won't be modified
README_COMMENT   = '<!-- DO ***NOT*** EDIT ANYTHING ABOVE THIS LINE, INCLUDING THIS COMMENT -->'

# The path of the directory in which to extract VBA code
VBA_DIRECTORY    = '~VBA Code'

# ======================
# =  Define Utilities  =
# ======================

def parse_vba(workbook_path, extract_path, KEEP_NAME=False):
    '''
    Given the path of a workbook, this will extract the VBA code from the workbook
    in workbook_path to the folder in extract_path.
    
    If KEEP_NAME is True, we keep the line "Attribute VB_Name" in the files
    
    Slightly modified from the code in
       https://www.xltrail.com/blog/auto-export-vba-commit-hook
    '''
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, filename, content in vba_modules:
        try:
            decoded_content = content.decode('latin-1')
        except:
            decoded_content = content
            
        lines = []
        if '\r\n' in decoded_content:
            lines = decoded_content.split('\r\n')
        else:
            lines = decoded_content.split('\n')
        if lines:
            content = []
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    if not os.path.exists(os.path.join(extract_path)):
                        os.makedirs(extract_path)
                    with open(os.path.join(extract_path, filename), 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content).lower())

def update_readme(filename, errors, version='ERROR'):
    '''
    This function will find the readme file in the repo, and
      - Ensure it is formatted correct; if it isn't, an error will be raised
      - If there are any errors in the errors list, create an error report at the
        top of the file
      - Otherwise, output the version number
    '''
    
    # Copy the errors to ensure we can't modify the list
    errors = list(errors)
    
    # If the readme file doesn't exist, an error will be thrown
    with open(README_FILE) as f:
        readme_file = f.readlines()
    
    # Find the line with the comment
    comment_line_n = [i for i, j in enumerate(readme_file) if README_COMMENT in j]
    assert len(comment_line_n) == 1, 'Readme file incorrectly formatted'
    comment_line_n = comment_line_n[0]
    
    # Remove everything before that line
    readme_file = readme_file[comment_line_n:]
    
    # Insert errors if there are any
    if len(errors) > 0:
        # Prepend an intro to the errors
        errors[0:0] = ['## ERROR REPORT',
                       '**Some errors were found while processing your '
                                'latest push. Please fix them before proceeding as follows**',
                       '  - Pull the latest commit (made by the VBA robot) from github',
                       '  - Make your changes',
                       '  - Push back to github',
                       '',
                       'The specific errors I found were as follows:']
        
        # Add line breaks
        errors = [i + '\n' for i in errors]
        
        # Add a line
        errors.append('---\n')
        
        # Add to the readme file
        readme_file[0:0] = errors
    else:
        # If not, insert the version number
        readme_file.insert(0, f'## Version: {version}\n')
    
    # Insert the readme header
    readme_file.insert(0, '# XLKitLearn\n')

    # Write the file back
    with open(README_FILE, 'w') as f:
        f.writelines(readme_file)

# =================
# =  Main Script  =
# =================

if __name__ == '__main__':
    # Create an empty list for potential errors
    errors = []
    
    # ---------------------------------------
    # -  Ensure expected files are present  -
    # ---------------------------------------
    
    file_errors = ["  - Some files I expected to find weren't quite right"]
    
    # Look for the Python file
    if not os.path.isfile(PYTHON_FILE):
        file_errors.append(f"      * I expected to find a file {PYTHON_FILE} containing the Python code for the add-in, but it wasn't there")
    
    # Look for the Excel file
    if not os.path.isfile(ADDIN_FILE):
        file_errors.append(f"      * I expected to find a file {ADDIN_FILE} containing the add-in, but it wasn't there")
      
    # If we had any file read errors, output them and leave
    if len(file_errors) > 1:
        update_readme(README_FILE, file_errors)
        sys.exit()
    
    # -----------------------------------------------------
    # -  Ensure the Python in the Excel matches the file  -
    # -----------------------------------------------------
    
    # Load the Python in the file
    with open(PYTHON_FILE) as f:
        python_from_file = f.readlines()
    
    # Load the Python from the Excel
    python_from_excel = pd.read_excel(ADDIN_FILE,
                                      sheet_name = 'code_text',
                                      header     = None,
                                      dtype      = str,
                                      na_filter  = False).values
        
    # Check whether we have a prod version by looking at the first row
    # of the code
    try:
        first_row = list(python_from_excel[0])
        prod_version = (first_row[first_row.index('prepped_for_prod')+1] == 'PROD')
    except:
        prod_version = False
    
    # Extract the code only
    python_from_excel = [i[0].replace('_x000D_', '\n') for i in python_from_excel[1:]]
    
    # Ensure the two match
    if python_from_file != python_from_excel:
        errors.append(f'  - The Python code in the Excel file does not match the Python code in {PYTHON_FILE}. To fix this, run the Excel file at least once in debug mode to load the new Python.')
    
    # ----------------------------
    # -  Get the version number  -
    # ----------------------------
    
    # The version number is in the seventh line of the Python code
    version_line = python_from_file[6]
    
    if 'ADDIN_VERSION ' not in version_line:
        errors.append(f'  - Could not read the version number from the Python code.')
    else:
        version = version_line.split()[-1][1:-1]
        
    # If we are not in prod mode, add "dev" to the version number
    if not prod_version:
        version = version + 'dev'
    
    # --------------------------
    # -  Extract the VBA code  -
    # --------------------------
    try:
        # If the VBA_DIRECTORY exists, delete it
        if os.path.exists(VBA_DIRECTORY):
            shutil.rmtree(VBA_DIRECTORY)
        
        # Extract the VBA code files
        parse_vba(ADDIN_FILE, VBA_DIRECTORY)
    
    except Exception as e:
        errors.append(f'  - Error extracting the VBA code from the workbook; the error was {str(e)}')
    
    # -------------------------------------
    # -  Update version on the html page  -
    # -------------------------------------
    
    try:
        # Read the HTML page
        with open(HTML_PAGE) as f:
            html_file = f.readlines()
        
        # Find the line with the version
        version_line = [i for i, j in enumerate(html_file) if '<h2>Latest Version:' in j]
        assert len(version_line) == 1
        version_line = version_line[0]
        
        # Edit that line with the version
        html_file[version_line] = f'<h2>Latest Version: {version}</h2>\n'
        
        with open(HTML_PAGE, 'w') as f:
            f.writelines(html_file)
    except Exception as e:
        errors.append(f'  - Error updating the version number in the HTML file; the error was {str(e)}')
    
    # ----------------------------
    # -  Update the readme file  -
    # ----------------------------
    
    update_readme(README_FILE, errors, version)