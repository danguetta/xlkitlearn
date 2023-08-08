# XLKitLearn
## Version: 11.04dev
<!-- DO ***NOT*** EDIT ANYTHING ABOVE THIS LINE, INCLUDING THIS COMMENT -->

This repo contains the latest version of [XLKitLearn](https://www.xlkitlearn.com). Please see the website for authorship, license, installation, and usage information - this repo provides information for those interested in seeing the add-in's code and/or contributing to it.

There is a separate GitHub repo for the add-in manual [here](https://github.com/danguetta/XLKitLearn_docs).

## Understanding the repo

The main file in this repo is `XLKitLearn.xltm` - this is an Excel template which, when opened, creates a blank version of the add-in. It contains the following sheets:
  - `Add-in`; together with the VBA code and user forms, this sheet comprises the front-end of the add-in.
  - `boston_housing`; a sample dataset (the [Boston housing dataset](http://lib.stat.cmu.edu/datasets/boston)).
  - `xlwings.conf`; the XLWings configuration settings (this sheet is deep-hidden in the final product - see [debug mode](#understanding-debug-mode) below).
  - `code_text`; the full Python code comprising the add-in's back-end (this sheet is deep-hidden in the final product - see [debug mode](#understanding-debug-mode) below). This code should never be edited directly in the sheet; instead, edit `XLKitLearn.py` (see [contributing](#contributing-to-the-add-in)).

In addition, the repo contains:
  - A file called `XLKitLearn.py`, which contains an exact copy of the code in the `code_text` sheet, for easier editing (see [contributing](#contributing-to-the-add-in) for details on how these are kept in sync). **The top of this file contains the version number for the add-in ; version numbers everywhere else propagate from here.**.
  - A folder called `~VBA Code`, which contains an exact copy of the VBA code in the add-in; this is to make sure diffs can be tracked in github. This should never be edited directly; instead, make changes directly in the VBA in the Excel file (see [contributing](#contributing-to-the-add-in) for details on how this is kept in sync).
  - A file called `requirements.txt`, which contains the Python packages needed to make XLKitLearn work.
  - A folder called `.github`, which contains all the scripts that automate the repo, and create the installer.

## Contributing to the add-in

**Begin by making a new branch; NEVER commit to `main` directly** 

To contribute to the add-in, create a new branch, and
  - To edit the VBA code and front-end, edit `XLKitLearn.xltm` directly. Make sure that as soon as you open the file, you immediately save it as an `xltm` file; this will ensure you are using the file in [debug mode](#understanding-debug-mode).
  - To edit the Python code, edit `XLKitLearn.py` directly. Whenever the add-in is run in [debug mode](#understanding-debug-mode), that entire file will be read, and the Python code in the `code_text` sheet will be replaced with the new Python code.

Whenever you commit to the repo, a bot will carry out the following steps:
  - Check the Python code in Excel matches the code in `XLKitLearn.py` exactly.
  - Extract the VBA code in the `xltm` file into the `~VBA Code` folder.
  - Update the README file to reflect the version number (if any errors are found, those errors will be printed at the top of the README file).

**IMPORTANT**: when the process is done, it will create a new commit in GitHub. *Immediately pull this new commit* so that your local copy isn't behind the origin. After this is done, look at the top of this `readme` file - if any errors occurred, they will be listed at the top of the file.

## Understanding Debug Mode

When users save the add-in file, it will save as an `xlsm` file. When you are devving against the add-in, you should save it as an `xltm` file (see [above](#contributing-to-the-add-in)). This will ensure the add-in runs in debug mode, which will have the following effect
  - Some `On Error Resume Next` statements in VBA will be ignored, to make sure errors are triggered in a way that is useful for debugging.
  - The `xlwings.conf` and `code_text` sheets will be visible.
  - Every time the add-in is run, the code from `XLKitLearn.py` will be read and loaded into `code_text`, to ensure the latest version of the code is in the Excel.

In some cases (for example, debugging a file with a user), you'll want the the file to launch in debug mode even with an `xlsm` extension. To make this happen, simply rename the file to contain the word `DEBUgG` (with two Gs).

## Releasing a new version of the add-in

When you are ready to release a new version of the add-in:
  - Run `prepare_for_prod` in the VBA immediate window to tidy up the workbook and prepare it for production (this will, for example, remove any extraneous sheets, and take the workbook out of debug mode). It will also require a password to update the version of the add-in on the server. Note that if `prepare_for_prod` has not been run, `dev` will be appended to the version name in the `README` file.
  - Commit your changes.
  - **WAIT** for the github action described [here](#contributing-to-the-add-in) to complete.
  - Create a new release in GitHub, based on the commit created by the VBA robot in the previous step. Github actions will check that everything is in order, and create installers that will be uploaded to the github release. Note that the release will fail if `preare_for_prod` was not run.
  
## Modifying external packages

  - When updating the version of XLWings, ensure the `ShowError` function includes the line `log_vba_error (Content)` to log errors to the server.
  - In the XLWings `CleanUp` function, add a call to `format_sheet` at the very end of the function
  - When updating `mdl_onedrive_path`, remove the `msgbox` line.