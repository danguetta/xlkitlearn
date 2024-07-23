option explicit

' ###############################################################
' #  mdl_global_vars                                            #
' #                                                             #
' #  defines all the global variables used throughout the vba   #
' #  project, names stored in the spreadsheet, and constants    #
' ###############################################################

' constants - cell addresses
' ==========================

' addresses of the cells storing the add-in settings and email on the
' add-in sheet
public const cell_current_settings = "'add-in'!d9"
public const cell_current_text_settings = "'add-in'!d14"
public const cell_email = "'add-in'!f17"

' address of the cells containing the status as the add-in runs; because
' this will need to be passed to python, we can't include the sheet in it
public const cell_status = "f7"
public const cell_status_text = "f12"

' address of the first cell containing permanent variables
public const cell_perm_vars = "'code_text'!a1"

' constants - utilities
' =====================

' missing value sentinel
public const missing = "xxxxxxxxxxxxx"

' color constants
public const white = -2147483643
public const red = 12632319
public const grey = 8421504

' width of button in frm_pred
' ===========================

public const formula_width_large = 180
public const formula_width_small = 114

' constants - add-in text
' =======================

public const blank_range_selector = "(click to select)"

' workbook welcome
public const t_welcome_text_pc = "welcome to xlkitlearn! please begin by pressing ctrl+s to save this file somewhere."
public const t_welcome_text_mac = "welcome to xlkitlearn! please begin by pressing command+s to save this file somewhere."

' formula checker
public const t_formula_without_tilde = "every formula needs to be in the format y ~ x. for example, " & _
                                       "y ~ x1 && y ~ x1 + x2 is correct, but y ~ x1 && x1 + x2 is incorrect."
public const t_formula_no_y = "at least one of your formulas does nto have a term to the left of the ~ (i.e. " & _
                              "it is missing an outcome variable). the problematic formula was "
public const t_formula_diff_y = "when you enter multiple formulas, the outcome variable to the left of ~ must " & _
                                "be the same for all formulas. k-fold cross validation wouldn't really make " & _
                                "sense otherwise."
public const t_missing_y_values = "you are using {{y_var}} as an output variable, but that column has missing " & _
                                  "values in some rows"
public const t_nonnumeric_y_values = "you are using {{y_var}} as an output variable, but that column has " & _
                                    "non-numeric values in some rows"
public const t_dot_formula_non_numeric = "you are using a dot-formula of the form y ~ ., but your columns are not all numeric. " & _
                                         "xlkitlearn cannot automatically generate categorical variables for you when you use dot " & _
                                         "formulas."
public const t_incomplete_formula_rhs = "please complete your formula. the problematic formula was "
public const t_lowercase_c = "you are using a lowercase 'c' to convert a categorical variable. please change this " & _
                             "to an uppercase 'c'. the problematic term was "
public const t_formula_invalid_terms = "it looks like at least one term in your formula(s) does not exist in your data. " & _
                                       "the problematic terms were "
public const t_formula_invalid_terms_end = ". please remember these variable names are case sensitive."
public const t_formula_check_unknown_error = "unknown error. note that this might be a bug in the error-checking code rather than " & _
                                              "in your formula, so feel free to try and run it anyway and see if it works!"

' header check
public const t_space_header = "this column name contains a space"
public const t_reserved_header = "the name of this column is a 'reserved' python keyword; try something else"
public const t_num_start_header = "column names cannot start with a number"
public const t_duplicate_header = "this column name is repeated; column names can only appear once"
public const t_invalid_var_char = "variable names can only contain letters, numbers, and underscores"
public const t_bad_headers = "some of the headers in your data are invalid. click on the button to the right to view errors."

' validation
public const t_invalid_file_type = "xlkitlearn can only accept xls, xlsx, and csv files."
public const t_invalid_range = "please enter a valid range."
public const t_non_contiguous_range = "please select a contiguous range."
public const t_book_unsaved = "please save this workbook before running the addin."
public const t_onedrive = "it looks like you saved this file in a onedrive folder. unfortunately, onedrive is a truly abysmal piece of software, " & _
                                "and it interferes with the way python " & _
                                "communicates with excel workbooks. it therefore cannot be used with xlkitlearn." & vbcrlf & vbcrlf & _
                                "while you are using this workbook, please move this file to a folder outside onedrive (google drive and dropbox are " & _
                                "ok)."

' other
public const t_no_files = "the folder in which this file is located does not contain any csv, xls, or xlsx files; perhaps " & _
                          "you didn't download them into the right folder?"

' constants - other
' =================

' models
public const models = "{'linear/logistic regression'|'lasso penalty'`'k-nearest neighbors'|'neighbors, weighting'`'decision tree'|'tree depth'`'boosted decision tree'|'tree depth,max trees,learning rate'`'random forest'|'tree depth,number of trees'}"

' blank settings
public const blank_settings = "{'model'|'linear/logistic regression'`'formula'|''`'param1'|''`'param2'|''`'param3'|''`'training_data'|''`'k'|''`'ts_data'|'false'`'evaluation_perc'|''`'evaluation_data'|''`'prediction_data'|''`'seed'|'123'`'output_model'|'true'`'output_evaluation_details'|'true'`'output_code'|'true'}"
public const blank_text_settings = "{'source_data'|''`'max_df'|''`'min_df'|''`'max_features'|'500'`'stop_words'|'true'`'tf_idf'|'false'`'wordtovec'|'false'`'lda_topics'|''`'seed'|'123'`'eval_perc'|'0'`'bigrams'|'false'`'stem'|'false'`'output_code'|'true'`'sparse_output'|'false'`'max_lda_iter'|''}"

' sample settings
public const sample_settings = "{'model'|'random forest'`'formula'|'median_property_value ~ crime_per_capita + prop_zoned_over_25k + prop_non_retail_acres + bounds_river + nox_concentration + av_rooms_per_dwelling + prop_owner_occupied_pre_1940 + dist_to_employment_ctr + highway_accessibility + tax_rate + pupil_teacher_ratio'`'param1'|'3 & 4 & 5 & 6'`'param2'|'25'`'param3'|''`'training_data'|'[xlkitlearn.xlsm]boston_housing!$a$1:$l$507'`'k'|'5'`'ts_data'|'false'`'evaluation_perc'|''`'evaluation_data'|''`'prediction_data'|''`'seed'|'123'`'output_model'|'true'`'output_evaluation_details'|'true'`'output_code'|'true'}"

' the maximum number of columns we want to read column names for
public const max_col_headers = 100

' ephemeral variables
' ===================

' the default excel save format
public save_format as variant

' whether we are in debug mode
public debug_mode as boolean

' create variables we will use to indicate whether there are any errors in the predictive
' settings or the text settings; these will be used to warn the user before they run the
' add-in in case errors are present
public pred_errors as boolean
public text_errors as boolean

' a variable to store the sheet name the user has requested
public requested_sheet_name as string

' create a variable to store the add-in start time, so we can output time diagnostics
public start_time as long

' create a variable to store the calculation mode so it can be restored
public calc_mode as integer

' create a variable to store the window the file selection box was called from,
' so that we know where to put the result
public file_selection_source as string

' variables mediating the identification
' of columns in the data
' --------------------------------------

' collections for the variable names and types
public vars as new collection
public var_types as new collection

' collection for any errors in the header names
public header_errors as new collection

' whether the dataset selected contains any blank columns
public blank_cols as boolean

' whether we have an excel file, a file with too many columns
' or a file from which headers could not be read
public excel_file as boolean
public too_many_cols as boolean
public unreadable_headers as boolean

' permanent variables
' ===================

public function wb_var(var_name as string, optional value as string = missing)
    ' this function loads or sets a permanent variable in the workbook.
    ' these are stored at the top of the code_text sheet
    
    ' find the cell for the permanent variable
    dim i as integer
    i = 0
    
    while (range(cell_perm_vars).offset(0, i).value <> var_name)
        if trim(range(cell_perm_vars).offset(0, i).value) = "" then
            msgbox "attempted to access name " & var_name & " which does not exist.", vbcritical
            exit function
        end if
    
        i = i + 2
    wend
    
    ' if we have no value, return the current one
    if value = missing then
        wb_var = range(cell_perm_vars).offset(0, i + 1).value
    else
        range(cell_perm_vars).offset(0, i + 1).value = value
    end if
    
end function

' xlwings conf
' ============

public sub set_xlwings_conf(param, value)
    ' this function sets the value of xlwings settings on xlwings.conf
    
    with sheets("xlwings.conf")
        dim row as integer
        row = 1
        
        while .range("a" & row).value <> ""
            if .range("a" & row) = param then
                .range("b" & row) = value
                exit sub
            end if
            
            row = row + 1
        wend
        
        ' if we haven't found the param, add it
        .range("a" & row) = param
        .range("b" & row) = value
    end with
end sub
