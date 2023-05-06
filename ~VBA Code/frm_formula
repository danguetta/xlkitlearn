dim autocomplete as boolean

private sub cmd_save_click()
    frm_pred.txt_formula.text = txt_formula_edit.text
    unload me
end sub

private sub txt_formula_edit_change()
    ' if autocomplete = true (which will be the case if the user hasn't just
    ' pressed backspace), attempt to autocomplete the variable the user is
    ' typing

    on error goto unknown_err
    
    if autocomplete then
        ' find the variable stub
        dim backlook_done as boolean
        dim cur_char as string
        dim var_stub as string
        dim var_match as string
        dim start_sel_pos as integer
        dim i as integer
        
        ' start at the current position in the text, and look backwards until we
        ' get to the end of that variable
        i = txt_formula_edit.selstart + 1
        while (not backlook_done) and (i > 1)
            i = i - 1
            cur_char = mid(txt_formula_edit.text, i, 1)
            if cur_char = "(" or cur_char = " " or cur_char = "+" or cur_char = ")" _
                    or cur_char = "~" or cur_char = "*" or i <= 1 then backlook_done = true
        wend
        
        ' if i = 1, we're at the very start of the string. if not, we will have
        ' gone one character back from the start of the word, so go one forward
        ' to get to the start of the word
        if i > 1 then i = i + 1
        
        ' if there is at least one character in the current variable name,
        ' autocomplete it. if not, populate the variable names in the list on
        ' the left
        if i <= txt_formula_edit.selstart then
            var_stub = mid(txt_formula_edit.text, i, txt_formula_edit.selstart - i + 1)
            var_match = autocomplete_variable(var_stub, true)
            
            if var_match <> "" then
                populate_list var_stub
                autocomplete = false
                
                start_sel_pos = txt_formula_edit.selstart
                txt_formula_edit.text = mid(txt_formula_edit.text, 1, txt_formula_edit.selstart) & _
                                          mid(var_match, len(var_stub) + 1) & _
                                          mid(txt_formula_edit.text, txt_formula_edit.selstart + 1)
                txt_formula_edit.selstart = start_sel_pos
                txt_formula_edit.sellength = len(var_match) - len(var_stub)
                
                autocomplete = true
            else
                populate_list
            end if
        else
            populate_list
        end if
    end if
    
    ' re-validate the formula
    dim error_message as string
    error_message = validate_formula(txt_formula_edit.text)
    
    goto after_err
unknown_err:
    error_message = t_formula_check_unknown_error
after_err:
    
    if error_message <> "" then
        txt_formula_edit.backcolor = red
        txt_error.text = error_message
    else
        txt_formula_edit.backcolor = white
        txt_error.text = ""
    end if
    
end sub

private sub txt_formula_edit_keydown(byval keycode as msforms.returninteger, byval shift as integer)
    ' if we're backspacing, do not autocomplete
    if keycode = 8 then
        autocomplete = false
        populate_list
    else
        autocomplete = true
    end if

    ' if the user tabs, treat this as a right arrow, which "accepts" the
    ' autocomplete suggestion. in vba, rightarrow is apparently keycode 39
    if keycode = 9 then
        keycode = 39
    end if
    
    ' implement paste for mac
    #if mac then
        if keycode = 86 and shift = 2 then
            keycode = 0
            txt_formula_edit.text = getclipboardtext()
        end if
    #end if
end sub

private sub userform_initialize()
    ' populate the variable names
    populate_list
    
    ' initialize the formula and check it
    txt_formula_edit.text = frm_pred.txt_formula.text
    
    dim error_message as string
    error_message = validate_formula(txt_formula_edit.text)
    
    if error_message <> "" then
        txt_formula_edit.backcolor = red
        txt_error.text = error_message
    else
        txt_formula_edit.backcolor = white
        txt_error.text = ""
    end if
    
    ' turn autocomplete on
    autocomplete = true

    ' put the window in the ocrrect place
    on error resume next
    me.startupposition = 0
    me.left = application.left + (0.5 * application.width) - (0.5 * me.width)
    me.top = application.top + (0.5 * application.height) - (0.5 * me.height)
end sub

private sub populate_list(optional stub as string = "")
    ' populate the list of variables; if stub is provided, only include variables
    ' that start with stub
    lst_variables.clear
    
    if mac_file then
        lst_variables.additem "mac security settings"
        lst_variables.additem "do not allow excel to"
        lst_variables.additem "access the column names"
        lst_variables.additem "from a file. you will"
        lst_variables.additem "need to open the file,"
        lst_variables.additem "look at the column"
        lst_variables.additem "names there, and type"
        lst_variables.additem "them manually into the"
        lst_variables.additem "formula."
    elseif invalid_file then
        lst_variables.additem "i've encoutered an error"
        lst_variables.additem "trying to find the"
        lst_variables.additem "columns in the file you"
        lst_variables.additem "selected. you'll need"
        lst_variables.additem "type columns here"
        lst_variables.additem "manually instead."
    elseif excel_file then
        lst_variables.additem "vba cannot find column"
        lst_variables.additem "names for non-csv files."
        lst_variables.additem "you will need to open"
        lst_variables.additem "the file, look at the"
        lst_variables.additem "column names there, and"
        lst_variables.additem "type them manually into"
        lst_variables.additem "the formula."
    elseif too_many_cols then
        lst_variables.additem "your data has a lot of"
        lst_variables.additem "columns! for the sake"
        lst_variables.additem "of not crashing your"
        lst_variables.additem "computer, i won't list"
        lst_variables.additem "them all here. you can"
        lst_variables.additem "manually type them into"
        lst_variables.additem "the formula."
    else
        dim i as integer
        for i = 1 to vars.count
            if stub = mid(vars.item(i), 1, len(stub)) then
                lst_variables.additem vars.item(i)
            end if
        next i
    end if
end sub