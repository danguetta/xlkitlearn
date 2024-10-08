option explicit

' ###############################################################
' #  mdl_sheet_code                                             #
' #                                                             #
' #  contains the code that runs the buttons on the add-in      #
' #  sheet                                                      #
' ###############################################################

dim sheet_names as new collection
dim sheet_settings as new collection
dim running_batch_process as boolean

' function to check whether we have a onedrive file
' =================================================

function check_onedrive() as string
    if thisworkbook.path = "" then
        check_onedrive = t_book_unsaved
        exit function
    end if
    
    on error goto onedriveerror
    if dir(thisworkbook.fullname) = "" then
        goto onedriveerror
        exit function
    end if
    
    on error goto 0
    
    exit function
    
onedriveerror:
    check_onedrive = t_onedrive
end function


' functions to enable and disable buttons
' =======================================

sub disable_buttons()
    thisworkbook.sheets("add-in").buttons("btn_run").font.color = grey
    thisworkbook.sheets("add-in").buttons("btn_run_text").font.color = grey
    
    thisworkbook.sheets("add-in").buttons("btn_edit").font.color = grey
    thisworkbook.sheets("add-in").buttons("btn_edit_text").font.color = grey
end sub

sub enable_buttons()
    thisworkbook.sheets("add-in").buttons("btn_run").font.color = 0
    thisworkbook.sheets("add-in").buttons("btn_run_text").font.color = 0
    
    thisworkbook.sheets("add-in").buttons("btn_edit").font.color = 0
    thisworkbook.sheets("add-in").buttons("btn_edit_text").font.color = 0
end sub

' settings buttons
' ================

sub btn_edit_settings()
    ' if code isn't running, show settings
    if run_id() = "" and validate_settings_string("predictive", true) then
        ' ensure this workbook is saved
        if check_onedrive() <> "" then
            msgbox check_onedrive(), vbexclamation
            exit sub
        end if
        
        frm_pred.show false
    end if
end sub

sub btn_edit_text_settings()
    ' if code isn't running, show settings
    if run_id() = "" and validate_settings_string("text", true) then
        ' ensure this workbook is saved
        if check_onedrive() <> "" then
            msgbox check_onedrive(), vbexclamation
            exit sub
        end if
    
        frm_text.show false
    end if
end sub

public function validate_settings_string(which_addin as string, allow_blank as boolean) as boolean
    ' this function checks whether the settings string is correctly formatted;
    ' if allow_blank is true, a template empty settings strings will be copied
    ' to the settings cell
    
    dim settings_cell as string
    dim blank_settings_string as string
    
    if which_addin = "predictive" then
        settings_cell = cell_current_settings
        blank_settings_string = blank_settings
    elseif which_addin = "text" then
        settings_cell = cell_current_text_settings
        blank_settings_string = blank_text_settings
    end if
    
    validate_settings_string = false
    if trim(range(settings_cell).value) = "" then
        if allow_blank then
            range(settings_cell).value = blank_settings_string
            validate_settings_string = true
        else
            msgbox "it look like your settings string (cell " & split(settings_cell, "!")(1) & ") is blank; please enter some settings and try again.", vbcritical
        end if
    elseif ubound(dict_utils(settings_cell)) = -1 then
        msgbox "it look like you messed with your settings string (cell " & split(settings_cell, "!")(1) & "); try clearing it and starting again.", vbcritical
    else
        validate_settings_string = true
    end if
    
end function

' run buttons
' ===========

sub btn_run_addin()
    ' if the add-in is not running, run it
    if run_id() = "" and validate_settings_string("predictive", false) then
        ' if any entries are missing from the dictionary, sync it
        range(cell_current_settings).value = sync_dicts(range(cell_current_settings).value, blank_settings)
                
        ' run
        run_addin "predictive_addin", cell_status
    end if
end sub

sub btn_run_text_addin()
    ' if the add-in is not running, run it
    if run_id() = "" and validate_settings_string("text", false) then
        ' if any entries are missing from the dictionary, sync it
        range(cell_current_text_settings).value = sync_dicts(range(cell_current_text_settings).value, blank_text_settings)
        
        ' run
        run_addin "text_addin", cell_status_text
    end if
end sub

sub run_addin(f_name as string, this_status_cell as string)
    ' this function runs the add-in
    
    ' ensure this workbook is saved
    if check_onedrive() <> "" then
        msgbox check_onedrive(), vbexclamation
        exit sub
    end if

    ' load code if it's here and if we're not in production
    if debug_mode = true then
        on error resume next
        load_code
        on error goto 0
    end if

    ' ensure an email was provided
    dim email_provided as string
    email_provided = range(cell_email).value
    if (trim(email_provided) = "") or (instr(1, email_provided, "@") = 0) or (instr(1, email_provided, ".") = 0) or (instr(1, email_provided, """") > 0) or (instr(1, email_provided, "|") > 0) then
        msgbox "please enter a valid email address on the addin tab to validate the add-in before it runs.", vbexclamation, "validation error"
        format_sheet
        exit sub
    end if
    
    ' ensure there are no settings errors
    ' ===================================
    
    ' first, load and unload each form to check settings
    if f_name = "predictive_addin" then
        load frm_pred
        pause_execution 0.1
        unload frm_pred
    end if
    
    if f_name = "text_addin" then
        load frm_text
        pause_execution 0.1
        unload frm_text
    end if
    
    ' if there are errors, warn
    if (f_name = "predictive_addin" and pred_errors = true) or (f_name = "text_addin" and text_errors = true) then
        dim msg_res as integer
        msg_res = msgbox("it looks like some of your settings contain errors. these will be highlighted in red in the " & _
                           "settings dialogue when you click 'edit settings'." & vbcrlf & vbcrlf & "would you like to run the add-in anyway? " & _
                           "click 'yes' to run, and click 'no' to abort this run and launch the settings dialogue to see " & _
                           "where the errors are.", vbyesno + vbexclamation, "settings errors")
        
        if msg_res = vbno then
            if f_name = "predictive_addin" then
                frm_pred.show false
            else
                frm_text.show false
            end if
            
            exit sub
        end if
    end if
    
    ' generate a run id
    dim email_stem as long
    on error resume next
    email_stem = asc(mid(email_provided, 1, 1)) + asc(mid(email_provided, 2, 1))
    on error goto 0
    run_id true, email_stem
    
    ' disable all buttons
    disable_buttons
    
    ' set the status cell to "launching"
    thisworkbook.sheets("add-in").range(this_status_cell) = "launching python"
    doevents
        
    ' log the start time
    start_time = timer()
    
    ' create a new sheet to output the results - save the sheet's name
    dim new_sheet_name as string
    new_sheet_name = sheets.add().name
    activewindow.displaygridlines = false
    thisworkbook.sheets.item(new_sheet_name).visible = xlsheethidden
    
    ' rename the sheet if requested by the user
    on error goto after_rename
    if requested_sheet_name <> "" then
        thisworkbook.sheets.item(new_sheet_name).name = requested_sheet_name
        new_sheet_name = requested_sheet_name
    end if
    
after_rename:
    requested_sheet_name = ""
    on error goto 0
    
    ' save the name of the sheet we're using
    wb_var "sheet_in_progress", new_sheet_name
    
    ' save the calculation mode, and then set it to manual
    calc_mode = application.calculation
    application.calculation = xlcalculationmanual
    
    ' figure out whether we're using the udf
    dim udf_server as string
    udf_server = "false"
    #if mac then
        doevents
    #else
        if sheets("add-in").checkboxes("chk_server").value = 1 then
            set_xlwings_conf "use udf server", "true"
        else
            set_xlwings_conf "use udf server", "false"
        end if
        
        if sheets("add-in").checkboxes("chk_server").value = 1 or sheets("add-in").checkboxes("chk_foreground").value = 1 then
            set_xlwings_conf "show console", "true"
        else
            set_xlwings_conf "show console", "false"
        end if
    #end if
    
    ' run
    dim code_text as string
    code_text = "import xlwings as xw; " & _
                "import types; " & _
                "mod = types.moduletype('mod'); " & _
                "xw.book.caller().sheets('add-in').range('" & this_status_cell & "').value = 'python launched; loading packages'; " & _
                "code_sheet = xw.book.caller().sheets('code_text'); " & _
                "code_range = 'a2:a' + str(code_sheet.range(f'a{code_sheet.cells.last_cell.row}').end('up').row); " & _
                "exec('\n'.join([str(i) if i is not none else '' for i in code_sheet.range(code_range).value]), mod.__dict__); " & _
                "mod.run_addin('" & f_name & "','" & new_sheet_name & "', " & udf_server & ")"
    
    'log the run on the server
    on error resume next
    
    dim platform as string
    #if mac then
        platform = "mac"
    #else
        platform = "windows"
    #end if
    
    make_request "post", "https://telemetry.xlkitlearn.com/log.php", false, _
                        "request_type", "run", _
                        "run_id", run_id(), _
                        "version", addin_version(), _
                        "email", email_provided, _
                        "platform", platform
    
    on error goto 0
    
    dim run_python_result as string
    run_python_result = runpython(code_text)
    
    ' if the python code failed to terminate, run format_sheet to clear the extraneous
    ' sheet that will have been created and reset the buttons
    on error resume next
    #if mac then
        if run_python_result <> "" then
            format_sheet
        end if
    #else
        if run_python_result <> 0 then
            format_sheet
        end if
    #end if
    
    on error goto 0
end sub

' kill add-in button (mac only)
' =============================

public sub kill_addin()
    dim code_text as string
    
    code_text = "import os\n" & _
                "import xlwings as xw\n" & _
                "import signal\n" & _
                "try:\n" & _
                "    os.kill(int(""" & wb_var("pid") & """), signal.sigterm)\n" & _
                "    xw.book.caller().macro(""format_sheet"")()\n" & _
                "except:\n" & _
                "    pass"
    
    code_text = "exec('" & code_text & "')"
    
    ' kill_code = replace(kill_code, "'", "\'")
    
    runpython code_text
end sub

' format the sheet
' ================

public sub format_sheet()
    ' this function will be called from python when the add-in is done running;
    ' it will format the sheet and reset the add-in

    ' only activate workbook if necessary; on a mac, it can sometimes
    ' cause the sheet to change
    if activeworkbook.name <> thisworkbook.name then
        thisworkbook.activate
    end if

    ' re-enable the buttons
    enable_buttons
    
    ' clear status cells
    thisworkbook.sheets("add-in").range(cell_status).value = ""
    thisworkbook.sheets("add-in").range(cell_status_text).value = ""

    ' clear any lingering blank sheets
    dim sheet as worksheet
    application.displayalerts = false
    for each sheet in thisworkbook.sheets
        if worksheetfunction.counta(sheet.usedrange) = 0 and sheet.visible = 0 and sheet.name <> "xlwings.conf" then
            sheet.delete
        end if
    next sheet
    application.displayalerts = true
    
    ' if we don't have a sheet in progress, leave
    if (wb_var("sheet_in_progress") = "") or (not sheetexists(thisworkbook, wb_var("sheet_in_progress"))) then
        wb_var "sheet_in_progress", ""
        wb_var "run_id", ""
        exit sub
    end if
    
    ' if we have no output for some reason, delete the sheet
    if worksheetfunction.counta(thisworkbook.sheets(wb_var("sheet_in_progress")).usedrange) = 0 then
        application.displayalerts = false
        thisworkbook.sheets(wb_var("sheet_in_progress")).delete
        application.displayalerts = true
        wb_var "run_id", ""
        exit sub
    end if
    
    ' format the results
    ' ==================
    ' the top rows will contain cells that need special formating
    ' treatment.
    dim font_medium, font_large, bottom_thick, top_thin, italics, align_center, align_right, bold, expand, courier, align_left, graph_formatting, number_format
    dim num_cols as integer
    dim num_rows as integer
    dim i as integer
    
    with thisworkbook.sheets(wb_var("sheet_in_progress"))
        font_medium = .range("a1").value
        font_large = .range("a2").value
        bottom_thick = .range("a3").value
        top_thin = .range("a4").value
        italics = .range("a5").value
        bold = .range("a6").value
        align_center = .range("a7").value
        align_right = .range("a8").value
        expand = .range("a9").value
        courier = .range("a10").value
        align_left = .range("a11").value
        number_format = .range("a12").value
        num_cols = .range("a13").value
        num_rows = .range("a14").value
        
        graph_formatting = .range("a15").value
        
        ' delete these rows
        .rows("1:15").delete
    
        if font_medium <> "" then .range(font_medium).font.size = 16
        if font_large <> "" then .range(font_large).font.size = 22
        
        if bottom_thick <> "" then .range(bottom_thick).borders(xledgebottom).linestyle = xlcontinuous
        if bottom_thick <> "" then .range(bottom_thick).borders(xledgebottom).weight = xlmedium
            
        if top_thin <> "" then .range(top_thin).borders(xledgetop).linestyle = xlcontinuous
        if top_thin <> "" then .range(top_thin).borders(xledgetop).weight = xlthin
        
        if italics <> "" then .range(italics).font.italic = true
        
        if bold <> "" then .range(bold).font.bold = true
        
        if align_center <> "" then .range(align_center).horizontalalignment = xlcenter
        
        if align_right <> "" then .range(align_right).horizontalalignment = xlright
        
        if number_format <> "" then .range(number_format).numberformat = "0.000"
        
        if courier <> "" then .range(courier).font.name = "courier new"
        
        if align_left <> "" then .range(align_left).horizontalalignment = xlleft
        
        if expand <> "" then .range(expand).columns.autofit
    
        for i = 1 to num_cols
            if .columns(i).columnwidth > 20 then .columns(i).columnwidth = 20
        next i
    
        .visible = xlsheetvisible
        
        .activate
        
        if graph_formatting <> "" then
        
            graph_formatting = split(graph_formatting, "|")
            
            for i = 0 to ubound(graph_formatting)
                dim this_instruction
                this_instruction = split(graph_formatting(i), ",")
                
                dim inst_2 as double
                dim inst_3 as double
                inst_2 = localize_number(this_instruction(2))
                inst_3 = localize_number(this_instruction(3))
                
                with .shapes.range(this_instruction(0))
                    .top = localize_number(range(this_instruction(1)).top)
                    .left = localize_number(range(this_instruction(1)).left)
                    
                    ' we use a value of 3 lines per inch in the graph
                    .width = localize_number(range(this_instruction(1), range(this_instruction(1)).offset(3 * inst_2 - 1)).height)
                    .height = localize_number(range(this_instruction(1), range(this_instruction(1)).offset(3 * inst_2 - 1)).height) * inst_3 / inst_2
                end with
            next i
        
        end if
    
        ' output the overhead time
        if .range("a1").value <> "text add-in error report" and .range("a1").value <> "add-in error report" then
            .range("d" & num_rows).value = round(timer - start_time - localize_number(.range("d" & num_rows).value), 2)
        end if
    end with
    
    ' reset the calculation mode
    if calc_mode <> 0 then
        application.calculation = calc_mode
    else
        application.calculation = xlcalculationautomatic
    end if
    calc_mode = 0
        
    ' log a success
    on error resume next
    
    make_request "post", "https://telemetry.xlkitlearn.com/log.php", false, _
                "request_type", "success", _
                "run_id", run_id()
    
    ' clear the run_id
    wb_var "run_id", ""
    wb_var "sheet_in_progress", ""
    
    on error goto 0
end sub

