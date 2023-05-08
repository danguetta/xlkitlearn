option explicit

public sub workbook_open()
    ' immediately clear the prod flag; if we're re-opening we're not longer in
    ' in prod mode and we need to run prepare_for_prod again if we want to publish
    wb_var "prepped_for_prod", ""
    
    ' check whether we're in debug mode and hide sheets accordingly
    check_debug_mode
    
    ' re-set the add-in
    format_sheet
    
    ' check whether python is correctly set up
    #if mac then
        #if mac_office_version < 15 then
            msgbox "you are using excel 2011 for mac or earlier. unfortunately, xlkitlearn was only " & _
                    "tested in excel 2016 or later. ", vbcritical + vbretrycancel + vbdefaultbutton2
            
            goto close_workbook
        #end if
        
        ' ensure the interpreter exists
        if not fileorfolderexistsonmac(parentfolder(parentfolder(getconfig("interpreter_mac", "")))) then
            msgbox "it looks like you are trying to run the xlkitlearn excel file without first installing the python interpreter " & _
                    "required. please seek help.", vbcritical
            
            goto close_workbook
        end if
        
        'deal with the advanced commands
        sheets("add-in").buttons("btn_kill").visible = true
        
        sheets("add-in").checkboxes("chk_server").visible = false
        sheets("add-in").checkboxes("chk_foreground").visible = false
        
    #else
        ' ensure the interpreter exists
        if not fileexists(getconfig("interpreter_win", "")) then
            msgbox "it looks like you are trying to run the xlkitlearn excel file without first installing the python interpreter " & _
                    "required. please seek help.", vbcritical
                    
            goto close_workbook
        end if
                
        ' deal with the advanced commands
        sheets("add-in").buttons("btn_kill").visible = false
        
        sheets("add-in").checkboxes("chk_server").visible = true
        sheets("add-in").checkboxes("chk_foreground").visible = true
    #end if
    
    ' check the version of the interpreter matches
    on error goto version_check_error
    
    dim version_n as string
    #if mac then
        version_n = readfile(parentfolder(parentfolder(getconfig("interpreter_mac", ""))) & "/data/version")
    #else
        version_n = readfile(parentfolder(getconfig("interpreter_win", "")) & "\data\version")
    #end if
    
    version_n = replace(replace(version_n, chr(10), ""), chr(13), "")
    
    if version_n <> addin_version() then
        if version_n = "" then
            version_n = "unknown"
        end if
        msgbox "it looks like something has gone awry with your xlkitlearn installation. your python interpreter and your " & _
               "xlkitlearn file are mismatched. the interpreter version is " & version_n & ", and the version of this excel is " & _
               addin_version() & ". please seek help.", vbcritical
        
        goto close_workbook
    end if
    
    ' prompt to save the notebook
    if thisworkbook.path = "" then
        #if mac then
            ' do events to hide the checkboxes and show the buttons immediately, before
            ' the user presses on "save"
            doevents
            activeworkbook.refreshall
        
            ' on a pc, the beforesave event triggers reliably, and we can use it to
            ' set the default save format to a macro enabled workbook. on a mac, the
            ' event sometimes fails to triggers, so we have to change the default
            ' save type here
            
            ' save the current default save type
            save_format = application.defaultsaveformat
            
            'set the default save format to 52 (macro enabled workbook)
            application.defaultsaveformat = 52
            
            msgbox t_welcome_text_mac, vbinformation
        #else
            msgbox t_welcome_text_pc, vbinformation
        #end if
    end if
    
    exit sub
version_check_error:
    msgbox "i was unable to check whether your python interpreter matches your xlkitlearn version. please seek help.", vbcritical
    
close_workbook:
    if debug_mode = false then
        thisworkbook.close false
    end if
end sub

private sub workbook_aftersave(byval success as boolean)
    ' this event triggers reliably on a pc *and* a mac; reset the save format
    ' back to what it was
    
    ' reset the default save format
    application.defaultsaveformat = save_format
    
    ' check whether we should be in debug mode
    check_debug_mode
    
    ' reset the add-in
    format_sheet
    
    #if mac then
        ' the current version of xlwings isn't able to detect paths when using onedrive
        ' for business. get_local_path is. until xlwings integrates that functionality,
        ' we try and do it here
        
        dim full_name as string
        dim local_name as string
        dim root as string
        
        full_name = thisworkbook.fullname
        local_name = getlocalpath(thisworkbook.fullname)
        
        if instr(1, full_name, "my.sharepoint") then
            on error goto one_drive_error
            
            ' remove everything after my.sharepoint
            full_name = mid(full_name, instr(1, full_name, "my.sharepoint") + len("my.sharepoint"))
            
            ' find the fourth slash, and remove everything before
            full_name = mid(full_name, instr(instr(instr(instr(1, full_name, "/") + 1, full_name, "/") + 1, full_name, "/") + 1, full_name, "/") + 1)
            
            ' remove that tail from the local_name to find the stem
            root = left(local_name, instr(1, local_name, full_name) - 2)
            
            ' put the stem in the conf
            set_xlwings_conf "onedrive_commercial_mac", root
            
            exit sub
one_drive_error:
            msgbox "xlkitlearn was not able to figure out where the onedrive folder is stored on your computer. please save this file somewhere outside onedrive before using it.", vbexclamation
        end if
    #end if
end sub

private sub workbook_beforesave(byval saveasui as boolean, cancel as boolean)
    ' we can only rely on this event triggering reliably on windows because
    ' of an excel bug, so make this explicit
    #if not mac then
        ' record the current "default save" format
        save_format = application.defaultsaveformat
        
        'set the default save format to 52 (macro enabled workbook)
        application.defaultsaveformat = 52
    #end if
end sub

private sub workbook_beforeclose(cancel as boolean)
    ' on a mac, we had to set the default save format when we opened the
    ' notebook. this *should* have been reset in the aftersave event, but
    ' on the off-chance the notebook wasn't saved, let's reset it here
    
    #if mac then
        application.defaultsaveformat = save_format
    #end if
end sub
