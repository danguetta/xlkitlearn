option explicit

dim change_running as boolean

private sub worksheet_change(byval target as range)
    ' ensure that the email address doesn't appear as a hyperlink
    target.hyperlinks.delete
end sub

private sub worksheet_selectionchange(byval target as range)
    ' only allow the three editable cells to be selected on the add-in sheet
    ' try and figure out the closest such cell whenever the use clicks on
    ' anything else

    on error resume next
    
    if not change_running and not debug_mode then
        change_running = true
        
        if target.address = "$d$8" then
            range(cell_email).select
        elseif target.address = "$d$10" then
            range(cell_current_text_settings).select
        elseif target.address = "$d$13" then
            range(cell_current_settings).select
        elseif target.address = "$d$15" then
            range(cell_email).select
        elseif target.address = "$f$16" then
            range(cell_current_text_settings).select
        elseif target.address = "$f$18" then
            range(cell_current_settings).select
        elseif target.row <= 9 then
            range(cell_current_settings).select
        elseif target.row <= 14 then
            range(cell_current_text_settings).select
        else
            range(cell_email).select
        end if
        
        change_running = false
    end if
end sub