option explicit

' ###############################################################
' #  mdl_dict_utils                                             #
' #                                                             #
' #  this module contains utilities that create a vba version   #
' #  of a python dictionary, which we use to save settings      #
' ###############################################################

public function dict_utils(byval dict_name as string, optional byval key as string = missing, optional byval value as string = missing, optional byval reverse as boolean = false, optional byval error_on_missing_key as boolean = true)
    ' this function takes a dictionary in one of the hidden names in
    ' this spreadsheet and either modifies one of the keys or returns
    ' it.
    '
    ' it takes four arguments
    '   - the name of the dictionary, which can contain one of three
    '     things
    '       * the name of an excel "name"
    '       * an actual dictionary string (in which case, no value
    '         can be set)
    '       * a range containing a dictionary
    '   - the key we want to access. if this is not provided, we
    '     return an array with the keys in this dictionary
    '   - the value we want to replace it with. if this is not
    '     provided, the dictionary is left as-is, and the value
    '     corresponding to this key is returned
    '   - if reverse is included and true, the values will be used
    '     as keys, and vice versa
    '   - if error_on_missing_keys is included and false, attempts
    '     to retrieve a missing key will return the missing sentinel
    '     rather than throw an error (or just add the key to an
    '     existing dict)
    
    on error goto error_out
    
    dim cur_dict as string
    if dict_name = "" then
        cur_dict = "{}"
    elseif instr(dict_name, "{") <> 0 then
        ' an actual dictionary was passed. ensure no value was passed
        cur_dict = dict_name
    elseif check_valid_range(dict_name) = true then
        cur_dict = range(dict_name).value
    else
        ' we want to make sure we're not using names anymore; trigger an
        ' error if we get here
        on error goto 0
        debug.print 1 / 0
        
        ' we have a variable name - load it from the variable and
        ' remove the first equal sign
        cur_dict = mid(thisworkbook.names.item(dict_name).value, 2)
        
        ' remove the quotes if they are there
        if left(cur_dict, 1) = "'" then
            cur_dict = trim_ends(cur_dict)
        end if
    end if
    
    ' remove opening and closing braces
    cur_dict = trim_ends(cur_dict)
    
    ' we might need to construct a dictionary or list, either as the new
    ' dictionary to be output, or as a list of keys to return. prepare
    ' a string to hold that dict/list
    dim output_string as string
    
    ' if both the key and value are provided, output_string will need to
    ' contain the new dictionary, so we should initialize it with a
    ' brace
    if key <> missing and value <> missing then
        output_string = "{"
    end if
        
    ' go through the dictionary
    dim split_dict, i as integer, this_key as string, this_value as string, found_key as boolean
    
    split_dict = split(cur_dict, "`")
    
    for i = 0 to ubound(split_dict)
        this_key = trim_ends(split(split_dict(i), "|")(0))
        this_value = trim_ends(split(split_dict(i), "|")(1))
        
        if reverse = true then
            dim buffer as string
            buffer = this_key
            this_key = this_value
            this_value = buffer
        end if
        
        if key = missing and value = missing then
            ' we're just cataloguing the keys
            output_string = output_string & this_key & "`"
        else
            ' we're either looking to retrieve a key or set one
            if this_key = key then
                ' we're in the correct key - figure out what to do
                if value = missing then
                    ' we've found the value we wanted - return
                    dict_utils = this_value
                    exit function
                else
                    ' add it to the dictionary
                    output_string = output_string & "'" & key & "'|'" & value & "'`"
                end if
                
                ' note that we found our key
                found_key = true
            else
                ' if we're just looking for a key, don't do anything. otherwise, just
                ' leave this key as is in the output dictionary
                if value <> missing then
                    output_string = output_string & "'" & this_key & "'|'" & this_value & "'`"
                end if
            end if
        end if
    next i
    
    ' if we got here, it's for one of the following reasons
    '   - we're retrieving a key and it wasn't found
    '   - we're updating a dictionary
    '   - we want to return a list of keys
    
    ' first, check if the key wasn't found
    if key <> missing and value = missing then
        ' retrieving a key that wasn't found
        if error_on_missing_key then
            msgbox "coding error: attempting to retrieve non-existent key " & key, vbcritical
            goto error_out
        else
            dict_utils = missing
        end if
        
        exit function
    elseif key <> missing and value <> missing then
        ' updating a dictionary. make sure we've found our key, and
        ' add it if needed
        if found_key = false then
            if error_on_missing_key then
                msgbox "coding error: attempting to set non-existent key " & key, vbcritical
                goto error_out
            else
                ' add the key
                output_string = output_string & "'" & key & "'|'" & value & "'`"
            end if
        end if
    end if
    
    ' there will be an extra backtick at the end; remove it
    output_string = mid(output_string, 1, len(output_string) - 1)
    
    ' now, find out what we're trying to do
    if key <> missing and value <> missing then
        ' add a closing brace
        output_string = output_string & "}"
    
        ' save it to the workbook or return it
        if instr(dict_name, "{") <> 0 then
            dict_utils = output_string
        elseif check_valid_range(dict_name) = true then
            range(dict_name).value = output_string
        else
            thisworkbook.names.item(dict_name).value = output_string
        end if
    else
        ' we need to output our keys
        dict_utils = split(output_string, "`")
        exit function
    end if
    
    exit function
error_out:
    dict_utils = split("", ",")
end function

public function sync_dicts(byval new_dict as string, byval template_dict as string)
    ' this function takes two dictionaries new_dict and template_dict.
    ' it outputs a dictionary that has all the entries in template_dict;
    ' if new_dict has an entry with a given key, that value is used in the
    ' output, if not the value in template_dict will be used as a "default"
    
    dim keys
    dim i as integer
    dim output_dict as string
    dim this_val as string
    
    ' retrieve the keys from the template_dict
    keys = dict_utils(template_dict)
    
    ' loop through them and add each entry. start with a blank dict
    output_dict = "{}"
    for i = 0 to ubound(keys)
        this_val = dict_utils(new_dict, keys(i), , , false)
        if this_val = missing then
            output_dict = dict_utils(output_dict, keys(i), dict_utils(template_dict, keys(i)), , false)
        else
            output_dict = dict_utils(output_dict, keys(i), this_val, , false)
        end if
    next i
    
    sync_dicts = output_dict
end function