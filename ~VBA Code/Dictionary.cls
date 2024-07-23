''
' dictionary v1.4.1
' (c) tim hall - https://github.com/timhall/vba-dictionary
'
' drop-in replacement for scripting.dictionary on mac
'
' @author: tim.hall.engr@gmail.com
' @license: mit (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
option explicit

' --------------------------------------------- '
' constants and private variables
' --------------------------------------------- '

#const usescriptingdictionaryifavailable = true

#if mac or not usescriptingdictionaryifavailable then

' dict_keyvalue 0: formattedkey, 1: originalkey, 2: value
private dict_pkeyvalues as collection
private dict_pkeys() as variant
private dict_pitems() as variant
private dict_pobjectkeys as collection
private dict_pcomparemode as comparemethod

#else

private dict_pdictionary as object

#end if

' --------------------------------------------- '
' types
' --------------------------------------------- '

public enum comparemethod
    binarycompare = vba.vbbinarycompare
    textcompare = vba.vbtextcompare
    databasecompare = vba.vbdatabasecompare
end enum

' --------------------------------------------- '
' properties
' --------------------------------------------- '

public property get comparemode() as comparemethod
#if mac or not usescriptingdictionaryifavailable then
    comparemode = dict_pcomparemode
#else
    comparemode = dict_pdictionary.comparemode
#end if
end property
public property let comparemode(value as comparemethod)
#if mac or not usescriptingdictionaryifavailable then
    if me.count > 0 then
        ' can't change comparemode for dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        err.raise 5 ' invalid procedure call or argument
    end if

    dict_pcomparemode = value
#else
    dict_pdictionary.comparemode = value
#end if
end property

public property get count() as long
#if mac or not usescriptingdictionaryifavailable then
    count = dict_pkeyvalues.count
#else
    count = dict_pdictionary.count
#end if
end property

public property get item(key as variant) as variant
#if mac or not usescriptingdictionaryifavailable then
    dim dict_keyvalue as variant
    dict_keyvalue = dict_getkeyvalue(key)

    if not isempty(dict_keyvalue) then
        if vba.isobject(dict_keyvalue(2)) then
            set item = dict_keyvalue(2)
        else
            item = dict_keyvalue(2)
        end if
    else
        ' not found -> returns empty
    end if
#else
    if vba.isobject(dict_pdictionary.item(key)) then
        set item = dict_pdictionary.item(key)
    else
        item = dict_pdictionary.item(key)
    end if
#end if
end property
public property let item(key as variant, value as variant)
#if mac or not usescriptingdictionaryifavailable then
    if me.exists(key) then
        dict_replacekeyvalue dict_getkeyvalue(key), key, value
    else
        dict_addkeyvalue key, value
    end if
#else
    dict_pdictionary.item(key) = value
#end if
end property
public property set item(key as variant, value as variant)
#if mac or not usescriptingdictionaryifavailable then
    if me.exists(key) then
        dict_replacekeyvalue dict_getkeyvalue(key), key, value
    else
        dict_addkeyvalue key, value
    end if
#else
    set dict_pdictionary.item(key) = value
#end if
end property

public property let key(previous as variant, updated as variant)
#if mac or not usescriptingdictionaryifavailable then
    dim dict_keyvalue as variant
    dict_keyvalue = dict_getkeyvalue(previous)

    if not vba.isempty(dict_keyvalue) then
        dict_replacekeyvalue dict_keyvalue, updated, dict_keyvalue(2)
    end if
#else
    dict_pdictionary.key(previous) = updated
#end if
end property

' ============================================= '
' public methods
' ============================================= '

''
' add an item with the given key
'
' @param {variant} key
' @param {variant} item
' --------------------------------------------- '
public sub add(key as variant, item as variant)
#if mac or not usescriptingdictionaryifavailable then
    if not me.exists(key) then
        dict_addkeyvalue key, item
    else
        ' this key is already associated with an element of this collection
        err.raise 457
    end if
#else
    dict_pdictionary.add key, item
#end if
end sub

''
' check if an item exists for the given key
'
' @param {variant} key
' @return {boolean}
' --------------------------------------------- '
public function exists(key as variant) as boolean
#if mac or not usescriptingdictionaryifavailable then
    exists = not isempty(dict_getkeyvalue(key))
#else
    exists = dict_pdictionary.exists(key)
#end if
end function

''
' get an array of all items
'
' @return {variant}
' --------------------------------------------- '
public function items() as variant
#if mac or not usescriptingdictionaryifavailable then
    if me.count > 0 then
        items = dict_pitems
    else
        ' split("") creates initialized empty array that matches dictionary keys and items
        items = vba.split("")
    end if
#else
    items = dict_pdictionary.items
#end if
end function

''
' get an array of all keys
'
' @return {variant}
' --------------------------------------------- '
public function keys() as variant
#if mac or not usescriptingdictionaryifavailable then
    if me.count > 0 then
        keys = dict_pkeys
    else
        ' split("") creates initialized empty array that matches dictionary keys and items
        keys = vba.split("")
    end if
#else
    keys = dict_pdictionary.keys
#end if
end function

''
' remove an item for the given key
'
' @param {variant} key
' --------------------------------------------- '
public sub remove(key as variant)
#if mac or not usescriptingdictionaryifavailable then
    dim dict_keyvalue as variant
    dict_keyvalue = dict_getkeyvalue(key)

    if not vba.isempty(dict_keyvalue) then
        dict_removekeyvalue dict_keyvalue
    else
        ' application-defined or object-defined error
        err.raise 32811
    end if
#else
    dict_pdictionary.remove key
#end if
end sub

''
' remove all items
' --------------------------------------------- '
public sub removeall()
#if mac or not usescriptingdictionaryifavailable then
    set dict_pkeyvalues = new collection

    erase dict_pkeys
    erase dict_pitems
#else
    dict_pdictionary.removeall
#end if
end sub

' ============================================= '
' private functions
' ============================================= '

#if mac or not usescriptingdictionaryifavailable then

private function dict_getkeyvalue(dict_key as variant) as variant
    on error resume next
    dict_getkeyvalue = dict_pkeyvalues(dict_getformattedkey(dict_key))
    err.clear
end function

private sub dict_addkeyvalue(dict_key as variant, dict_value as variant, optional dict_index as long = -1)
    if me.count = 0 then
        redim dict_pkeys(0 to 0)
        redim dict_pitems(0 to 0)
    else
        redim preserve dict_pkeys(0 to ubound(dict_pkeys) + 1)
        redim preserve dict_pitems(0 to ubound(dict_pitems) + 1)
    end if

    dim dict_formattedkey as string
    dict_formattedkey = dict_getformattedkey(dict_key)

    if dict_index >= 0 and dict_index < dict_pkeyvalues.count then
        ' shift keys/items after + including index into empty last slot
        dim dict_i as long
        for dict_i = ubound(dict_pkeys) to dict_index + 1 step -1
            dict_pkeys(dict_i) = dict_pkeys(dict_i - 1)
            if vba.isobject(dict_pitems(dict_i - 1)) then
                set dict_pitems(dict_i) = dict_pitems(dict_i - 1)
            else
                dict_pitems(dict_i) = dict_pitems(dict_i - 1)
            end if
        next dict_i

        ' add key/item at index
        dict_pkeys(dict_index) = dict_key
        if vba.isobject(dict_value) then
            set dict_pitems(dict_index) = dict_value
        else
            dict_pitems(dict_index) = dict_value
        end if

        ' add key-value at proper index
        dict_pkeyvalues.add array(dict_formattedkey, dict_key, dict_value), dict_formattedkey, before:=dict_index + 1
    else
        ' add key-value as last item
        if vba.isobject(dict_key) then
            set dict_pkeys(ubound(dict_pkeys)) = dict_key
        else
            dict_pkeys(ubound(dict_pkeys)) = dict_key
        end if
        if vba.isobject(dict_value) then
            set dict_pitems(ubound(dict_pitems)) = dict_value
        else
            dict_pitems(ubound(dict_pitems)) = dict_value
        end if

        dict_pkeyvalues.add array(dict_formattedkey, dict_key, dict_value), dict_formattedkey
    end if
end sub

private sub dict_replacekeyvalue(dict_keyvalue as variant, dict_key as variant, dict_value as variant)
    dim dict_index as long
    dim dict_i as integer

    dict_index = dict_getkeyindex(dict_keyvalue(1))

    ' remove existing dict_value
    dict_removekeyvalue dict_keyvalue, dict_index

    ' add new dict_key dict_value back
    dict_addkeyvalue dict_key, dict_value, dict_index
end sub

private sub dict_removekeyvalue(dict_keyvalue as variant, optional byval dict_index as long = -1)
    dim dict_i as long
    if dict_index = -1 then
        dict_index = dict_getkeyindex(dict_keyvalue(1))
    end if

    if dict_index >= 0 and dict_index <= ubound(dict_pkeys) then
        ' shift keys/items after index down
        for dict_i = dict_index to ubound(dict_pkeys) - 1
            dict_pkeys(dict_i) = dict_pkeys(dict_i + 1)

            if vba.isobject(dict_pitems(dict_i + 1)) then
                set dict_pitems(dict_i) = dict_pitems(dict_i + 1)
            else
                dict_pitems(dict_i) = dict_pitems(dict_i + 1)
            end if
        next dict_i

        ' resize keys/items to remove empty slot
        if ubound(dict_pkeys) = 0 then
            erase dict_pkeys
            erase dict_pitems
        else
            redim preserve dict_pkeys(0 to ubound(dict_pkeys) - 1)
            redim preserve dict_pitems(0 to ubound(dict_pitems) - 1)
        end if
    end if

    dict_pkeyvalues.remove dict_keyvalue(0)
    dict_removeobjectkey dict_keyvalue(1)
end sub

private function dict_getformattedkey(dict_key as variant) as string
    if vba.isobject(dict_key) then
        dict_getformattedkey = dict_getobjectkey(dict_key)
    elseif vartype(dict_key) = vba.vbboolean then
        dict_getformattedkey = iif(dict_key, "-1__-1", "0__0")
    elseif vartype(dict_key) = vba.vbstring then
        dict_getformattedkey = dict_key

        if me.comparemode = comparemethod.binarycompare then
            ' collection does not have method of setting key comparison
            ' so case-sensitive keys aren't supported by default
            ' -> approach: append lowercase characters to original key
            '    abc -> abc___b_, abc -> abc__abc, abc -> abc_____
            dim dict_lowercase as string
            dict_lowercase = ""

            dim dict_i as integer
            dim dict_char as string
            dim dict_ascii as integer
            for dict_i = 1 to vba.len(dict_getformattedkey)
                dict_char = vba.mid$(dict_getformattedkey, dict_i, 1)
                dict_ascii = vba.asc(dict_char)
                if dict_ascii >= 97 and dict_ascii <= 122 then
                    dict_lowercase = dict_lowercase & dict_char
                else
                    dict_lowercase = dict_lowercase & "_"
                end if
            next dict_i

            if dict_lowercase <> "" then
                dict_getformattedkey = dict_getformattedkey & "__" & dict_lowercase
            end if
        end if
    else
        ' for numbers, add duplicate to distinguish from strings
        ' ->  123  -> "123__123"
        '    "123" -> "123"
        dict_getformattedkey = vba.cstr(dict_key) & "__" & cstr(dict_key)
    end if
end function

private function dict_getobjectkey(dict_objkey as variant) as string
    dim dict_i as integer
    for dict_i = 1 to dict_pobjectkeys.count
        if dict_pobjectkeys.item(dict_i) is dict_objkey then
            dict_getobjectkey = "__object__" & dict_i
            exit function
        end if
    next dict_i

    dict_pobjectkeys.add dict_objkey
    dict_getobjectkey = "__object__" & dict_pobjectkeys.count
end function

private sub dict_removeobjectkey(dict_objkey as variant)
    dim dict_i as integer
    for dict_i = 1 to dict_pobjectkeys.count
        if dict_pobjectkeys.item(dict_i) is dict_objkey then
            dict_pobjectkeys.remove dict_i
            exit sub
        end if
    next dict_i
end sub

private function dict_getkeyindex(dict_key as variant) as long
    dim dict_i as long
    for dict_i = 0 to ubound(dict_pkeys)
        if vba.isobject(dict_pkeys(dict_i)) and vba.isobject(dict_key) then
            if dict_pkeys(dict_i) is dict_key then
                dict_getkeyindex = dict_i
                exit for
            end if
        elseif vba.isobject(dict_pkeys(dict_i)) or vba.isobject(dict_key) then
            ' both need to be objects to check equality, skip
        elseif dict_pkeys(dict_i) = dict_key then
            dict_getkeyindex = dict_i
            exit for
        end if
    next dict_i
end function

#end if

private sub class_initialize()
#if mac or not usescriptingdictionaryifavailable then
    set dict_pkeyvalues = new collection

    erase dict_pkeys
    erase dict_pitems
    set dict_pobjectkeys = new collection
#else
    set dict_pdictionary = createobject("scripting.dictionary")
#end if
end sub

private sub class_terminate()
#if mac or not usescriptingdictionaryifavailable then
    set dict_pkeyvalues = nothing
    set dict_pobjectkeys = nothing
#else
    set dict_pdictionary = nothing
#end if
end sub
