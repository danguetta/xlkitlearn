option explicit

' ###############################################################
' #  mdl_telemetry                                              #
' #                                                             #
' #  contains the code that will be used to log data to a       #
' #  server                                                     #
' ###############################################################

public sub windows_curl(request as string)
    ' this function runs a curl on windows
    
    ' replace spaces in the request with '%20' and double quotes with single quotes
    request = replace(replace(request, " ", "%20"), """", "'")
    
    ' add the curl to the front of the request, with double quotes around the request
    request = "curl """ & request & """"
    
    ' run without waiting for return
    
    dim waitonreturn as boolean: waitonreturn = false
    dim windowstyle as integer: windowstyle = 0
    dim wsh as object
    set wsh = createobject("wscript.shell")
    wsh.run request, windowstyle, waitonreturn
    set wsh = nothing
end sub

public sub log_vba_error(content as string)
    ' log a vba error

    on error resume next
    dim request as string
    
    #if mac then
        request = replace(content, "'", """")
        runpython "import requests; requests.post(url = 'http://guetta.org/addin/error.php'," & _
                        "data = {'run_id':'" & run_id() & "', " & _
                        "        'source':'unknown', 'error_type'='vba_error', 'platform'='mac', " & _
                        "        'error_text' : '" & request & "'}, timeout = 10)"
    #else
        request = replace(content, vbcrlf, "\n")
        request = replace(request, chr(10), "\n")
        request = replace(request, "&", "|")
        
        windows_curl "http://guetta.org/addin/error.php?run_id=" & run_id() & _
                           "&source=unknown&error_type=vba_error&platform=windows&error_text=" & request
    #end if

end sub