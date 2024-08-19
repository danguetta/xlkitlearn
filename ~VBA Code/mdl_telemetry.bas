option explicit

' ###############################################################
' #  mdl_telemetry                                              #
' #                                                             #
' #  contains the code that will be used to log data to a       #
' #  server                                                     #
' ###############################################################

private declare ptrsafe function web_popen lib "libc.dylib" alias "popen" (byval command as string, byval mode as string) as longptr
private declare ptrsafe function web_pclose lib "libc.dylib" alias "pclose" (byval file as longptr) as long
private declare ptrsafe function web_fread lib "libc.dylib" alias "fread" (byval outstr as string, byval size as longptr, byval items as longptr, byval stream as longptr) as long
private declare ptrsafe function web_feof lib "libc.dylib" alias "feof" (byval file as longptr) as longptr


public function make_request(request_type as string, url as string, capture_output as boolean, paramarray args() as variant) as string
    ' sends a request from vba in a cross/platform fashion. accepts the following
    ' arguments
    '   - request_type : accepts get or post
    '   - url : the url
    '   - capture_output : true if we should wait for the result from the api and
    '     return it, false, if we should run asynchronously and return without
    '     waiting for the result
    '   - optionally, an arbitrary number of arguments containing data that should
    '     be sent with the request. these should come in pairs - the argument name,
    '     followed by a value
    '
    ' the request will be made using the curl commandon pc and mac.
    '
    ' if any data arguments are passed, they will be send using the -d argument in
    ' the following form
    '     {name:value, name:value, ...}
    '
    ' none of the arguments passed to this function can include double quotes or the
    ' | sign
    '
    ' if capture_output = true, the return value will be a ">" followed by the
    ' response from the api call. if there is an error, the return value will be "-1"
    '
    ' if capture_output = false, the return value will always be "-1", since we can
    ' check whether the request was successful
    
    make_request = -1
    
    on error goto abort_function
    
    ' ======================
    ' =   error checking   =
    ' ======================
    
    ' check the request type is valid
    if (request_type <> "post") and (request_type <> "get") then
        exit function
    end if
    
    ' ensure there are no forbidden symbols in the url
    if instr(1, url, """") + instr(1, url, "|") <> 0 then
        exit function
    end if
    
    ' ============================================
    ' =   create the data argument if required   =
    ' ============================================
    
    dim data_arg as string
    if ubound(args) - lbound(args) + 1 > 0 then
        ' make sure we have an even number of arguments
    
        if ((ubound(args) - lbound(args) + 1) mod 2) <> 0 then
            exit function
        end if
    
        dim i as integer
        for i = lbound(args) to ubound(args) step 2
            if instr(1, args(i), """") + instr(1, args(i), "|") + instr(1, args(i + 1), """") + instr(1, args(i + 1), "|") <> 0 then
                exit function
            end if
            data_arg = data_arg & "\""" & args(i) & "\"":\""" & replace(args(i + 1), """", "\""") & "\"","
        next i
        data_arg = mid(data_arg, 1, len(data_arg) - 1)
        data_arg = "{" & data_arg & "}"
    end if
    
    ' ===============================
    ' =   create the curl command   =
    ' ===============================
    dim curl_command as string
    
    curl_command = "curl -x """ & request_type & """ -h ""content-type:application/json"" "
    
    if data_arg <> "" then
        curl_command = curl_command & "-d """ & data_arg & """ "
    end if
    
    curl_command = curl_command & """" & url & """"
    
    ' ========================
    ' =   make the request   =
    ' ========================
    
    #if mac then
    
        if capture_output then
        
            dim web_file as longptr
            dim web_chunk as string
            dim web_read as long
    
            on error goto web_cleanup
    
            web_file = web_popen(curl_command, "r")
    
            if web_file = 0 then
                exit function
            end if
            
            make_request = ""
            do while web_feof(web_file) = 0
                web_chunk = vba.space$(50)
                web_read = web_fread(web_chunk, 1, len(web_chunk) - 1, web_file)
                if web_read > 0 then
                    web_chunk = vba.left$(web_chunk, web_read)
                    make_request = make_request & web_chunk
                end if
            loop
            
            ' prepend '>' to show an error hasn't happened
            make_request = ">" & make_request

web_cleanup:
            web_pclose (web_file)
            
        else
            
            ' use an applescript to run the call asynchronously
            applescripttask "xlkitlearn_dorequest.applescript", "dorequest", request_type & "|" & data_arg & "|" & url
            
        end if
    
    #else
    
        dim wsh as object
        
        if capture_output then
            dim exec as object
            dim output as string
            dim request as string
            
            'create the wscript.shell object
            set wsh = createobject("wscript.shell")
            
            ' execute the command and capture the output
            set exec = wsh.exec(curl_command)
            
            ' read the output
            make_request = exec.stdout.readall
            
            ' clean up
            set exec = nothing
            set wsh = nothing
            
            ' pre-pend ">"
            make_request = ">" & make_request
        else
            dim waitonreturn as boolean: waitonreturn = false
            dim windowstyle as integer: windowstyle = 0
            set wsh = createobject("wscript.shell")
            wsh.run curl_command, windowstyle, waitonreturn
            set wsh = nothing
        end if
        
    #end if
    
abort_function:
    
end function

public sub log_vba_error(byval content as string)
    ' log a vba error

    on error resume next
    
    dim platform as string
    #if mac then
        platform = "mac"
    #else
        platform = "windows"
    #end if
    
    ' remove forbiden characters from the contents
    content = replace(content, """", "{quote}")
    content = replace(content, "|", "{pipe}")
    content = replace(content, "\", "\\")
    content = replace(content, chr(10), "\n")      ' lf (\n)
    content = replace(content, chr(13), "")        ' cr (\r)
    
    make_request "post", "https://telemetry.xlkitlearn.com/log.php", false, _
                    "request_type", "error", _
                    "run_id", run_id(), _
                    "source", "unknown", _
                    "error_type", "vba_error", _
                    "platform", platform, _
                    "error_text", content
end sub