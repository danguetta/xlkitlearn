' ###############################################################
' #  mdl_onedrive_path                                          #
' #                                                             #
' #  contains function to resolve the true workbook path; see   #
' #  comment on function below                                  #
' ###############################################################

'attribute vb_name = "getlocalonedrivepath"
' cross-platform vba function to get the local path of onedrive/sharepoint
' synchronized microsoft office files (works on windows and on macos)
'
' author: guido witt-dörring
' created: 2022/07/01
' updated: 2023/05/02
' license: mit
'
' ————————————————————————————————————————————————————————————————
' https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
' https://stackoverflow.com/a/73577057/12287457
' ————————————————————————————————————————————————————————————————
'
' copyright (c) 2023 guido witt-dörring
'
' mit license:
' permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "software"), to
' deal in the software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the software, and to permit persons to whom the software is
' furnished to do so, subject to the following conditions:
'
' the above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the software.
'
' the software is provided "as is", without warranty of any kind, express or
' implied, including but not limited to the warranties of merchantability,
' fitness for a particular purpose and noninfringement. in no event shall the
' authors or copyright holders be liable for any claim, damages or other
' liability, whether in an action of contract, tort or otherwise, arising
' from, out of or in connection with the software or the use or other dealings
' in the software.
'
'———————————————————————————————————————————————————————————————————————————————
' comments regarding the implementation:
' 1) background and alternative
'    this function was intended to be written as a single procedure without any
'    dependencies, for maximum portability between projects, as it implements a
'    functionality that is very commonly needed for many vba applications
'    working inside onedrive/sharepoint synchronized directories. i followed
'    this paradigm because it was not clear to me how complicated this simple
'    sounding endeavour would turn out to be.
'    unfortunately, more and more complications arose, and little by little,
'    the procedure turned incredibly complicated. i do not condone the coding
'    style applied here, and that this is not how i usually write code,
'    nevertheless, i'm not open to rewriting this code in a different style,
'    because a clean implementation of this algorithm already exists, as pointed
'    out in the following.
'
'    if you would like to understand the underlying algorithm of how the local
'    path can be found with only the url-path as input, i recommend following
'    the much cleaner implementation by cristian buse:
'    https://github.com/cristianbuse/vba-filetools
'    we developed the algorithm together and wrote separate implementations
'    concurrently. his solution is contained inside a module-level library,
'    split into many procedures and using features like private types and api-
'    functions, that are not available when trying to create a single procedure
'    without dependencies like below. this makes his code more readable.
'
'    both of our solutions are well tested and actively supported with bugfixes
'    and improvements, so both should be equally valid choices for use in your
'    project. the differences in performance/features are marginal and they can
'    often be used interchangeably. if you need more file-system interaction
'    functionality, use cristians library, and if you only need getlocalpath,
'    just copy this function to any module in your project and it will work.
'
' 2) how does this function not work?
'    most other solutions for this problem circulating online (a list of most
'    can be found here: https://stackoverflow.com/a/73577057/12287457) are using
'    one of two approaches, :
'     1. they use the environment variables set by onedrive:
'         - environ(onedrive)
'         - environ(onedrivecommercial)
'         - environ(onedriveconsumer)
'        and replace part of the url with it. there are many problems with this
'        approach:
'         1. they are not being set by onedrive on macos.
'         2. it is unclear exactly which part of the url needs to be replaced.
'         3. environment variables can be changed by the user.
'         4. only there three exist. if more onedrive accounts are logged in,
'            they just overwrite the previous ones.
'        or,
'     2. they use the mount points onedrive writes to the registry here:
'         - \hkey_current_user\software\syncengines\providers\onedrive\
'        this also has several drawbacks:
'         1. the registry is not available on macos.
'         2. it's still unclear exactly what part of the url should be replaced.
'         3. these registry keys can contain mistakes, like for example, when:
'             - synchronizing a folder called "personal" from someone else's
'               personal onedrive
'             - synchronizing a folder called "business1" from someone else's
'               personal onedrive and then relogging your own first business
'               onedrive account
'             - relogging you personal onedrive can change the "cid" property
'               from a folderid formatted cid (e.g. 3dea8a9886f05935!125) to a
'               regular private cid (e.g. 3dea8a9886f05935) for synced folders
'               from other people's onedrives
'
'    for these reasons, this solution uses a completely different approach to
'    solve this problem.
'
' 3) how does this function work?
'    this function builds the web to local translation dictionary by extracting
'    the mount points from the onedrive settings files.
'    it reads files from...
'    on windows:
'        - the "...\appdata\local\microsoft" directory
'    on mac:
'        - the "~/library/containers/com.microsoft.onedrive-mac/data/" & _
'              "library/application support" directory
'        - and/or the "~/library/application support"
'    it reads the following files:
'      - \onedrive\settings\personal\clientpolicy.ini
'      - \onedrive\settings\personal\????????????????.dat
'      - \onedrive\settings\personal\????????????????.ini
'      - \onedrive\settings\personal\global.ini
'      - \onedrive\settings\personal\groupfolders.ini
'      - \onedrive\settings\business#\????????-????-????-????-????????????.dat
'      - \onedrive\settings\business#\????????-????-????-????-????????????.ini
'      - \onedrive\settings\business#\clientpolicy*.ini
'      - \onedrive\settings\business#\global.ini
'      - \office\clp\* (just the filename)
'
'    where:
'     - "*" ... 0 or more characters
'     - "?" ... one character [0-9, a-f]
'     - "#" ... one digit
'     - "\" ... path separator, (= "/" on macos)
'     - the "???..." filenames represent cids)
'
'    on macos, the \office\clp\* exists for each microsoft office application
'    separately. depending on whether the application was already used in
'    active syncing with onedrive it may contain different/incomplete files.
'    in the code, the path of this directory is stored inside the variable
'    "clppath". on macos, the defined clppath might not exist or not contain
'    all necessary files for some host applications, because environ("home")
'    depends on the host app.
'    this is not a big problem as the function will still work, however in
'    this case, specifying a preferredmountpointowner will do nothing.
'    to make sure this directory and the necessary files exist, a file must
'    have been actively synchronized with onedrive by the application whose
'    "home" folder is returned by environ("home") while being logged in
'    to that application with the account whose email is given as
'    preferredmountpointowner, at some point in the past!
'
'    if you are usually working with excel but are using this function in a
'    different app, you can instead use an alternative (excels clp folder) as
'    the clppath as it will most likely contain all the necessary information
'    the alternative clppath is commented out in the code, if you prefer to
'    use excels clp folder per default, just un-comment the respective line
'    in the code.
'———————————————————————————————————————————————————————————————————————————————

'———————————————————————————————————————————————————————————————————————————————
' comments regarding the usage:
' this function can be used as a user defined function (udf) from the worksheet.
' (more on that, see "usage examples")
'
' this function offers three optional parameters to the user, however using
' these should only be necessary in extremely rare situations.
' the best rule regarding their usage: don't use them.
'
' in the following these parameters will still be explained.
'
'1) returnall
'   in some exceptional cases it is possible to map one onedrive webpath to
'   multiple different localpaths. this can happen when multiple business
'   onedrive accounts are logged in on one device, and multiple of these have
'   access to the same onedrive folder and they both decide to synchronize it or
'   add it as link to their mysite library.
'   calling the function with returnall:=true will return all valid localpaths
'   for the given webpath, separated by two forward slashes (//). this should be
'   used with caution, as the return value of the function alone is, should
'   multiple local paths exist for the input webpath, not a valid local path
'   anymore.
'   an example of how to obtain all of the local paths could look like this:
'   dim localpath as string, localpaths() as string
'   localpath = getlocalpath(webpath, true)
'   if not localpath like "http*" then
'       localpaths = split(localpath, "//")
'   end if
'
'2) preferredmountpointowner
'   this parameter deals with the same problem as 'returnall'
'   if the function gets called with returnall:=false (default), and multiple
'   localpaths exist for the given webpath, the function will just return any
'   one of them, as usually, it shouldn't make a difference, because the result
'   directories at both of these localpaths are mirrored versions of the same
'   webpath. nevertheless, this option lets the user choose, which mountpoint
'   should be chosen if multiple localpaths are available. each localpath is
'  'owned' by an onedrive account. if a webpath is synchronized twice, this can
'   only happen by synchronizing it with two different accounts, because
'   onedrive prevents you from synchronizing the same folder twice on a single
'   account. therefore, each of the different localpaths for a given webpath
'   has a unique 'owner'. preferredmountpointowner lets the user select the
'   localpath by specifying the account the localpath should be owned by.
'   this is done by passing the email address of the desired account as
'   preferredmountpointowner.
'   for example, you have two different business onedrive accounts logged in,
'   foo.bar@business1.com and foo.bar@business2.com
'   both synchronize the webpath:
'   webpath = "https://business1.sharepoint.com/sites/testlib/documents/" & _
              "test/test/test/test.xlsm"
'
'   the first one has added it as a link to his personal onedrive, the local
'   path looks like this:
'   c:\users\username\onedrive - business1\testlinkparent\test - testlinklib\...
'   ...test\test.xlsm
'
'   the second one just synchronized it normally, the localpath looks like this:
'   c:\users\username\business1\testlinklib - test\test\test.xlsm
'
'   calling getlocalpath like this:
'   getlocalpath(webpath,,, "foo.bar@business1.com") will return:
'   c:\users\username\onedrive - business1\testlinkparent\test - testlinklib\...
'   ...test\test.xlsm
'
'   calling it like this:
'   getlocalpath(webpath,,, "foo.bar@business2.com") will return:
'   c:\users\username\business1\testlinklib - test\test\test.xlsm
'
'   and calling it like this:
'   getlocalpath(webpath,, true) will return:
'   c:\users\username\onedrive - business1\testlinkparent\test - testlinklib\...
'   ...test\test.xlsm//c:\users\username\business1\testlinklib - test\test\...
'   ...test.xlsm
'
'   calling it normally like this:
'   getlocalpath(webpath) will return any one of the two localpaths, so:
'   c:\users\username\onedrive - business1\testlinkparent\test - testlinklib\...
'   ...test\test.xlsm
'   or
'   c:\users\username\business1\testlinklib - test\test\test.xlsm
'
'3) rebuildcache
'   the function creates a "translation" dictionary from the onedrive settings
'   files and then uses this dictionary to "translate" webpaths to localpaths.
'   this dictionary is implemented as a static variable to the function doesn't
'   have to recreate it every time it is called. it is written on the first
'   function call and reused on all the subsequent calls, making them faster.
'   if the function is called with rebuildcache:=true, this dictionary will be
'   rewritten, even if it was already initialized.
'   note that it is not necessary to use this parameter manually, even if a new
'   mountpoint was added to the onedrive, or a new onedrive account was logged
'   in since the last function call because the function will automatically
'   determine if any of those cases occurred, without sacrificing performance.
'———————————————————————————————————————————————————————————————————————————————
option explicit

''——————————————————————————————————————————————————————————————————————————————
'' usage examples:
'' excel:
'private sub testgetlocalpathexcel()
'    debug.print getlocalpath(thisworkbook.fullname)
'    debug.print getlocalpath(thisworkbook.path)
'end sub
'
' usage as user defined function (udf):
' you might have to replace ; with , in the formulas depending on your settings.
' add this formula to any cell, to get the local path of the workbook:
' =getlocalpath(left(cell("filename";a1);find("[";cell("filename";a1))-1))
'
' to get the local path including the filename (the fullname), use this formula:
' =getlocalpath(left(cell("filename";a1);find("[";cell("filename";a1))-1) &
' textafter(textbefore(cell("filename";a1);"]");"["))
'
''word:
'private sub testgetlocalpathword()
'    debug.print getlocalpath(thisdocument.fullname)
'    debug.print getlocalpath(thisdocument.path)
'end sub
'
''powerpoint:
'private sub testgetlocalpathpowerpoint()
'    debug.print getlocalpath(activepresentation.fullname)
'    debug.print getlocalpath(activepresentation.path)
'end sub
''——————————————————————————————————————————————————————————————————————————————


'this function will convert a onedrive/sharepoint url path, e.g. url containing
'https://d.docs.live.net/; .sharepoint.com/sites; my.sharepoint.com/personal/...
'to the locally synchronized path on your current pc or mac, e.g. a path like
'c:\users\username\onedrive\ on windows; or /users/username/onedrive/ on macos,
'if you have the remote directory locally synchronized with the onedrive app.
'if no local path can be found, the input value will be returned unmodified.
'author: guido witt-dörring
'source: https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
'        https://stackoverflow.com/a/73577057/12287457
public function getlocalpath(byval path as string, _
                    optional byval returnall as boolean = false, _
                    optional byval preferredmountpointowner as string = "", _
                    optional byval rebuildcache as boolean = false) _
                             as string
    #if mac then
        const vberrpermissiondenied            as long = 70
        const vberrinvalidformatinresourcefile as long = 325
        const noerrjustdecodeutf8              as long = 20
        const syncidfilename as string = ".849c9593-d756-4e56-8d6e-42412f2a707b"
        const ismac as boolean = true
        const ps as string = "/" 'application.pathseparator doesn't work
    #else 'windows               'in all host applications (e.g. outlook), hence
        const ps as string = "\" 'conditional compilation is preferred here.
        const ismac as boolean = false
    #end if
    const methodname as string = "getlocalpath"
    const vberrfilenotfound     as long = 53
    const vberroutofmemory      as long = 7
    const vberrkeyalreadyexists as long = 457
    static loctowebcoll as collection, lastcacheupdate as date
    
    if not left(path, 8) = "https://" then getlocalpath = path: exit function

    dim webroot as string, locroot as string, s as string, vitem as variant
    dim pmpo as string: pmpo = lcase$(preferredmountpointowner)
    if not loctowebcoll is nothing and not rebuildcache then
        dim rescoll as collection: set rescoll = new collection
        'if the loctowebcoll is initialized, this logic will find the local path
        for each vitem in loctowebcoll
            locroot = vitem(0): webroot = vitem(1)
            if instr(1, path, webroot, vbtextcompare) = 1 then _
                rescoll.add key:=vitem(2), _
                   item:=replace(replace(path, webroot, locroot, , 1), "/", ps)
        next vitem
        if rescoll.count > 0 then
            if returnall then
                for each vitem in rescoll: s = s & "//" & vitem: next vitem
                getlocalpath = mid$(s, 3): exit function
            end if
            on error resume next: getlocalpath = rescoll(pmpo): on error goto 0
            if getlocalpath <> "" then exit function
            getlocalpath = rescoll(1): exit function
        end if
        'local path was not found with cached mountpoints
        getlocalpath = path 'no exit function here! check if cache needs rebuild
    end if

    dim settpaths as collection: set settpaths = new collection
    dim settpath as variant, clppath as string
    #if mac then 'the settings directories can be in different locations
        dim cloudstoragepath as string, cloudstoragepathexists as boolean
        s = environ("home")
        clppath = s & "/library/application support/microsoft/office/clp/"
        s = left$(s, instrrev(s, "/library/containers/", , vbbinarycompare))
        settpaths.add s & _
                      "library/containers/com.microsoft.onedrive-mac/data/" & _
                      "library/application support/onedrive/settings/"
        settpaths.add s & "library/application support/onedrive/settings/"
        cloudstoragepath = s & "library/cloudstorage/"

        'excels clp folder:
        'clppath = left$(s, instrrev(s, "/library/containers", , 0)) & _
                  "library/containers/com.microsoft.excel/data/" & _
                  "library/application support/microsoft/office/clp/"
    #else 'on windows, the settings directories are always in this location:
        settpaths.add environ("localappdata") & "\microsoft\onedrive\settings\"
        clppath = environ("localappdata") & "\microsoft\office\clp\"
    #end if

    dim i as long
    #if mac then 'request access to all possible directories at once
        dim arrdirs() as variant: redim arrdirs(1 to settpaths.count * 11 + 1)
        for each settpath in settpaths
            for i = i + 1 to i + 9
                arrdirs(i) = settpath & "business" & i mod 11
            next i
            arrdirs(i) = settpath: i = i + 1
            arrdirs(i) = settpath & "personal"
        next settpath
        arrdirs(i + 1) = cloudstoragepath
        if getsetting("getlocalpath", "accessrequestinfomsg", "displayed", _
                      "false") = "false" then doevents 'msgbox "the current " _
            & "vba project requires access to the onedrive settings files to " _
            & "translate an onedrive url to the local path of the locally " & _
            "synchronized file/folder on your mac. because these files are " & _
            "located outside of excels sandbox, file-access must be granted " _
            & "explicitly. please approve the access requests following this " _
            & "message.", vbinformation
        if not grantaccesstomultiplefiles(arrdirs) then _
            err.raise vberrpermissiondenied, methodname
        savesetting "getlocalpath", "accessrequestinfomsg", "displayed", "true"
    #end if
    
    'find all subdirectories in onedrive settings folder:
    dim onedrivesettdirs as collection: set onedrivesettdirs = new collection
    for each settpath in settpaths
        dim dirname as string: dirname = dir(settpath, vbdirectory)
        do until dirname = vbnullstring
            if dirname = "personal" or dirname like "business#" then _
                onedrivesettdirs.add item:=settpath & dirname & ps
            dirname = dir(, vbdirectory)
        loop
    next settpath

    if not loctowebcoll is nothing or ismac then
        dim requiredfiles as collection: set requiredfiles = new collection
        'get collection of all required files
        dim vdir as variant
        for each vdir in onedrivesettdirs
            dim cid as string: cid = iif(vdir like "*" & ps & "personal" & ps, _
                                         "????????????????", _
                                         "????????-????-????-????-????????????")
            dim filename as string: filename = dir(vdir, vbnormal)
            do until filename = vbnullstring
                if filename like cid & ".ini" _
                or filename like cid & ".dat" _
                or filename like "clientpolicy*.ini" _
                or strcomp(filename, "groupfolders.ini", vbtextcompare) = 0 _
                or strcomp(filename, "global.ini", vbtextcompare) = 0 then _
                    requiredfiles.add item:=vdir & filename
                filename = dir
            loop
        next vdir
    end if

    'this part should ensure perfect accuracy despite the mount point cache
    'while sacrificing almost no performance at all by querying filedatetimes.
    if not loctowebcoll is nothing and not rebuildcache then
        'check if a settings file was modified since the last cache rebuild
        dim vfile as variant
        for each vfile in requiredfiles
            if filedatetime(vfile) > lastcacheupdate then _
                rebuildcache = true: exit for 'full cache refresh is required!
        next vfile
        if not rebuildcache then exit function
    end if

    'if execution reaches this point, the cache will be fully rebuilt...
    dim filenum as long, syncid as string, b() as byte
    #if mac then 'variables for manual decoding of utf-8, utf-32 and ansi
        dim j as long, k as long, m as long, ansi() as byte, sansi as string
        dim utf16() as byte, sutf16 as string, utf32() as byte
        dim utf8() as byte, sutf8 as string, numbytesofcodepoint as long
        dim codepoint as long, lowsurrogate as long, highsurrogate as long
    #end if

    lastcacheupdate = now()
    #if mac then 'prepare building syncidtosyncdir dictionary. this involves
        'reading the ".849c9593-d756-4e56-8d6e-42412f2a707b" files inside the
        'subdirs of "~/library/cloudstorage/", list of files and access required
        dim coll as collection: set coll = new collection
        dirname = dir(cloudstoragepath, vbdirectory)
        do until dirname = vbnullstring
            if dirname like "onedrive*" then
                cloudstoragepathexists = true
                vdir = cloudstoragepath & dirname & ps
                vfile = cloudstoragepath & dirname & ps & syncidfilename
                coll.add item:=vdir
                requiredfiles.add item:=vdir 'for pooling file access requests
                requiredfiles.add item:=vfile
            end if
            dirname = dir(, vbdirectory)
        loop

        'pool access request for these files and the onedrive/settings files
        if loctowebcoll is nothing then
            dim vfiles as variant
            if requiredfiles.count > 0 then
                redim vfiles(1 to requiredfiles.count)
               for i = 1 to ubound(vfiles): vfiles(i) = requiredfiles(i): next i
                if not grantaccesstomultiplefiles(vfiles) then _
                    err.raise vberrpermissiondenied, methodname
            end if
        end if

        'more access might be required if some folders inside cloudstoragepath
        'don't contain the hidden file ".849c9593-d756-4e56-8d6e-42412f2a707b".
        'in that case, access to their first level subfolders is also required.
        if cloudstoragepathexists then
            for i = coll.count to 1 step -1
                dim fattr as long: fattr = 0
                on error resume next
                fattr = getattr(coll(i) & syncidfilename)
                dim isfile as boolean: isfile = false
                if err.number = 0 then isfile = not cbool(fattr and vbdirectory)
                on error goto 0
                if not isfile then 'hidden file does not exist
                'dir(path, vbhidden) is unreliable and doesn't work on some macs
                'if dir(coll(i) & syncidfilename, vbhidden) = vbnullstring then
                    dirname = dir(coll(i), vbdirectory)
                    do until dirname = vbnullstring
                        if not dirname like ".trash*" and dirname <> "icon" then
                            coll.add coll(i) & dirname & ps
                            coll.add coll(i) & dirname & ps & syncidfilename, _
                                     coll(i) & dirname & ps  '<- key for removal
                        end if
                        dirname = dir(, vbdirectory)
                    loop          'remove the
                    coll.remove i 'folder if it doesn't contain the hidden file.
                end if
            next i
            if coll.count > 0 then
                redim arrdirs(1 to coll.count)
                for i = 1 to coll.count: arrdirs(i) = coll(i): next i
                if not grantaccesstomultiplefiles(arrdirs) then _
                    err.raise vberrpermissiondenied, methodname
            end if
            'remove all files from coll (not the folders!): reminder:
            on error resume next 'coll(coll(i)) = coll(i) & syncidfilename
            for i = coll.count to 1 step -1
                coll.remove coll(i)
            next i
            on error goto 0

            'write syncidtosyncdir collection
            dim syncidtosyncdir as collection
            set syncidtosyncdir = new collection
            for each vdir in coll
                fattr = 0
                on error resume next
                fattr = getattr(vdir & syncidfilename)
                isfile = false
                if err.number = 0 then isfile = not cbool(fattr and vbdirectory)
                on error goto 0
                if isfile then 'hidden file exists
                'dir(path, vbhidden) is unreliable and doesn't work on some macs
                'if dir(vdir & syncidfilename, vbhidden) <> vbnullstring then
                    filenum = freefile(): s = "": vfile = vdir & syncidfilename
                    'somehow reading these files with "open" doesn't always work
                    dim readsucceeded as boolean: readsucceeded = false
                    on error goto readfailed
                    open vfile for binary access read as #filenum
                        redim b(0 to lof(filenum)): get filenum, , b: s = b
                        readsucceeded = true
readfailed:             on error goto -1
                    close #filenum: filenum = 0
                    on error goto 0
                    if readsucceeded then
                        'debug.print "used open statement to read file: " & _
                                    vdir & syncidfilename
                        ansi = s 'if open was used: decode ansi string manually:
                        if lenb(s) > 0 then
                            redim utf16(0 to lenb(s) * 2 - 1): k = 0
                            for j = lbound(ansi) to ubound(ansi)
                                utf16(k) = ansi(j): k = k + 2
                            next j
                            s = utf16
                        else: s = vbnullstring
                        end if
                    else 'reading the file with "open" failed with an error. try
                        'using applescript. also avoids the manual transcoding.
                        'somehow applscript fails too, sometimes. seems whenever
                        '"open" works, applescript fails and vice versa (?!?!)
                        vfile = macscript("return path to startup disk as " & _
                                    "string") & replace(mid$(vfile, 2), ps, ":")
                        s = macscript("return read file """ & _
                                      vfile & """ as string")
                       'debug.print "used apple script to read file: " & vfile
                    end if
                    if instr(1, s, """guid"" : """, vbbinarycompare) then
                        s = split(s, """guid"" : """)(1)
                        syncid = left$(s, instr(1, s, """", 0) - 1)
                        syncidtosyncdir.add key:=syncid, _
                             item:=vba.array(syncid, left$(vdir, len(vdir) - 1))
                    else
                        debug.print "warning, empty syncidfile encountered!"
                    end if
                end if
            next vdir
        end if
    #end if

    'declare all variables that will be used in the loop over onedrive settings
    dim line as variant, parts() as string, n as long, libnr as string
    dim tag as string, mainmount as string, relpath as string, email as string
    dim parentid as string, folderid as string, foldername as string
    dim folderidpattern as string, foldertype as string, keyexists as boolean
    dim siteid as string, libid as string, webid as string, lnkid as string
    dim mainsyncid as string, syncfind as string, mainsyncfind as string
    'the following are "constants" and needed for reading the .dat files:
    dim sig1 as string:       sig1 = chrb$(2)
    dim sig2 as string * 4:   midb$(sig2, 1) = chrb$(1)
    dim vbnullbyte as string: vbnullbyte = chrb$(0)
    #if mac then
        dim sig3 as string: sig3 = vbnullchar & vbnullchar
    #else 'windows
        dim sig3 as string: sig3 = vbnullchar
    #end if

    'writing loctowebcoll using .ini and .dat files in the onedrive settings:
    'here, a scripting.dictionary would be nice but it is not available on mac!
    dim lastaccountupdates as collection, lastaccountupdate as date
    set lastaccountupdates = new collection
    set loctowebcoll = new collection
    for each vdir in onedrivesettdirs 'one folder per logged in od account
        dirname = mid$(vdir, instrrev(vdir, ps, len(vdir) - 1, 0) + 1)
        dirname = left$(dirname, len(dirname) - 1)

        'read global.ini to get cid
        if dir(vdir & "global.ini", vbnormal) = "" then goto nextfolder
        filenum = freefile()
        open vdir & "global.ini" for binary access read as #filenum
            redim b(0 to lof(filenum)): get filenum, , b
        close #filenum: filenum = 0
        #if mac then 'on mac, the onedrive settings files use utf-8 encoding
            sutf8 = b: on error goto decodeutf8: err.raise noerrjustdecodeutf8
            on error goto 0: b = sutf16 'b = strconv(b, vbunicode) <- unreliable
        #end if
        for each line in split(b, vbnewline)
            if line like "cid = *" then cid = mid$(line, 7): exit for
        next line

        if cid = vbnullstring then goto nextfolder
        if (dir(vdir & cid & ".ini") = vbnullstring or _
            dir(vdir & cid & ".dat") = vbnullstring) then goto nextfolder
        if dirname like "business#" then
            folderidpattern = replace(space$(32), " ", "[a-f0-9]")
        elseif dirname = "personal" then
            folderidpattern = replace(space$(16), " ", "[a-f0-9]") & "!###*"
        end if

        'get email for business accounts
        '(only necessary to let user choose preferredmountpointowner)
        filename = dir(clppath, vbnormal)
        do until filename = vbnullstring
            i = instrrev(filename, cid, , vbtextcompare)
            if i > 1 and cid <> vbnullstring then _
                email = lcase$(left$(filename, i - 2)): exit do
            filename = dir
        loop

        #if mac then
            on error goto -1 '(error from calling code might still be set)
            on error resume next
            lastaccountupdate = lastaccountupdates(dirname)
            keyexists = (err.number = 0)
            on error goto 0
            if keyexists then
                if filedatetime(vdir & cid & ".ini") < lastaccountupdate then
                    goto nextfolder
                else
                    for i = loctowebcoll.count to 1 step -1
                        if loctowebcoll(i)(5) = dirname then
                            loctowebcoll.remove i
                        end if
                    next i
                    lastaccountupdates.remove dirname
                    lastaccountupdates.add key:=dirname, _
                                         item:=filedatetime(vdir & cid & ".ini")
                end if
            else
                lastaccountupdates.add key:=dirname, _
                                      item:=filedatetime(vdir & cid & ".ini")
            end if
        #end if

        'read all the clientploicy*.ini files:
        dim clipolcoll as collection: set clipolcoll = new collection
        filename = dir(vdir, vbnormal)
        do until filename = vbnullstring
            if filename like "clientpolicy*.ini" then
                filenum = freefile()
                open vdir & filename for binary access read as #filenum
                    redim b(0 to lof(filenum)): get filenum, , b
                close #filenum: filenum = 0
                #if mac then 'on mac, onedrive settings files use utf-8 encoding
                    sutf8 = b: on error goto decodeutf8
                    err.raise noerrjustdecodeutf8 'this is not an error!
                    on error goto 0: b = sutf16 'strconv(b, vbunicode)unreliable
                #end if
                clipolcoll.add key:=filename, item:=new collection
                for each line in split(b, vbnewline)
                    if instr(1, line, " = ", vbbinarycompare) then
                        tag = left$(line, instr(1, line, " = ", 0) - 1)
                        s = mid$(line, instr(1, line, " = ", 0) + 3)
                        select case tag
                        case "davurlnamespace"
                            clipolcoll(filename).add key:=tag, item:=s
                        case "siteid", "irmlibraryid", "webid" 'only used for
                            s = replace(lcase$(s), "-", "") 'backup method later
                            if len(s) > 3 then s = mid$(s, 2, len(s) - 2)
                            clipolcoll(filename).add key:=tag, item:=s
                        end select
                    end if
                next line
            end if
            filename = dir
        loop

        'read cid.dat file
        const chunkoverlap          as long = 1000
        const maxdirname            as long = 255
        dim buffsize as long: buffsize = -1 'buffer uninitialized
try:    on error goto catch
        dim odfolders as collection: set odfolders = new collection
        dim lastchunkendpos as long: lastchunkendpos = 1
        dim lastfileupdate as date:  lastfileupdate = filedatetime(vdir & _
                                                                   cid & ".dat")
        i = 0 'i = current reading pos.
        do
            'ensure file is not changed while reading it
            if filedatetime(vdir & cid & ".dat") > lastfileupdate then goto try
            filenum = freefile
            open vdir & cid & ".dat" for binary access read as #filenum
                dim lendatfile as long: lendatfile = lof(filenum)
                if buffsize = -1 then buffsize = lendatfile 'initialize buffer
                'overallocate a bit so read chunks overlap to recognize all dirs
                redim b(0 to buffsize + chunkoverlap)
                get filenum, lastchunkendpos, b: s = b
                dim size as long: size = lenb(s)
            close #filenum: filenum = 0
            lastchunkendpos = lastchunkendpos + buffsize

            for vitem = 16 to 8 step -8
                i = instrb(vitem + 1, s, sig2, 0) 'sarch pattern in cid.dat
                do while i > vitem and i < size - 168 'and confirm with another
                    if strcomp(midb$(s, i - vitem, 1), sig1, 0) = 0 then 'one
                        i = i + 8: n = instrb(i, s, vbnullbyte, 0) - i
                        if n < 0 then n = 0           'i:start pos, n: length
                        if n > 39 then n = 39
                        #if mac then 'strconv doesn't work reliably on mac ->
                            ansi = midb$(s, i, n) 'decode ansi string manually:
                            j = ubound(ansi) - lbound(ansi) + 1
                            if j > 0 then
                                redim utf16(0 to j * 2 - 1): k = 0
                                for j = lbound(ansi) to ubound(ansi)
                                    utf16(k) = ansi(j): k = k + 2
                                next j
                                folderid = utf16
                            else: folderid = vbnullstring
                            end if
                        #else 'windows
                            folderid = strconv(midb$(s, i, n), vbunicode)
                        #end if
                        i = i + 39: n = instrb(i, s, vbnullbyte, 0) - i
                        if n < 0 then n = 0
                        if n > 39 then n = 39
                        #if mac then 'strconv doesn't work reliably on mac ->
                            ansi = midb$(s, i, n) 'decode ansi string manually:
                            j = ubound(ansi) - lbound(ansi) + 1
                            if j > 0 then
                                redim utf16(0 to j * 2 - 1): k = 0
                                for j = lbound(ansi) to ubound(ansi)
                                    utf16(k) = ansi(j): k = k + 2
                                next j
                                parentid = utf16
                            else: parentid = vbnullstring
                            end if
                        #else 'windows
                            parentid = strconv(midb$(s, i, n), vbunicode)
                        #end if
                        i = i + 121
                        n = instr(-int(-(i - 1) / 2) + 1, s, sig3) * 2 - i - 1
                        if n > maxdirname * 2 then n = maxdirname * 2
                        if n < 0 then n = 0
                        if folderid like folderidpattern _
                        and parentid like folderidpattern then
                            #if mac then 'encoding of folder names is utf-32-le
                                do while n mod 4 > 0
                                    if n > maxdirname * 4 then exit do
                                    n = instr(-int(-(i + n) / 2) + 1, s, sig3) _
                                        * 2 - i - 1
                                loop
                                if n > maxdirname * 4 then n = maxdirname * 4
                                utf32 = midb$(s, i, n)
                                'utf-32 can only be converted manually to utf-16
                                redim utf16(lbound(utf32) to ubound(utf32))
                                j = lbound(utf32): k = lbound(utf32)
                                do while j < ubound(utf32)
                                    if utf32(j + 2) + utf32(j + 3) = 0 then
                                        utf16(k) = utf32(j)
                                        utf16(k + 1) = utf32(j + 1)
                                        k = k + 2
                                    else
                                        if utf32(j + 3) <> 0 then err.raise _
                                            vberrinvalidformatinresourcefile, _
                                            methodname
                                        codepoint = utf32(j + 2) * &h10000 + _
                                                    utf32(j + 1) * &h100& + _
                                                    utf32(j)
                                        m = codepoint - &h10000
                                        highsurrogate = &hd800& or (m \ &h400&)
                                        lowsurrogate = &hdc00& or (m and &h3ff)
                                        utf16(k) = highsurrogate and &hff&
                                        utf16(k + 1) = highsurrogate \ &h100&
                                        utf16(k + 2) = lowsurrogate and &hff&
                                        utf16(k + 3) = lowsurrogate \ &h100&
                                        k = k + 4
                                    end if
                                    j = j + 4
                                loop
                                if k > lbound(utf16) then
                                    redim preserve utf16(lbound(utf16) to k - 1)
                                    foldername = utf16
                                else: foldername = vbnullstring
                                end if
                            #else 'on windows encoding is utf-16-le
                                foldername = midb$(s, i, n)
                            #end if
                            'vba.array() instead of just array() is used in this
                            'function because it ignores option base 1
                            odfolders.add vba.array(parentid, foldername), _
                                          folderid
                        end if
                    end if
                    i = instrb(i + 1, s, sig2, 0) 'find next sig2 in cid.dat
                loop
                if odfolders.count > 0 then exit for
            next vitem
        loop until lastchunkendpos >= lendatfile _
                or buffsize >= lendatfile
        goto continue
catch:
        select case err.number
        case vberrkeyalreadyexists
            'this can happen at chunk boundries, folder might get added twice:
            odfolders.remove folderid 'make sure the folder gets added new again
            resume 'to avoid foldernames truncated by chunk ends
        case is <> vberroutofmemory: err.raise err, methodname
        end select
        if buffsize > &hfffff then buffsize = buffsize / 2: resume try
        err.raise err, methodname 'raise error if less than 1 mb ram available
continue:
        on error goto 0

        'read cid.ini file
        filenum = freefile()
        open vdir & cid & ".ini" for binary access read as #filenum
            redim b(0 to lof(filenum)): get filenum, , b
        close #filenum: filenum = 0
        #if mac then 'on mac, the onedrive settings files use utf-8 encoding
            sutf8 = b: on error goto decodeutf8: err.raise noerrjustdecodeutf8
            on error goto 0: b = sutf16 'b = strconv(b, vbunicode) <- unreliable
        #end if
        select case true
        case dirname like "business#" 'settings files for a business od account
        'max 9 business onedrive accounts can be signed in at a time.
           dim libnrtowebcoll as collection: set libnrtowebcoll = new collection
            mainmount = vbnullstring
            for each line in split(b, vbnewline)
                webroot = "": locroot = "": parts = split(line, """")
                select case left$(line, instr(1, line, " = ", 0) - 1)
                case "libraryscope" 'one line per synchronized library
                    locroot = parts(9)
                    syncfind = locroot: syncid = split(parts(10), " ")(2)
                    if locroot = vbnullstring then libnr = split(line, " ")(2)
                    foldertype = parts(3): parts = split(parts(8), " ")
                    siteid = parts(1): webid = parts(2): libid = parts(3)
                    if mainmount = vbnullstring and foldertype = "odb" then
                        mainmount = locroot: filename = "clientpolicy.ini"
                        mainsyncid = syncid: mainsyncfind = syncfind
                    else: filename = "clientpolicy_" & libid & siteid & ".ini"
                    end if
                    on error resume next 'on error try backup method...
                    webroot = clipolcoll(filename)("davurlnamespace")
                    on error goto 0
                    if webroot = "" then 'backup method to find webroot:
                        for each vitem in clipolcoll
                            if vitem("siteid") = siteid _
                            and vitem("webid") = webid _
                            and vitem("irmlibraryid") = libid then
                                webroot = vitem("davurlnamespace"): exit for
                            end if
                        next vitem
                    end if
                    if webroot = vbnullstring then err.raise vberrfilenotfound _
                                                           , methodname
                    if locroot = vbnullstring then
                        libnrtowebcoll.add vba.array(libnr, webroot), libnr
                    else
                        loctowebcoll.add vba.array(locroot, webroot, email, _
                                        syncid, syncfind, dirname), key:=locroot
                    end if
                case "libraryfolder" 'one line per synchronized library folder
                    libnr = split(line, " ")(3)
                    locroot = parts(1): syncfind = locroot
                    syncid = split(parts(4), " ")(1)
                    s = vbnullstring: parentid = left$(split(line, " ")(4), 32)
                    do  'if not synced at the bottom dir of the library:
                        '   -> add folders below mount point to webroot
                        on error resume next: odfolders parentid
                        keyexists = (err.number = 0): on error goto 0
                        if not keyexists then exit do
                        s = odfolders(parentid)(1) & "/" & s
                        parentid = odfolders(parentid)(0)
                    loop
                    webroot = libnrtowebcoll(libnr)(1) & s
                    loctowebcoll.add vba.array(locroot, webroot, email, _
                                             syncid, syncfind, dirname), locroot
                case "addedscope" 'one line per folder added as link to personal
                    relpath = parts(5): if relpath = " " then relpath = ""  'lib
                    parts = split(parts(4), " "): siteid = parts(1)
                    webid = parts(2): libid = parts(3): lnkid = parts(4)
                    filename = "clientpolicy_" & libid & siteid & lnkid & ".ini"
                    on error resume next 'on error try backup method...
                    webroot = clipolcoll(filename)("davurlnamespace") & relpath
                    on error goto 0
                    if webroot = "" then 'backup method to find webroot:
                        for each vitem in clipolcoll
                            if vitem("siteid") = siteid _
                            and vitem("webid") = webid _
                            and vitem("irmlibraryid") = libid then
                                webroot = vitem("davurlnamespace") & relpath
                                exit for
                            end if
                        next vitem
                    end if
                    if webroot = vbnullstring then err.raise vberrfilenotfound _
                                                           , methodname
                    s = vbnullstring: parentid = left$(split(line, " ")(3), 32)
                    do 'if link is not at the bottom of the personal library:
                        on error resume next: odfolders parentid
                        keyexists = (err.number = 0): on error goto 0
                        if not keyexists then exit do       'add folders below
                        s = odfolders(parentid)(1) & ps & s 'mount point to
                        parentid = odfolders(parentid)(0)   'locroot
                    loop
                    locroot = mainmount & ps & s
                    loctowebcoll.add vba.array(locroot, webroot, email, _
                                     mainsyncid, mainsyncfind, dirname), locroot
                case else: exit for
                end select
            next line
        case dirname = "personal" 'settings files for a personal od account
        'only one personal onedrive account can be signed in at a time.
            for each line in split(b, vbnewline) 'loop should exit at first line
                if line like "library = *" then
                    parts = split(line, """"): locroot = parts(3)
                    syncfind = locroot: syncid = split(parts(4), " ")(2)
                    exit for
                end if
            next line
            on error resume next 'this file may be missing if the personal od
            webroot = clipolcoll("clientpolicy.ini")("davurlnamespace") 'account
            on error goto 0                  'was logged out of the onedrive app
            if locroot = "" or webroot = "" or cid = "" then goto nextfolder
            loctowebcoll.add vba.array(locroot, webroot & "/" & cid, email, _
                                       syncid, syncfind, dirname), key:=locroot
            if dir(vdir & "groupfolders.ini") = "" then goto nextfolder
            'read groupfolders.ini file
            cid = vbnullstring: filenum = freefile()
            open vdir & "groupfolders.ini" for binary access read as #filenum
                redim b(0 to lof(filenum)): get filenum, , b
            close #filenum: filenum = 0
            #if mac then 'on mac, the onedrive settings files use utf-8 encoding
                sutf8 = b: on error goto decodeutf8
                err.raise noerrjustdecodeutf8
                on error goto 0: b = sutf16 'strconv(b, vbunicode) is unreliable
            #end if 'two lines per synced folder from other peoples personal ods
            for each line in split(b, vbnewline)
                if line like "*_baseuri = *" and cid = vbnullstring then
                    cid = lcase$(mid$(line, instrrev(line, "/", , 0) + 1, 16))
                    folderid = left$(line, instr(line, "_") - 1)
                elseif cid <> vbnullstring then
                    loctowebcoll.add vba.array(locroot & ps & odfolders( _
                                     folderid)(1), webroot & "/" & cid & "/" & _
                                     mid$(line, len(folderid) + 9), email, _
                                     syncid, syncfind, dirname), _
                                key:=locroot & ps & odfolders(folderid)(1)
                    cid = vbnullstring: folderid = vbnullstring
                end if
            next line
        end select
nextfolder:
        cid = vbnullstring: s = vbnullstring: email = vbnullstring
    next vdir

    'clean the finished "dictionary" up, remove trailing "\" and "/"
    dim tmpcoll as collection: set tmpcoll = new collection
    for each vitem in loctowebcoll
        locroot = vitem(0): webroot = vitem(1): syncfind = vitem(4)
        if right$(webroot, 1) = "/" then _
            webroot = left$(webroot, len(webroot) - 1)
        if right$(locroot, 1) = ps then _
            locroot = left$(locroot, len(locroot) - 1)
        if right$(syncfind, 1) = ps then _
            syncfind = left$(syncfind, len(syncfind) - 1)
        tmpcoll.add vba.array(locroot, webroot, vitem(2), _
                              vitem(3), syncfind), locroot
    next vitem
    set loctowebcoll = tmpcoll

    #if mac then 'deal with syncids
        if cloudstoragepathexists then
            set tmpcoll = new collection
            for each vitem in loctowebcoll
                locroot = vitem(0): syncid = vitem(3): syncfind = vitem(4)
                locroot = replace(locroot, syncfind, _
                                           syncidtosyncdir(syncid)(1), , 1)
                tmpcoll.add vba.array(locroot, vitem(1), vitem(2)), locroot
            next vitem
            set loctowebcoll = tmpcoll
        end if
    #end if

    getlocalpath = getlocalpath(path, returnall, pmpo, false): exit function
    exit function
decodeutf8: 'by abusing error handling, code duplication is avoided
    #if mac then     'strconv doesn't work reliably, therefore utf-8 must
        utf8 = sutf8 'be transcoded to utf-16 manually (yes, this is insane)
        redim utf16(0 to (ubound(utf8) - lbound(utf8) + 1) * 2)
        i = lbound(utf8): k = 0
        do while i <= ubound(utf8) 'loop through the utf-8 byte array
            'determine the number of bytes in the current utf-8 codepoint
            numbytesofcodepoint = 1
            if utf8(i) and &h80 then
                if utf8(i) and &h20 then
                    if utf8(i) and &h10 then
                        numbytesofcodepoint = 4
                    else: numbytesofcodepoint = 3: end if
                else: numbytesofcodepoint = 2: end if
            end if
            if i + numbytesofcodepoint - 1 > ubound(utf8) then _
                err.raise vberrinvalidformatinresourcefile, methodname
            'calculate the unicode codepoint value from the utf-8 bytes
            if numbytesofcodepoint = 1 then
                codepoint = utf8(i)
            else: codepoint = utf8(i) and (2 ^ (7 - numbytesofcodepoint) - 1)
                for j = 1 to numbytesofcodepoint - 1
                    codepoint = (codepoint * 64) + (utf8(i + j) and &h3f)
                next j
            end if
            'convert the unicode codepoint to utf-16le bytes
            if codepoint < &h10000 then
                utf16(k) = codepoint and &hff&
                utf16(k + 1) = codepoint \ &h100&
                k = k + 2
            else 'codepoint must be encoded as surrogate pair
                m = codepoint - &h10000
                highsurrogate = &hd800& or (m \ &h400&)
                lowsurrogate = &hdc00& or (m and &h3ff)
                utf16(k) = highsurrogate and &hff&
                utf16(k + 1) = highsurrogate \ &h100&
                utf16(k + 2) = lowsurrogate and &hff&
                utf16(k + 3) = lowsurrogate \ &h100&
                k = k + 4
            end if
            i = i + numbytesofcodepoint 'move to the next utf-8 codepoint
        loop
        if k > 0 then
            redim preserve utf16(0 to k - 1)
            sutf16 = utf16
        else: sutf16 = ""
        end if
        resume next 'clear the error object, and jump back to the statement
    #end if         'after which the "pseudo-error" was raised.
end function
