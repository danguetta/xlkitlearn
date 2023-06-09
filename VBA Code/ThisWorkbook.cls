Option Explicit

Public Sub Workbook_Open()
    If InStr(1, ThisWorkbook.Name, "DEBUG") > 0 Then
        'We're deving. Show the code sheet
        DEBUG_VERSION = True
        Sheet2.Visible = -1
        Sheet4.Visible = -1
    Else
        DEBUG_VERSION = False
        Sheet2.Visible = 2
        Sheet4.Visible = 2
    End If

    ' Update the names in this workbook
    update_names
    
    ' Update the configuration
    update_conf True
    
    ' Re-set the add-in
    format_sheet
    
    ' Check whether Python is correctly set up
    #If Mac Then
        #If MAC_OFFICE_VERSION < 15 Then
            MsgBox "You are using Excel 2011 for Mac or earlier. Unfortunately, xlkitlearn was only " & _
                    "tested in Excel 2016 or later. For earlier version support, please contact " & _
                    "the authors.", vbCritical + vbRetryCancel + vbDefaultButton2
            
            GoTo close_workbook
        #End If
        
        ' Ensure the interpreter exists
        If Not FileOrFolderExistsOnMac(ParentFolder(ParentFolder(GetConfig("INTERPRETER_MAC", "")))) Then
            Dim msg_output
            msg_output = MsgBox("It looks like you are trying to run the XLKitLearn Excel file without first installing the python interpreter " & _
                    "required. Please go to guetta.org/xlkitlearn and follow the installation instructions there to the letter, " & _
                    "then try again." & vbCrLf & vbCrLf & "If you know what you're doing, please press Retry below to display " & _
                    "internal settings to try and fix this. If not, press Cancel to close xlkitlearn.", vbCritical + vbRetryCancel + vbDefaultButton2)
            
            If msg_output = vbCancel Then
                GoTo close_workbook
            Else
                ThisWorkbook.Sheets("xlwings.conf").Visible = xlSheetVisible
                ThisWorkbook.Sheets("xlwings.conf").Activate
                Exit Sub
            End If
        End If
        
        'Deal with the advanced commands
        Sheets("Add-in").Buttons("btn_kill").Visible = True
        
        Sheets("Add-in").CheckBoxes("chk_server").Visible = False
        Sheets("Add-in").CheckBoxes("chk_foreground").Visible = False
        
    #Else
        ' Ensure the interpreter exists
        If Not FileExists(GetConfig("INTERPRETER_WIN", "")) Then
            Dim msg_output
            msg_output = MsgBox("It looks like you are trying to run the XLKitLearn Excel file without first installing the XLKitLearn.exe " & _
                    "file. Please go to guetta.org/xlkitlearn, download the file XLKitLearn.exe, and install it before trying " & _
                    "again." & vbCrLf & vbCrLf & "If you know what you're doing, please press Retry below to display " & _
                    "internal settings to try and fix this. If not, press Cancel to close xlkitlearn.", vbCritical + vbRetryCancel + vbDefaultButton2)
            
            If msg_output = vbCancel Then
                GoTo close_workbook
            Else
                ThisWorkbook.Sheets("xlwings.conf").Visible = xlSheetVisible
                ThisWorkbook.Sheets("xlwings.conf").Activate
                Exit Sub
            End If
        End If
                
        ' Deal with the advanced commands
        Sheets("Add-in").Buttons("btn_kill").Visible = False
        
        Sheets("Add-in").CheckBoxes("chk_server").Visible = True
        Sheets("Add-in").CheckBoxes("chk_foreground").Visible = True
    #End If
    
    ' Check the version of the interpreter matches
    On Error GoTo after_version_check
    
    Dim version_n As String
    #If Mac Then
        version_n = ReadFile(ParentFolder(ParentFolder(GetConfig("INTERPRETER_MAC", ""))) & "/data/version")
    #Else
        version_n = ReadFile(ParentFolder(GetConfig("INTERPRETER_WIN", "")) & "\data\version")
    #End If
    
    version_n = Replace(Replace(version_n, Chr(10), ""), Chr(13), "")
    
    If version_n <> Split(addin_version(), ".")(0) Then
        If version_n = "" Then
            version_n = "unknown"
        End If
        MsgBox "It looks like the version of XLKitLearn you have installed on your computer is not compatible with " & _
               "this Excel file. The version on your computer is " & version_n & ", and the version of this Excel is " & _
               Split(addin_version(), ".")(0) & ". Please go back to guetta.org/xlkitlearn and follow the installation " & _
               "instructions again.", vbCritical
    End If
after_version_check:
    
    Exit Sub
close_workbook:
    If DEBUG_VERSION = False Then
        ThisWorkbook.Close False
    End If
End Sub