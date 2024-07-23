private sub userform_initialize()
    lst_errors.columncount = 5
    lst_errors.columnwidths = "50,15,170,15,500"
    
    ' populate the header errors
    dim i as integer
    for i = 1 to header_errors.count()
        lst_errors.additem
        
        lst_errors.list(i - 1, 0) = split(header_errors.item(i), "`")(0)
        lst_errors.list(i - 1, 1) = ""
        lst_errors.list(i - 1, 2) = split(header_errors.item(i), "`")(1)
        lst_errors.list(i - 1, 3) = ""
        lst_errors.list(i - 1, 4) = split(header_errors.item(i), "`")(2)
    next i
end sub