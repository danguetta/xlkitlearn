-- This AppleScript file allows VBA to carry out asyn curls on a Mac

on doRequest(input_params)
        -- Split the input parameters appropriately
        set {request_type, request_args, target_url} to SplitString(input_params, "|")

        -- Define the curl command
        set curl_command to "curl -X " & (quoted form of request_type) & " -H 'Content-Type:application/json' -d " & (quoted form of UnescapeString(request_args)) & " " & (quoted form of target_url) & " &>/dev/n$

        --display dialog curl_command

        -- Run the curl command without capturing the output
        do shell script curl_command
end doRequest

on SplitString(TheBigString, fieldSeparator)
        # From Ron de Bruin's "Mail from Excel 2016 with Mac Mail example": www.rondebruin.nl
        tell AppleScript
                set oldTID to text item delimiters
                set text item delimiters to fieldSeparator
                set theItems to text items of TheBigString
                set text item delimiters to oldTID
        end tell
        return theItems
end SplitString

on UnescapeString(inputText)
        # Replaces all instances of \" with "

        -- Set the text item delimiters to the escaped quote (\" )
        set AppleScript's text item delimiters to "\\\""

        -- Break the inputText into a list of text items
        set textItems to text items of inputText

        -- Set the text item delimiters to a regular quote (")
        set AppleScript's text item delimiters to "\""

        -- Join the text items back together with the regular quote
        set resultText to textItems as text

        -- Reset the text item delimiters to default
        set AppleScript's text item delimiters to ""

        -- Return the modified text
        return resultText
end UnescapeString