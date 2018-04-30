Dim wmiDateTime, wmi, query, eventItems, eventItem
Dim timeWritten, eventDate, eventTime, description
Dim eventsDict, eventInfo, errorCount, i

'----------------------------------------------------------------------------------------------------------------------------
'Name       : SQ          -> Places single quotes around a string
'Parameters : stringValue -> String containing the value to place single quotes around
'Return     : SQ          -> Returns a single quoted string
'----------------------------------------------------------------------------------------------------------------------------
Function SQ(ByVal stringValue)
   If VarType(stringValue) = vbString Then
      SQ = "'" & stringValue & "'"
   End If
End Function

'----------------------------------------------------------------------------------------------------------------------------
'Name       : DQ          -> Place double quotes around a string and replace double quotes
'           :             -> within the string with pairs of double quotes.
'Parameters : stringValue -> String value to be double quoted
'Return     : DQ          -> Double quoted string.
'----------------------------------------------------------------------------------------------------------------------------
Function DQ (ByVal stringValue)
   If stringValue <> "" Then
      DQ = """" & Replace (stringValue, """", """""") & """"
   Else
      DQ = """"""
   End If
End Function

Function ConvertWMIDateTime(wmiDateTimeString)
   Dim integerValues, i
   '-------------------------------------------------------------------------------------------------------------------------
   'Ensure the wmiDateTimeString contains a "+" or "-" character. If it doesn't it is not a valid WMI date time so exit.
   '-------------------------------------------------------------------------------------------------------------------------
   If InStr(1, wmiDateTimeString, "+", vbTextCompare) = 0 And _
      InStr(1, wmiDateTimeString, "-", vbTextCompare) = 0 Then
      ConvertWMIDateTime = ""
      Exit Function
   End If
   '-------------------------------------------------------------------------------------------------------------------------
   'Replace any "." or "+" or "-" characters in the wmiDateTimeString and check each character is a valid integer.
   '-------------------------------------------------------------------------------------------------------------------------   
   integerValues = Replace(Replace(Replace(wmiDateTimeString, ".", ""), "+", ""), "-", "")
   For i = 1 To Len(integerValues)
      If Not IsNumeric(Mid(integerValues, i, 1)) Then
         ConvertWMIDateTime = ""
         Exit Function
      End If
   Next
   '-------------------------------------------------------------------------------------------------------------------------
   'Convert the WMI Date Time string to a String that can be formatted as a valid Date Time value.
   '-------------------------------------------------------------------------------------------------------------------------
   ConvertWMIDateTime = Left(wmiDateTimeString, 4) & "-" & _ 
							  Mid(wmiDateTimeString, 5, 2)  & "-" & _
                              Mid(wmiDateTimeString, 7, 2)  & " " & _
                              Mid(wmiDateTimeString, 9, 2)  & ":" & _
                              Mid(wmiDateTimeString, 11, 2) & ":" & _
                              Mid(wmiDateTimeString, 13, 2)
End Function

query = "Select * from Win32_NTLogEvent Where Logfile = 'Security' And EventCode = 4624"
Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
If Err.Number <> 0 Then
 WScript.echo "Creating WMI Object to connect to " & DQ(hostName)
 
End If
'----------------------------------------------------------------------------------------------------------------------
'Create the "SWbemDateTime" Object for converting WMI Date formats. Supported in Windows Server 2003 & Windows XP.
'----------------------------------------------------------------------------------------------------------------------
Set wmiDateTime = CreateObject("WbemScripting.SWbemDateTime")
If Err.Number <> 0 Then
 Wscript.Echo "Creating " & DQ("WbemScripting.SWbemDateTime") & " object"
 
End If
'----------------------------------------------------------------------------------------------------------------------
'Build the WQL query and execute it.
'----------------------------------------------------------------------------------------------------------------------
startDateTime = DateAdd("d", -7, Now)
wmiDateTime.SetVarDate startDateTime, True
query          = query & " And (TimeWritten >= " & SQ(wmiDateTime.Value) & ")"

Set eventItems = wmi.ExecQuery(query)
If Err.Number <> 0 Then
 WScript.echo "Executing WMI Query " & DQ(query)
 
End If


For Each eventItem In eventItems
	If eventItem.InsertionStrings(8) = 2 Then
		If eventItem.InsertionStrings(6) = "PMH" Then
			timeWritten = ConvertWMIDateTime(eventItem.TimeWritten)
			wscript.echo "<WINUSERS>"
			wscript.echo "<NAME>" & LCase(eventItem.InsertionStrings(5)) & "</NAME>"
			wscript.echo "<LOGINTIME>" & timeWritten & "</LOGINTIME>"
			wscript.echo "</WINUSERS>"
		End If
	End If
Next
