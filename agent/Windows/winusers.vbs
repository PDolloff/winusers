Dim oWMIDateTime, oWMI, sQuery, oEventItems, oEventItem
Dim sTimeWritten

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
   set dtConverter = CreateObject("WbemScripting.SWbemDateTime")
   '------------------------------------------------------------------------------------------------------------
   'Ensure the wmiDateTimeString contains a "+" or "-" character. If it doesn't it is not a valid oWMI date time so exit.
   '---------------------------------------------------------------------------------------------------------------
   If InStr(1, wmiDateTimeString, "+", vbTextCompare) = 0 And _
      InStr(1, wmiDateTimeString, "-", vbTextCompare) = 0 Then
      ConvertWMIDateTime = ""
      Exit Function
   End If

   ' convert to local time
   dtConverter.Value = wmiDateTimeString
   temptime = CDate(dtConverter.GetVarDate(true))
   sYear = DatePart("yyyy",temptime)
   sMonth = Right("00" & DatePart("m",temptime),2)
   sDay = Right("00" & DatePart("d",temptime),2)
   sHour = Right("00" & DatePart("h",temptime),2)
   sMinute = Right("00" & DatePart("n",temptime),2)
   sSecond = Right("00" & DatePart("s",temptime),2)

   ConvertWMIDateTime = sYear & "-" & sMonth & "-" & sDay & " " & sHour & ":" & sMinute & ":" & sSecond

End Function


sQuery = "Select * from Win32_NTLogEvent Where Logfile = 'Security' And EventCode = 4624"
Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
If Err.Number <> 0 Then
   WScript.echo "Creating WMI Object to connect to " & DQ(hostName)
End If
'----------------------------------------------------------------------------------------------------------------------
'Create the "SWbemDateTime" Object for converting WMI Date formats. Supported in Windows Server 2003 & Windows XP.
'----------------------------------------------------------------------------------------------------------------------
Set oWMIDateTime = CreateObject("WbemScripting.SWbemDateTime")
If Err.Number <> 0 Then
   Wscript.Echo "Creating " & DQ("WbemScripting.SWbemDateTime") & " object"
End If
'----------------------------------------------------------------------------------------------------------------------
'Build the WQL query and execute it.
'----------------------------------------------------------------------------------------------------------------------
startDateTime = DateAdd("d", -7, Now)
oWMIDateTime.SetVarDate startDateTime, True
sQuery = sQuery & " And (TimeWritten >= " & SQ(oWMIDateTime.Value) & ")"

Set oEventItems = oWMI.ExecQuery(sQuery)
If Err.Number <> 0 Then
   WScript.echo "Executing WMI Query " & DQ(query)
End If

' get userdomain environment variable
Set oShell = CreateObject("Wscript.Shell")
sUserDomain = LCase(oShell.ExpandEnvironmentStrings("%USERDOMAIN%"))

' Use a dictionary object to remove duplicates
Set oLogEntries = CreateObject("scripting.dictionary")

For Each oEventItem In oEventItems
   If oEventItem.InsertionStrings(8) = 2 Then
      sTimeWritten = ConvertWMIDateTime(oEventItem.TimeWritten)
      sUsername = LCase(oEventItem.InsertionStrings(5))
      sUserDomainTemp = LCase(oEventItem.InsertionStrings(6))

      mykey = myUsername & " | " & sTimeWritten
      if (oLogEntries.Item(mykey) <> 1) And (sUserDomain = sUserDomainTemp)  Then
         wscript.echo "<WINUSERS>"
		 wscript.echo "<NAME>" & sUsername & "</NAME>"
		 wscript.echo "<LOGINTIME>" & sTimeWritten & "</LOGINTIME>"
		 wscript.echo "</WINUSERS>"
         oLogEntries.Item(mykey) = 1
      End If

   End If
Next
