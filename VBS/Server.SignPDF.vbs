Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' ActivePDF Service runs under the 'Local System' account, as such it will only see
' certificates located in HKEY_LOCAL_MACHINE. To use certificates located under
' a specific user set the impersonate user options in the GUI or API
oSVR.InvisiblySignFile "localhost", "My", True, "Mission Viejo, CA", _
    "Approval", "949-555-1212", 1

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.SignPDF.pdf")

' Output conversion result
WriteResult result

' Process Complete
Wscript.Quit

Sub WriteResult(oResult)
  message = "Result Status: " & result.ServerStatus
  If result.ServerStatus = 0 Then
      message = message & ", Success!"
  Else
      message = message &", " & result.Details
  End If
  Wscript.Echo message
End Sub