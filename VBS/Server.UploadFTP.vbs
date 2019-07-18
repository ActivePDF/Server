Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Setup the FTP request supplying credentials if needed
oSVR.AddFTPRequest "#.#.#.#", "/folder"
oSVR.SetFTPCredentials "john.doe", "asdfasdf"

' Set which files will upload with the FTP request
' To attach a binary file use AddFTPBinaryAttachment
oSVR.FTPAttachOutput = True
oSVR.AddFTPAttachment strPath & "Server.Input.ps"

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.UploadFTP.pdf")

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