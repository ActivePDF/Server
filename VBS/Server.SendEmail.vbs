Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Add an email
oSVR.AddEMail

' Set server information
' UPDATE THIS LINE WITH YOUR SERRVER INFORMATION
oSVR.SetSMTPInfo "#.#.#.#", 0
oSVR.SetSMTPCredentials "john.doe", "activePDF", "asdfasdf"

' Set email addresses
oSVR.SetSenderInfo "John Doe", "john.doe@asdidlwenra.com"
oSVR.SetReplyToInfo "John Doe", "john.doe@asdidlwenra.com"
oSVR.SetRecipientInfo "Jane Doe", "jane.doe@asdidlwenra.com"
oSVR.AddToCC "Jim Doe", "jim.doe@asdidlwenra.com"
oSVR.AddToBcc "Janice Doe", "janice.doe@asdidlwenra.com"

' Subject and Body
oSVR.EMailSubject = "PDF Delivery from activePDF"
oSVR.SetEMailBody "<html><body style='background-color: #EEE; padding: 4px;'>Here is your PDF!</body></html>", True

' Attachments - Binary attachments can be added with
' AddEMailBinaryAttachment
oSVR.AddEMailAttachment strPath & "Server.Input.ps"

' Other email options
oSVR.EMailReadReceipt = False
oSVR.EMailAttachOutput = True

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.SendEmail.pdf")

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