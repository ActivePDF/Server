Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Stamp Images and Text onto the output PDF
oSVR.AddStampCollection "TXTinternal"
oSVR.StampFont = "Helvetica"
oSVR.StampFontSize = 108
oSVR.StampFontTransparency = 0.3
oSVR.StampRotation = 45.0

oSVR.StampFillMode = 2
oSVR.SetStampColor 255, 0, 0, 0
oSVR.SetStampStrokeColor 100, 0, 0, 0

oSVR.AddStampText 116.0, 156.0, "Internal Only"

' Set whether the stamp collection(s) appears in the background or
' foreground
oSVR.StampBackground = 0

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.AddStampText.pdf")

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