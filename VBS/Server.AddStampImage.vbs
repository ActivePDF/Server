Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Create a stamp collection for the image stamp
oSVR.AddStampCollection "IMGimage"

' Add an image stamp to the lower right corner of each page.
oSVR.AddStampImage strPath & "Server.ImageInput.jpg", 508.0, 50.0, 64.0, 64.0, True

' Set whether the stamp collection(s) appears in the background or
' foreground
oSVR.StampBackground = 0

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.AddStampImage.pdf")

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