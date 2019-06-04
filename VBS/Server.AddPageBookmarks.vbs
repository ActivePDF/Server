' Copyright (c) 2019 ActivePDF, Inc.
' ActivePDF Server

Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Add bookmarks to pages in the PDF
oSVR.AddPageBookmark "Page 1", 0, 1, "Fit"
oSVR.AddPageBookmark "Page 2", 0, 2, "Fit"

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
	strPath & "Server.Input.ps", _
	strPath & "Server.AddPageBookmark.pdf")

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