' Copyright (c) 2019 ActivePDF, Inc.
' ActivePDF Server

Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' The below font options only work with conversions that
' go through the printer or postscript file conversions

' Whether to embed all fonts other than base14 fonts
oSVR.EmbedAllFonts = true

' Whether or not to embed Base14 fonts
oSVR.EmbedBase14Fonts = false

' Whether or not embedded fonts should be a subset
oSVR.SubsetFonts = true

' If TrueType fonts should be substituting for the version in the
' x:\windows\fonts folder or not
oSVR.SubstituteTTFonts = false

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF(strPath & "Server.Input.ps", _
    strPath & "Server.FontOptions.pdf")

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