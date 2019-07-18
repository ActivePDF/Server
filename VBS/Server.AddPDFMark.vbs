Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Add PDF marks to be used in the converted PDF
' PDF Mark Reference:
' http:'www.adobe.com/content/dam/Adobe/en/devnet/acrobat/pdfs/pdfmark_reference.pdf

' Notes (Page 20 - PDF Marks Reference)
oSVR.AddPDFMark "[ /SrcPg 1 /Rect [32 32 216 144] /Open false /Title (ActivePDF Comment) /Contents (Note Type Comment Example.) /Color [1 0 0] /Subtype /Text /ANN pdfmark"

' Free Text (Page 17 - PDF Marks Reference)
' Used for the text in the link created below
oSVR.AddPDFMark "[ /SrcPg 1 /Rect [262 26 350 46] /Contents (ActivePDF.com) /DA ([0 0 1] rg /Helv 12 Tf) /BS << /W 0 >> /Q 1 /Subtype /FreeText /ANN pdfmark"

' Links (Page  - PDF Marks Reference)
' Add a link around the activePDF.com text created above
oSVR.AddPDFMark "[ /SrcPg 1 /Rect [262 26 350 46] /Contents (ActivePDF.com) /BS << /W 0 >> /Action << /Subtype /URI /URI (http:'ActivePDF.com) >> /Subtype /Link /ANN pdfmark"

' Bookmarks (Page 26 - PDF Marks Reference)
oSVR.AddPDFMark "[ /Count -5 /Title (ActivePDF Server - AddPDFMark) /Page 1 /F 2 /OUT pdfmark"
oSVR.AddPDFMark "[ /Page 1 /View [/Fit] /Title (AddPDFMark - Page 1) /C [0 0 0] /F 2 /OUT pdfmark"
oSVR.AddPDFMark "[ /Page 2 /View [/Fit] /Title (AddPDFMark - Page 2) /C [0 0 0] /F 2 /OUT pdfmark"

' Document Information (Page 28 - PDF Marks Reference)
oSVR.AddPDFMark "[ /Title (PDF Marks) /Author (ActivePDF Server) /Subject (PDF Marks) /Keywords (pdfmark, server, example) /DOCINFO pdfmark"

' Document View Options (Page 29 - PDF Marks Reference)
oSVR.AddPDFMark "[ /PageMode /UseOutlines /Page 1 /View [/Fit] /DOCVIEW pdfmark"

' Document Open Options
oSVR.AddPDFMark "[ {Catalog} << /ViewerPreferences << /HideToolbar true  /HideMenubar true >> >> /PUT pdfmark"

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.AddPDFMark.pdf")

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