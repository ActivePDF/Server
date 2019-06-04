' Copyright (c) 2019 ActivePDF, Inc.
' ActivePDF Server 

Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Path and filename of output
oSVR.NewDocumentName = "Server.PowerPointPrinting.pdf"
oSVR.OutputDirectory = strPath

' Start the print job
Set result = oSVR.BeginPrintToPDF()
If result.ServerStatus = 0 Then
    ' Automate PowerPoint to print a document to activePDF Server
    Set objPPT = CreateObject("PowerPoint.Application")
    Set objDoc = objPPT.Presentations.Open(strPath & "Server.PowerPoint.Input.pptx", -1, 0, 0)
    Set objOptions = objDoc.PrintOptions
    objOptions.ActivePrinter = oSVR.NewPrinterName
    objOptions.PrintInBackground = 0
    objDoc.PrintOut 1, 9999, "", 1, 0
    objDoc.Saved = 1
    objDoc.Close
    objPPT.Quit

    ' Wait(seconds) for job to complete
    Set result = oSVR.EndPrintToPDF(30)
End If

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