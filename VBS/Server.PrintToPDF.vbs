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
oSVR.NewDocumentName = "Server.PrintToPDF.pdf"
oSVR.OutputDirectory = strPath

' Start the print job
Set result = oSVR.BeginPrintToPDF()
If result.ServerStatus = 0 Then
    ' Here is where you can print to activePDF Server to create
    ' a PDF from any print job, set your application to print to
    ' a static activePDF Server printer or call oSVR.NewPrinterName
    ' to dynamically create a new printer on the fly
    ' This example simply calls oSVR.TestPrintToPDF for testing purposes
    Set result = oSVR.TestPrintToPDF("Hello World!")
    If result.ServerStatus = 0 Then
        ' Wait(seconds) for job to complete
        Set result = oSVR.EndPrintToPDF(30)
    End If
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