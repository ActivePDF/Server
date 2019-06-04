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
oSVR.NewDocumentName = "Server.WordPrinting.pdf"
oSVR.OutputDirectory = strPath

' Start the print job
Set result = oSVR.BeginPrintToPDF()
If result.ServerStatus = 0 Then
    ' Automate Word to print a document to Server
    Set objWord = CreateObject("Word.Application")
    objWord.DisplayAlerts = False
    Set objDoc = objWord.Documents.Open(strPath & "Server.Word.Input.doc", False, True)
    Set objWordDialog = objWord.Dialogs(97)
    objWordDialog.Printer = oSVR.NewPrinterName
    objWordDialog.DoNotSetAsSysDefault = 1
    objWordDialog.Execute
    objDoc.PrintOut False
    objDoc.Close False
    objWord.Quit False

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