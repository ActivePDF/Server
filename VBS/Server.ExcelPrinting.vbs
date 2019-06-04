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
oSVR.NewDocumentName = "Server.ExcelPrinting.pdf"
oSVR.OutputDirectory = strPath

' Start the print job
Set results = oSVR.BeginPrintToPDF()
If results.ServerStatus = 0 Then
    ' Automate Excel to print a document to activePDF Server
    Set objXLS = CreateObject("Excel.Application")
    objXLS.DisplayAlerts = False
    Set objDoc = objXLS.Workbooks.Open(strPath & "Server.Excel.Input.xlsx", _
	    , True, , , , True, , , False, False, , False)
    objDoc.Activate
    objDoc.PrintOut 1, 999, 1, False, oSVR.NewPrinterName, False, False
    objDoc.Close 0
    objXLS.Quit

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