Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Add bookmarks to pages in the PDF
oSVR.AddPageBookmark "Parent", 1, 1, "Fit"
oSVR.AddPageBookmark "Child 1", 0, 2, "Fit"

' Add bookmarks to URLs
oSVR.AddURLBookmark "Parent URL Bookmark", 1, "http:'www.activepdf.com"
oSVR.AddURLBookmark "Child URL Bookmark", 0, _
                    "https:'www.activepdf.com/products/server"

' Add bookmarks pointing to pages in external PDF
' Both Local and UNC file paths are accepted
oSVR.AddLinkedPDFBookmark "Parent PDF Bookmark", 1, _
                          strPath & "Server.Sample.pdf", 1, "Fit"
oSVR.AddLinkedPDFBookmark "Child PDF Bookmark", 0, _
                          strPath & "Server.Sample.pdf", 2, "Fit"

' Add bookmarks pointing to any external file
' Both Local and UNC file paths are accepted
oSVR.AddFileBookmark "Parent File Bookmark", 1, _
                     strPath & "Server.PowerPoint.Input.pptx"
oSVR.AddFileBookmark "Child  File Bookmark", 0, _
                     strPath & "Server.Word.Input.doc"

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF( _
    strPath & "Server.Input.ps", _
	strPath & "Server.AddBookmarks.pdf")

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