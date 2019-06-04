' Copyright (c) 2019 ActivePDF, Inc.
' ActivePDF Server

Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oSVR = CreateObject("APServer.Object")

' Set the quality options for the created PDF
' For custom settings to take effect set the configuration to custom
oSVR.PredefinedSetting = 0

' Specifies if ASCII85 encoding should be applied to binary streams
oSVR.ASCIIEncode = true

' Automatically control the page orientation based on text flow
oSVR.AutoRotate = true

' Color Image Quality Settings
oSVR.ColorImageDownsampleThreshold = 1
oSVR.ColorImageDownsampleType = 0
oSVR.ColorImageFilter = 2
oSVR.ColorImageResolution = 72

' Specifies if CMYK colors should be converted to RGB
oSVR.ConvertCMYKToRGB = true

' Gray Image Quality Settings
oSVR.GrayImageDownsampleThreshold = 1
oSVR.GrayImageDownsampleType = 0
oSVR.GrayImageFilter = 2
oSVR.GrayImageResolution = 72

' Monochrome Image Quality Settings
oSVR.MonoImageDownsampleThreshold = 1
oSVR.MonoImageDownsampleType = 0
oSVR.MonoImageFilter = 2
oSVR.MonoImageResolution = 72

' Set whether existing halftone settings should be preserved
oSVR.PreserveHalftone = 0

' Set whether existing overprint settings should be preserved
oSVR.PreserveOverprint = 0

' Set how transfer functions from the input file are handled
oSVR.PreserveTransferFunction = 0

' Set the DPI for the created PDF
oSVR.Resolution = 300.0

' Set whether the UCRandBGInfo, from the input file, should be preserved
oSVR.UCRandBGInfo = 0

' Convert the PostScript file into PDF
Set result = oSVR.ConvertPSToPDF(strPath & "Server.Input.ps", _
	strPath & "Server.OutputQuality.pdf")

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