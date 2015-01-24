VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "slangGetHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// slangGetHoliday.cls
'//
'// Created by slangsoft on 2015/01/25.
'// Copyright (c) 2015 Slang Software. All rights reserved.
'//
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private baseUri As String
Private saveAs As String
Private dom As MSXML2.DOMDocument
Private tmpFilePath As String
Private tmpFileTime As String

Private Sub Class_Initialize()

    baseUri = "https://www.google.com/calendar/feeds/ja.japanese%23holiday%40group.v.calendar.google.com/public/basic"
    tmpFilePath = vbNullString
    tmpFileTime = Str(DateDiff("s", "1970/1/1 9:00", Now))
    saveAs = vbNullString
    Set dom = New MSXML2.DOMDocument60
    dom.async = False
'    dom.setProperty "ServerHTTPRequest", True
    
End Sub

Private Function getXmlFile() As Boolean

    Dim Ret As Long

    saveAs = tmpFilePath & "slangGetHoliday_" & tmpFileTime & ".xml"
    
    Ret = URLDownloadToFile(0, baseUri, saveAs, 0, 0)
    If Ret = 0 Then
        getXmlFile = True
    Else
        getXmlFile = False
    End If

End Function

Public Function getHolidayData() As Object

    If Not (getXmlFile()) Then
        MsgBox "É_ÉEÉìÉçÅ[Éhé∏îs", vbExclamation
        End
    End If
    
    If Not (dom.Load(saveAs)) Then
        Dim strErr As String
        With dom.parseError
            strErr = "XMLì«çûé∏îs" & vbCrLf & _
            vbCrLf & _
            "ErrorCode    : " & .errorCode & vbCrLf & _
            "ErrorReason  : " & .reason & vbCrLf & _
            "Line         : " & .Line & vbCrLf & _
            "LinePosition : " & .linepos & vbCrLf & _
            "FilePosition : " & .filepos & vbCrLf & _
            "SourceText   : " & .srcText & vbCrLf & _
            "DocumentUrl  : " & .url
        End With
        MsgBox strErr, vbExclamation
        End
    End If
    Set getHolidayData = dom.childNodes.Item(1).childNodes
    
    'delete temporary file
'    Kill openXmlFullPath

End Function

Public Sub setTmpFilePath(argTmpFilePath As String)
    tmpFilePath = argTmpFilePath
End Sub
