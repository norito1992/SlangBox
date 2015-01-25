VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "slangGetIcalHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// slangGetIcalHoliday.cls
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
Private tmpFilePath As String
Private tmpFileTime As String

Private Sub Class_Initialize()
    
    baseUri = "https://www.google.com/calendar/ical/ja.japanese%23holiday%40group.v.calendar.google.com/public/basic.ics"
    saveAs = vbNullString
    tmpFilePath = vbNullString
    tmpFileTime = Str(DateDiff("s", "1970/1/1 9:00", Now))
    
End Sub

Private Function getFile() As Boolean

    Dim ret As Long

    saveAs = tmpFilePath & "slangGetIcalHoliday_" & tmpFileTime & ".ics"
    
    ret = URLDownloadToFile(0, baseUri, saveAs, 0, 0)
    If ret = 0 Then
        getFile = True
    Else
        getFile = False
    End If

End Function

Public Function getIcalHolidayData() As Collection

    Dim inStream As Object
    Dim filePath As String
    Dim textStream As Object
    Dim bufLine As String
    Dim flgDate As Boolean
    Dim flgSummary As Boolean
    Dim wkData As Object
    Dim hDataset As Collection
    
    Const STR_DATE As String = "DTSTART;VALUE=DATE:"
    Const STR_SUMMARY As String = "SUMMARY:"
    
    Set hDataset = New Collection

    If Not (getFile()) Then
        MsgBox "ダウンロード失敗", vbExclamation
        End
    End If
    
    Set inStream = CreateObject("ADODB.Stream")
    inStream.Type = 2           '1:binary/2:text
    inStream.Charset = "UTF-8"
    inStream.Open
    inStream.LoadFromFile saveAs
    
    flgDate = False
    flgSummary = False
    Set wkData = New slangGetIcalHolidayType
    Do While inStream.EOS = False
        bufLine = inStream.ReadText(-2) '-1:read all line/-2:read one line
        If STR_DATE = Left(bufLine, Len(STR_DATE)) Then
            wkData.hDate = Format(Replace(bufLine, STR_DATE, ""), "####/##/##")
            flgDate = True
        ElseIf STR_SUMMARY = Left(bufLine, Len(STR_SUMMARY)) Then
            wkData.hName = Replace(bufLine, STR_SUMMARY, "")
            flgSummary = True
        End If
        If flgDate = True And flgSummary = True Then
            hDataset.Add wkData
            Set wkData = New slangGetIcalHolidayType
            flgDate = False
            flgSummary = False
        End If
    Loop
    
    inStream.Close
    Set inStream = Nothing
    
    'delete temporary file
    Kill saveAs
    
    Set getIcalHolidayData = hDataset

End Function

Public Sub setTmpFilePath(argTmpFilePath As String)
    tmpFilePath = argTmpFilePath
End Sub


