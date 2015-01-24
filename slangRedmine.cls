VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "slangRedmine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// slangRedmine.cls
'//
'// Created by slangsoft on 2015/01/24.
'// Copyright (c) 2015 Slang Software. All rights reserved.
'//
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private baseUri As String
Private apiKey As String
Private dom As MSXML2.DOMDocument
Private tmpFilePath As String
Private tmpFileTime As String

Private Sub Class_Initialize()

    baseUri = vbNullString
    apiKey = vbNullString
    tmpFilePath = vbNullString
    tmpFileTime = Str(DateDiff("s", "1970/1/1 9:00", Now))
    Set dom = New MSXML2.DOMDocument60
    dom.async = False
    
End Sub

Private Function getXmlFile(xmlName As String) As Boolean

    Dim downloadFile As String
    Dim saveAs As String
    Dim Ret As Long
    
    downloadFile = baseUri & xmlName & ".xml?key=" & apiKey
    saveAs = tmpFilePath & xmlName & "_" & tmpFileTime & ".xml"
    
    If xmlName = "" Then Exit Function
    
    Ret = URLDownloadToFile(0, downloadFile, saveAs, 0, 0)
    If Ret = 0 Then
        getXmlFile = True
    Else
        getXmlFile = False
    End If

End Function

Public Function getUserData() As Object

    Dim xmlName As String
    Dim openXmlFullPath As String
    
    xmlName = "users"
    openXmlFullPath = tmpFilePath & xmlName & "_" & tmpFileTime & ".xml"

    If Not (getXmlFile(xmlName)) Then
        MsgBox "É_ÉEÉìÉçÅ[Éhé∏îs", vbExclamation
        End
    End If
    
    openXmlFullPath = openXmlFullPath

    If Not (dom.Load(openXmlFullPath)) Then
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
    Set getUserData = dom.childNodes.Item(1).childNodes
    
    'delete temporary file
    Kill openXmlFullPath

End Function

Public Sub setBaseUri(argBaseUri As String)
    baseUri = argBaseUri
End Sub

Public Sub setApiKey(argApiKey As String)
    apiKey = argApiKey
End Sub

Public Sub setTmpFilePath(argTmpFilePath As String)
    tmpFilePath = argTmpFilePath
End Sub
