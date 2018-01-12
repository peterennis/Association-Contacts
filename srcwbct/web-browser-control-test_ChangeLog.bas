Option Compare Database
Option Explicit

Public gblnHideFormHeader As Boolean
Public gblnDeveloper As Boolean
Public gintUserId As Integer

' Constants for settings of "WBCT"
Public Const gblnTEST As Boolean = True
Public Const gstrPROJECT_WBCT As String = "WBCT"
Private Const mstrVERSION_WBCT As String = "0.0.2"
Private Const mstrDATE_WBCT As String = "January 11, 2018"

Public Const WBCT_SQL_FRONT_END = False
Public Const WBCT_AZSQL_FRONT_END = False
Public Const WBCT_STAFF_PERMISSIONS = False
Public Const WBCT_SHOW_LOGIN_FORM = False
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = mstrVERSION_WBCT
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = mstrDATE_WBCT
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_WBCT
End Function

Public Sub WBCT_EXPORT(Optional ByVal varDebug As Variant)

    Const THE_FRONT_END_APP = True
    Const THE_SOURCE_FOLDER = ".\srcwbct\"
    Const THE_XML_FOLDER = ".\srcwbct\xml\"
    Const THE_XML_DATA_FOLDER = ".\srcwbct\xmldata\"
    Const THE_BACK_END_SOURCE_FOLDER = "NONE"
    Const THE_BACK_END_XML_FOLDER = "NONE"
    Const THE_BACK_END_DB1 = "NONE"

    On Error GoTo PROC_ERR

    'Debug.Print "THE_BACK_END_DB1 = " & THE_BACK_END_DB1
    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varSrcFldrBe:=THE_BACK_END_SOURCE_FOLDER, _
                        varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER, _
                        varFrontEndApp:=THE_FRONT_END_APP, _
                        varBackEndDbOne:=THE_BACK_END_DB1
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varSrcFldrBe:=THE_BACK_END_SOURCE_FOLDER, _
                        varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER, _
                        varFrontEndApp:=THE_FRONT_END_APP, _
                        varBackEndDbOne:=THE_BACK_END_DB1
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ACDB_EXPORT"
    Resume Next

End Sub
'
'
'=============================================================================================================================
' Tasks:
' %005 -
' %004 -
' %003 -
'=============================================================================================================================
'
'
'20180111 - v002 -
    ' FIXED - %002 - Add rda svgomg logo and use in test (6KB -> 2KB)
    ' FIXED - %001 - First export