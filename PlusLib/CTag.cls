VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TTag
    sTagID        As String
    sTag          As String
    nWidth        As Long
    nHeight       As Long
    nClss         As Integer
    nDefectClss   As Integer
    nDefHeight    As Integer
    nDefBaseY     As Integer
    nDefBaseX1    As Integer
    nDefBaseX2    As Integer
    nDefBaseX3    As Integer
    nDefGapY      As Integer
    nDefGapX1     As Integer
    nDefGapX2     As Integer
    nDefLength    As Integer
    nDefHCount    As Integer
    nDefBarClss   As Integer
    nGap          As Integer
    sDirect       As String
'    nDefValue(10) As Integer
'    nDefBarClss   As Integer
End Type

Public Type TTagSub
    sName         As String
    nType         As Integer
    nAlign        As Integer
    x             As Integer
    y             As Integer
    nFont         As Integer
    nLength       As Integer
    nHMulti       As Integer
    nVMulti       As Integer
    nRotation     As Integer
    sText         As String
    nSpace        As Integer
    nRelation     As Integer
    nPrevItem     As Integer
    nBarType      As Integer
    nBarHeight    As Integer
    nFigureWidth  As Integer
    nFigureHeight As Integer
    nThickness    As Integer
    sImageFile    As String
    nWidth        As Integer
    nHeight       As Integer
    nVisible      As Integer
End Type

Private Const CLASS_NAME As String = "CTag"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetTag(Optional nChkTagID As Integer, Optional sTagID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "sp_Common_GetTag"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkTagID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sTagID)

        Set GetTag = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetTag"), Err.Description
End Function

Public Function GetTagSub(sTagID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "sp_Common_GetTagSub"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sTagID)

        Set GetTagSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetTagSub"), Err.Description
End Function
