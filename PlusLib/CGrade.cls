VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TGrade
    GradeID        As String
    Grade          As String
End Type

Private Const CLASS_NAME As String = "CGrade"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2002-03-18 (MON)
'* UPDATE :
'*
'* --->>> CodeType 속성을 먼저 설정하세요.
'*
'********************************************************************************
Public Function GetGrade() As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sGrade"

        Set GetGrade = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetGrade"), Err.Description
End Function

Public Function AddNewGrade(tGr As TGrade) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Grade_iGrade"

        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 1, tGr.GradeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tGr.Grade)

        .Execute

        tGr.GradeID = .Parameters(0).Value
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CGrade.AddNewGrade"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    With tGr
        sLog(0) = "xp_Grade_iGrade '" & .GradeID & "', '" & .Grade & "'"
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        AddNewGrade = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewGrade"), Err.Description
    Else
        Call LogService(sLog)

        AddNewGrade = True
    End If
End Function

Public Function UpdateGrade(tGr As TGrade) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Grade_uGrade"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tGr.GradeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tGr.Grade)
    
        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CGrade.UpdateGrade"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    With tGr
        sLog(0) = "xp_Grade_uGrade '" & .GradeID & "', '" & .Grade & "'"
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateGrade = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateGrade"), Err.Description
    Else
        Call LogService(sLog)

        UpdateGrade = True
    End If
End Function

Public Function DeleteGrade(sGradeID As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Grade_dGrade"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sGradeID)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CGrade.DeleteGrade"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    
    sLog(0) = "xp_Grade_dGrade '" & sGradeID & "'"

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteGrade = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteGrade"), Err.Description
    Else
        Call LogService(sLog)

        DeleteGrade = True
    End If
End Function
