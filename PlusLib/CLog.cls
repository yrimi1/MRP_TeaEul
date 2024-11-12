VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TLOG
    LogID  As Long
    LogSeq As Integer
End Type

Private Const CLASS_NAME As String = "CLog"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetComputerID() As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "MRPLog2.dbo.xp_sComputerID"

        Set GetComputerID = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetComputerID"), Err.Description
End Function

Public Function GetUserID() As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "MRPLog2.dbo.xp_sUserID"

        Set GetUserID = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetUserID"), Err.Description
End Function

Public Function GetLog(bErrLog As Boolean, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkComputerID As Integer, sComputerID As String, nChkUserID As Integer, sUserID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "MRPLog2.dbo." & IIf(bErrLog, "xp_sErrLog", "xp_sLog")

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkComputerID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 15, sComputerID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkUserID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 15, sUserID)

        Set GetLog = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetLog"), Err.Description
End Function

Public Function DeleteLog(bErrLog As Boolean, logData() As TLOG) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "MRPLog2.dbo." & IIf(bErrLog, "xp_dErrLog", "xp_dLog")
        .Prepared = True

        For i = 0 To UBound(logData)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, logData(i).LogID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, logData(i).LogSeq)

            .Execute

            Call ClearParameter(adoCmd)
        Next i
    End With
    Set adoCmd = Nothing

    g_adoCon.CommitTrans

    DeleteLog = True

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    g_adoCon.RollbackTrans

    DeleteLog = False
    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteLog"), Err.Description
End Function
