VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDyeCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TWorkPattern
    DyeKind         As Integer
    DyeID           As Integer
    PtNo            As Integer
    PtName          As String
    Section         As Integer
    Seq             As Integer
    ModeNo          As Integer
    SelNo1          As Integer
    SelNo2          As Integer
End Type

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetMode() As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeCon_sMode"
        Set GetMode = .Execute
    End With
    Set adoCmd = Nothing
End Function

Public Function GetPatternGroup(nDyeKind As Integer, nDyeID As Integer, nPTNo As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeCon_sPatternGroup"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeKind)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nPTNo)
        
        Set GetPatternGroup = .Execute
    End With
    Set adoCmd = Nothing
End Function

Public Function GetPattern(nDyeKind As Integer, nDyeID As Integer, nPTNo As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeCon_sPatternAndMode"
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeKind)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nPTNo)
        
        Set GetPattern = .Execute
    End With
    Set adoCmd = Nothing
End Function

Public Function DeletePattern(ByVal nDyeKind As Integer, ByVal nDyeID As Integer, ByVal nPTNo As Integer) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_DyeCon_dDyeWorkPattern"
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeKind)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nPTNo)
        .Execute
    End With
    Set adoCmd = Nothing
    DeletePattern = True
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    DeletePattern = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "삭제도중 에러"
End Function

Public Function InsertPattern(tData As TWorkPattern) As Boolean
    Dim adoCmd As ADODB.Command
    
    On Error GoTo ErrHandler
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeCon_iDyeWorkPattern"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.DyeKind)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.DyeID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.PtNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, tData.PtName)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.Section)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.Seq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.ModeNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.SelNo1)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tData.SelNo2)
        .Execute
    End With
    Set adoCmd = Nothing
    InsertPattern = True
    Exit Function
    
ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    InsertPattern = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "입력도중 에러"
End Function













