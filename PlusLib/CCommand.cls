VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Type TCommand
    sOrderID    As String
    nOrderIDSeq As Integer
    nOrderSeq   As String
    Instdate    As String
    InstQty     As Long
    nApplyLoss  As Single
    sPersonID    As String
End Type

Private Const CLASS_NAME As String = "CCommand"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetCommandByOrder(nChkDate As Integer, sSDate As String, sEDate As String, _
                         nChkCustom As Integer, sCustom As String, nChkArticle As Integer, sArticle As String, _
                         nChkOrder As Integer, sOrder As String, nChkClose As Integer) As ADODB.Recordset
                        
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Command_sCommandByOrder"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustom)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticle)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticle)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sOrder)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkClose)
        Set GetCommandByOrder = .Execute
    
    End With

    Set adoCmd = Nothing
    Exit Function
    
ErrHandler:
    Set adoCmd = Nothing
    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetCommandByOrder"), Err.Description
End Function


Public Function GetCommandByColor(sOrder As String) As ADODB.Recordset
                        
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Command_sCommandByColor"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sOrder)
        
        Set GetCommandByColor = .Execute
    
    End With

    Set adoCmd = Nothing
    Exit Function
    
ErrHandler:
    Set adoCmd = Nothing
    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetCommandByColor"), Err.Description
End Function



Public Function GetCommandByDate(sOrder As String, nOrderSeq As Integer) As ADODB.Recordset
                        
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Command_sCommandByDate"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sOrder)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)
        
        Set GetCommandByDate = .Execute
    
    End With

    Set adoCmd = Nothing
    Exit Function
    
ErrHandler:
    Set adoCmd = Nothing
    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetCommandByDate"), Err.Description
End Function


' 투입 마감 설정및 해제
Public Function UpdateStuffClose(sOrderID As String, sDate As String, nChkClose As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '품명 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Command_uStuffCloseClss"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkClose)
        
        .Execute
        
        
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    sLog(0) = "xp_Command_uStuffCloseClss '" & sOrderID & "'. '" & sDate & "', " & nChkClose
   
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateStuffClose = False
        Err.Raise Err.Number, "CArticle.UpdateStuffClose", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdateStuffClose = True
    End If

End Function


'
Public Function AddNewCommand(NewCommand As TCommand) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '품명 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Command_iCommandByColor"
                    
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewCommand.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, NewCommand.nOrderIDSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewCommand.nOrderSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewCommand.Instdate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewCommand.InstQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, NewCommand.nApplyLoss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewCommand.sPersonID)
        
        .Execute
        
        NewCommand.nOrderIDSeq = .Parameters(1).Value
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With NewCommand
        sLog(0) = "xp_Command_iCommandByColor '" & .sOrderID & "'. " & .nOrderIDSeq & ", " & .nOrderSeq & ", '" & _
                        .Instdate & "', " & .InstQty & ", " & .nApplyLoss & ", '" & .sPersonID
    End With
   
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCommand = False
        Err.Raise Err.Number, "CCommand.AddNewCommand", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewCommand = True
    End If

End Function



