VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TPlanInput
    sInstDate   As String
    nInstSeq    As Integer
    sOrderID    As String
    nOrderSeq    As Integer
    nInstQty    As Long
    sExpectDate As String
    sPersonID   As String
    sPatternID  As String
    sStuffCloseClss As String
End Type

Type TPlanInputSub
    sInstDate   As String
    nInstSeq    As Integer
    nProcSeq    As Integer
    sProcessID  As String
    nNeedWidth  As Single
    sInstRemark As String
    sRemark     As String
End Type

Private Const CLASS_NAME As String = "CPlanInput"

Public Property Let Connection(g_adoConnect As Object)
    Set g_adoCon = g_adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Public Function GetPattern() As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sPattern"

        Set GetPattern = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetPatternOne(sPatternID As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sPatternOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sPatternID)
        
        Set GetPatternOne = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

'- 수주별 작업진행
Public Function GetOrderHistory(Optional nChkDate As Integer, Optional sSDate As String, Optional sEDate As String, _
                    Optional nChkCustomID As Integer, Optional sCustomID As String, _
                    Optional nChkArticleID As Integer, Optional sArticleID As String, _
                    Optional nChkOrder As Integer, Optional sOrder As String, _
                    Optional nChkCloseClss As Integer, Optional nChkStuffClose As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sOrderReWork"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCloseClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkStuffClose)

        Set GetOrderHistory = .Execute
    End With
    
    Set adoCmd = Nothing
End Function



Public Function GetOrder(Optional nChkDate As Integer, Optional sSDate As String, Optional sEDate As String, _
                    Optional nChkCustomID As Integer, Optional sCustomID As String, _
                    Optional nChkArticleID As Integer, Optional sArticleID As String, _
                    Optional nChkOrder As Integer, Optional sOrder As String, _
                    Optional nChkCloseClss As Integer, Optional nChkStuffClose As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sOrder"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCloseClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkStuffClose)

        Set GetOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetPlanInput(nChkDate As Integer, sSDate As String, sEDate As String, _
                    nChkCustomID As Integer, sCustomID As String, _
                    nChkArticleID As Integer, sArticleID As String, _
                    nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sPlanInput"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)

        Set GetPlanInput = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetOrderSub(sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sOrderSub"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)

        Set GetOrderSub = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetInstQtyByDate(sDate As String) As Double
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanInput_sInstQtyByDate"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adSingle, adParamOutput, 4, 0)
        
        .Execute
             
        GetInstQtyByDate = IIf(IsNull(.Parameters(1).Value), 0, .Parameters(1).Value)

    End With

    Set adoCmd = Nothing
End Function


Public Function AddNewPlanInPut(tPlan As TPlanInput, tPlanSub() As TPlanInputSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    bError = False
    iSQL = 0
    
    g_adoCon.BeginTrans
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_PlanInput_uOrderPatternID"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, tPlan.sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, tPlan.sPatternID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, tPlan.sStuffCloseClss)

        .Execute
        
        iSQL = iSQL + 1
            
        Call ClearParameter(adoCmd)
            
        .CommandText = "xp_PlanInput_iPlanInput"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tPlan.sInstDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamOutput, 2, tPlan.nInstSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, tPlan.sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tPlan.nOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, tPlan.nInstQty)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tPlan.sExpectDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tPlan.sPersonID)
            
        .Execute
        iSQL = iSQL + 1
        
        tPlan.nInstSeq = .Parameters(1).Value
        
        For i = 0 To UBound(tPlanSub)
             Call ClearParameter(adoCmd)
             
            .CommandText = "xp_PlanInput_iPlanInputSub"
            .Prepared = True
            
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, tPlanSub(i).sInstDate)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tPlan.nInstSeq)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, tPlanSub(i).nProcSeq)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, tPlanSub(i).sProcessID)
            .Parameters.Append .CreateParameter(.CommandText, adSingle, adParamInput, 4, tPlanSub(i).nNeedWidth)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 50, tPlanSub(i).sInstRemark)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 100, tPlanSub(i).sRemark)
            
            .Execute
            
            iSQL = iSQL + 1
        Next i
    End With
    g_adoCon.CommitTrans
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPlanInput.AddNewPlanInput"
LogMessage:
    Set adoCmd = Nothing
    ReDim sLog(iSQL - 1)

    sLog(i) = "xp_PlanInput_uOrderPatternID '" & tPlan.sOrderID & "', '" & tPlan.sPatternID & "', '" & tPlan.sStuffCloseClss
    
    With tPlan
        sLog(i + 1) = "xp_PlanInput_iPlanInput '" & .sInstDate & "', " & .nInstSeq & ", '" & .sOrderID & "', " & .nOrderSeq & ", " & _
                            .nInstQty & ", '" & .sExpectDate & "', '" & .sPersonID & "' "
    End With
    For i = 0 To UBound(tPlanSub)
        With tPlanSub(i)
            sLog(i + 2) = "xp_PlanInput_iPlanInputSub '" & .sInstDate & "', " & tPlan.nInstSeq & ", " & .nProcSeq & ", '" & .sProcessID & "', " & _
                            .nNeedWidth & ", '" & .sInstRemark & "', '" & .sRemark & " "
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewPlanInPut = False
''        Err.Raise Err.Number, CLASS_NAME & ".AddNewPlanInput", Err.Description
    Else
        Call LogService(sLog)

        AddNewPlanInPut = True
    End If
End Function

Public Function DeletePlanInput(sInstDate As String, nInstSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    bError = False
    
    g_adoCon.BeginTrans
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_PlanInput_dPlanInput"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sInstDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, nInstSeq)

        .Execute
        
    End With
    g_adoCon.CommitTrans
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPlanInput.DeletePlanInput"
LogMessage:
    Set adoCmd = Nothing
    ReDim sLog(0)

    sLog(0) = "xp_PlanInput_dPlanInput '" & sInstDate & "', " & nInstSeq & " "
    
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)

        DeletePlanInput = False
''        Err.Raise Err.Number, CLASS_NAME & ".DeletePlanInput", Err.Description
    Else
        Call LogService(sLog)

        DeletePlanInput = True
    End If
End Function

