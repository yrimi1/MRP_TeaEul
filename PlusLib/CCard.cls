VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TCard
    sCardID    As String
    sSplitID   As String
    sOrderID    As String
    nOrderSeq    As Integer
    sCompProcID  As String
    nCompProcSeq As Integer
    sWaitProcID  As String
    nWaitProcSeq As Integer
    sSplitProcID As String
    nSplitProcSeq    As Integer
    sWorkUnitID  As String
    nWorkUnitSeq As Integer
    sBatJaNo As String
    sReWorkClss As String
    sLotNo  As String
    sUseClss    As String
    sEmerClss   As String
    sPauseClss  As String
    nRoll   As Integer
    nQty    As Integer
    sPatternID  As String
    sAfterProc  As String
    sPrePlanProc    As String
    sPostPlanProc   As String
    sModiClss   As String
    nChkOrder   As Integer
    nChkColor   As Integer
    nChkUseClss As Integer
    sNewSplitID As String
    sPersonID   As String
    sThreadName As String
    sStuffCustom    As String
    nStuffWidth As Single
    nStuffDensity   As Integer
    nTubeNo     As Integer
End Type

Type TCardPattern
    sCardID         As String
    sSplitID        As String
    nPlanSeq        As Integer
    sProcessID      As String
    sCompleteClss   As String
    nNeedWidth      As Single
    sInstRemark     As String
    sRemark         As String
    sReWorkClss     As String  '- 재작업 구분
End Type

Type TPlanPattern
    sProcessID      As String
    sCompleteClss   As String
    nNeedWidth      As Single
    sInstRemark     As String
    sRemark         As String
End Type

Type TCardHold
    nCheckDate      As Integer
    sDate           As String
    eDate           As String
    nCheckOrderId   As Integer
    OrderID         As String
    nCheckOrderNo   As Integer
    OrderNo         As String
    nCheckCardID    As Integer
    CardID          As String
    SplitID         As String
    nCheckProcID    As Integer
    WriteProcID     As String
End Type

Type THoldOpinion
    nAffected    As Integer
    WriteDate    As String
    WriteProcID  As String
    WriteSeq     As Integer
    ProcPerson   As String
    ProcDate     As String
    ProcOpinion  As String
    CardID       As String
    SplitID      As String
End Type


Private Const CLASS_NAME As String = "CCard"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Public Function GetProcess(Optional nChkGroup As Integer = 0) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Common_sProcess"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkGroup)
        
        Set GetProcess = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetOrder(nChkCustomID As Integer, sCustomID As String, _
                    nChkArticleID As Integer, sArticleID As String, _
                    nChkOrder As Integer, sOrder As String, _
                    nChkCardID As Integer, sCardID As String, sSplitID, _
                    nChkProcessID As Integer, sProcessID As String, _
                    nChkCloseClss As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_sOrder"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCardID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sSplitID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCloseClss)

        Set GetOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetOrderSub(sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_sOrderSub"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)

        Set GetOrderSub = .Execute

    End With

    Set adoCmd = Nothing
End Function

'''-- 새로운 패턴 가져오기
''Public Function GetCardPatternReWork(ByVal sCardID As String, ByVal sSplitID As String, ByVal sProcessID As String, ByRef bData As String) As ADODB.Recordset
''    Dim adoCmd As ADODB.Command
''
''    Dim bCheckRec As String
''
''    Set adoCmd = New ADODB.Command
''
''    With adoCmd
''        .ActiveConnection = m_adoCon
''        .CommandType = adCmdStoredProc
''        .CommandText = "xp_CardPattern_sReWork"
''
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
''        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sProcessID)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamOutput, 1, bCheckRec)
''
''        Set GetCardPatternReWork = .Execute
''        bData = .Parameters(3).Value
''
''    End With
''
''    Set adoCmd = Nothing
''End Function

Public Function GetCardPattern(sCardID As String, sSplitID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_sCardPattern"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)

        Set GetCardPattern = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetHoldingCard(tHold As TCardHold) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Hold_sHoldcard"
        
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, tHold.nCheckDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold.sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold.eDate)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, tHold.nCheckOrderId)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, tHold.OrderID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, tHold.nCheckOrderNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 24, tHold.OrderNo)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, tHold.nCheckCardID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold.CardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, tHold.SplitID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, tHold.nCheckProcID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, tHold.WriteProcID)
        
        Set GetHoldingCard = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function UpdateCardProcANDRecipe(tHold() As THoldOpinion, tRec As TRecipe, _
                                    tRecSub() As TRecipeSub, tPattern() As TPlanPattern, AfterProc As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, j%, iSQL%, iPlanSeq%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        For i = 0 To UBound(tHold)
            iPlanSeq = 0
            
            .CommandText = "xp_Hold_uProcOpinion"
            .Prepared = True
            Call ClearParameter(adoCmd)
        
            .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamOutput, 1, tHold(i).nAffected)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold(i).WriteDate)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, tHold(i).WriteProcID)
            .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, tHold(i).WriteSeq)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold(i).ProcPerson)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold(i).ProcDate)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 100, tHold(i).ProcOpinion)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tHold(i).CardID)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, tHold(i).SplitID)
            .Execute
            iSQL = iSQL + 1
            
            .CommandText = "xp_plPattern_dplPattern"
            .Prepared = True
            Call ClearParameter(adoCmd)
        
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tHold(i).CardID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, tHold(i).SplitID)
            .Parameters.Append .CreateParameter(, adTinyInt, adParamOutput, 1, iPlanSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tPattern(0).sProcessID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, AfterProc)
        
            .Execute
            iSQL = iSQL + 1
        
            iPlanSeq = .Parameters(2).Value
        
            .CommandText = "xp_Card_iCardPattern"
            .Prepared = True
    
            For j = 0 To UBound(tPattern)
                Call ClearParameter(adoCmd)
                
                iPlanSeq = iPlanSeq + 1
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tHold(i).CardID)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, tHold(i).SplitID)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iPlanSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tPattern(j).sProcessID)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tPattern(j).sCompleteClss)
                .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tPattern(j).nNeedWidth)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, tPattern(j).sInstRemark)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, tPattern(j).sRemark)
                
                .Execute
                iSQL = iSQL + 1
            Next j
            
        Next i
        
        ' Insert Inspect
        .CommandText = "xp_Recipe_iRecipe"
        .Prepared = True
        Call ClearParameter(adoCmd)

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.OrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, tRec.RecipeSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.ModifySeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.RecipeNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tRec.RecipeDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tRec.PersonID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.UnitWght)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tRec.ChunkRate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tRec.ModiClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tRec.Qty)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, tRec.Remark)

        .Execute
        iSQL = iSQL + 1
        ' 재처방 순위 받음
        tRec.RecipeSeq = .Parameters(2).Value

        ' Insert InspectSub
        .CommandText = "xp_Recipe_iRecipeSub"
        .Prepared = True

        For i = 0 To UBound(tRecSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.RecipeSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.ModifySeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRecSub(i).DyeAuxSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tRecSub(i).DyeAuxID)
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 5, tRecSub(i).DyeAuxRate)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCard.UpdateCardProcANDRecipe"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tRecSub) + 1)
    With tRec
        sLog(0) = "xp_Recipe_iRecipeNEW '" & .OrderID & "', '" & .OrderSeq & "', " & .RecipeSeq & ", " & .ModifySeq & ", '" & _
            .RecipeNO & "', '" & .RecipeDate & "', '" & .PersonID & "'"
    End With
    For i = 0 To UBound(tRecSub)
        With tRecSub(i)
            sLog(i + 1) = "xp_Recipe_iRecipeSubNEW '" & tRec.OrderID & "', '" & tRec.OrderSeq & "', " & tRec.RecipeSeq & ", " & _
                tRec.ModifySeq & ", " & .DyeAuxSeq & ", '" & .DyeAuxID & "', " & .DyeAuxRate
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateCardProcANDRecipe = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateCardProcANDRecipe"), Err.Description
    Else
        Call LogService(sLog)

        UpdateCardProcANDRecipe = True
    End If
End Function

Public Function UpdateCardChange(TItem As TCard) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_Card_uCardChange"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem.sSplitID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nOrderSeq)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 2, TItem.nRoll)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.nQty)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem.sUseClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sReWorkClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sEmerClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nChkOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nChkColor)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, TItem.sModiClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nChkUseClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.sThreadName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, TItem.sStuffCustom)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nStuffWidth)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.nStuffDensity)
        
        .Execute

    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCard.UpdateCardChange"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(0)
    With TItem
        sLog(0) = "xp_Card_uCardChange '" & .sCardID & "', '" & .sSplitID & "', '" & .sOrderID & "', " & .nOrderSeq & ", " & _
                    .nRoll & ", " & .nQty & ", '" & .sUseClss & "', '" & .sReWorkClss & "', '" & .sEmerClss & "', " & _
                    .nChkOrder & ", " & .nChkColor & ", '" & .sPersonID & "', '" & .sModiClss & "', " & .nChkUseClss & ", '" & _
                    .sThreadName & "', '" & .sStuffCustom & "', " & .nStuffWidth & ", " & .nStuffDensity & " "
    End With
    
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)

        UpdateCardChange = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateCardChange"), Err.Description
    Else
        Call LogService(sLog)

        UpdateCardChange = True
    End If
End Function


Public Function UpdateCardDivide(TItem() As TCard) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%
    
    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_Card_uCardDivide"
        .Prepared = True

        For i = 0 To UBound(TItem)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, i)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem(i).sCardID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem(i).sSplitID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem(i).sNewSplitID)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 2, TItem(i).nRoll)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem(i).nQty)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem(i).sPersonID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, TItem(i).sModiClss)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem(i).sPatternID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem(i).sUseClss)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem(i).nChkUseClss)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem(i).nTubeNo)
            
            .Execute
            iSQL = iSQL + 1
            Call ClearParameter(adoCmd)
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCard.UpdateCardDivide"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(TItem))
    For i = 0 To UBound(TItem)
        With TItem(i)
            sLog(i) = "xp_Card_uCardDivide " & i & ", '" & .sCardID & "', '" & .sSplitID & "', '" & .sNewSplitID & "', " & .nRoll & ", " & _
                            .nQty & ", ," & .sPersonID & "', '" & .sModiClss & "', ," & .sPatternID & "', '" & .sUseClss & "', " & _
                            .nChkUseClss & ", " & .nTubeNo & " "
        End With
    Next i
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateCardDivide = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateCardDivide"), Err.Description
    Else
        Call LogService(sLog)

        UpdateCardDivide = True
    End If
End Function

Public Function UpdateCardPattern(TItem As TCard, tItemSub() As TCardPattern) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%
    
    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_uCardPattern"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem.sSplitID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWaitProcID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nWaitProcSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem.sUseClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nChkUseClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, TItem.sPrePlanProc)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, TItem.sPostPlanProc)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, TItem.sAfterProc)

        .Execute
        iSQL = iSQL + 1

        .CommandText = "xp_Card_iCardPattern"
        .Prepared = True

        For i = 0 To UBound(tItemSub)
            Call ClearParameter(adoCmd)
                        
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tItemSub(i).sCardID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, tItemSub(i).sSplitID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItemSub(i).nPlanSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tItemSub(i).sProcessID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tItemSub(i).sCompleteClss)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nNeedWidth)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, tItemSub(i).sInstRemark)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, tItemSub(i).sRemark)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tItemSub(i).sReWorkClss)
            
            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCard.UpdateCardPattern"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tItemSub) + 1)
    With TItem
        sLog(0) = "xp_Card_uCardPattern '" & .sCardID & "', '" & .sSplitID & "', '" & .sWaitProcID & "', " & .nWaitProcSeq & ", '" & _
                        .sUseClss & "', " & .nChkUseClss & ", '" & .sPersonID & "', '" & .sPrePlanProc & "', '" & .sPostPlanProc & "', '" & _
                        .sAfterProc & "' "
    End With
    
    For i = 0 To UBound(tItemSub)
        With tItemSub(i)
            sLog(i + 1) = "xp_Card_iCardPattern '" & .sCardID & "', '" & .sSplitID & "', " & .nPlanSeq & ", '" & .sProcessID & "', '" & _
                        .sCompleteClss & "', " & .nNeedWidth & ", '" & .sInstRemark & "', '" & .sRemark
        End With
    Next i
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateCardPattern = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateCardPattern"), Err.Description
    Else
        Call LogService(sLog)

        UpdateCardPattern = True
    End If
End Function

Public Function GetWorkCard(sCardID As String, sSplitID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkCard_sWorkCard"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)

        Set GetWorkCard = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetWorkCardSub(sCardID As String, sSplitID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkCard_sWorkCardSub"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)

        Set GetWorkCardSub = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetWorkCardResult(sCardID As String, sSplitID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkCard_sWorkCardResult"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)

        Set GetWorkCardResult = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetWorkProcessPlan(sCardID As String, sSplitID As String, sPatternID As String) As String
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkCard_sProcessPlan"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sPatternID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamOutput, 255, "")
        
        .Execute
        
        GetWorkProcessPlan = IIf(IsNull(.Parameters(3).Value), "", .Parameters(3).Value)
    End With

    Set adoCmd = Nothing
End Function

Public Function GetWorkPattern(sCardID As String, sSplitID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkCard_sPattern"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)
        
        Set GetWorkPattern = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetCardAllList(sCardID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = " xp_Card_sResultProcCard"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        
        Set GetCardAllList = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetCardModiList(sCardID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = " xp_Card_sModiProcCard"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        
        Set GetCardModiList = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetCardWorking(sDate As String, sTime As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sProcWorking"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sTime)

        Set GetCardWorking = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetCardWaiting(sProcID As String, iClss As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sProcWaiting"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iClss)

        Set GetCardWaiting = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetThreadName(sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WizCard_sThreadName"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)

        Set GetThreadName = .Execute
    End With
    Set adoCmd = Nothing

End Function

Public Function GetStuffCustom(sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WizCard_sStuffINCustom"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)

        Set GetStuffCustom = .Execute
    End With
    Set adoCmd = Nothing
End Function

Public Function GetCardTube(nChkProcess As Integer, sProcessID As String, _
                            nChkCustom As Integer, sCustomID As String, _
                            nChkArticle As Integer, sArticleID As String, _
                            nChkOrder As Integer, sOrder As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_sTube"
       
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkProcess)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        
        Set GetCardTube = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


