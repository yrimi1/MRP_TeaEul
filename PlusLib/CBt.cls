VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Type tBt
    sBTID           As String
    nBTSeq          As Integer
    sCustom         As String
    sCustomID       As String
    sBTNo           As String
    sArticle        As String
    sArticleID      As String
    nColorCnt       As Integer  ' 색상수
    sRecpDate       As String   ' 접수일
    sPersonID       As String   ' 실험자
    RecpPersonID    As String   ' 접수 작성자
    RecpDTime       As String   ' 접수 등록일
    Remark          As String
    sSendDate        As String
    sSendPersonID    As String
End Type

Type TBtSub
    sBTID       As String
    nBTSeq      As Integer
    nColorSeq   As Integer
    sColor      As String
End Type

Type tBtDyeAux
    sBTID       As String
    nBTSeq      As Integer
    nColorSeq   As Integer
    nLot        As Integer
    nDyeAuxSeq  As Integer
    sDyeAuxID   As String
    nRecipeQty  As Single
End Type

Type tBtPrevDyeAux
    sDyeAux     As String
    sDyeAuxID   As String
    nRecipeQty  As Single
End Type

Private Const CLASS_NAME As String = "CBt"


Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetMaxBtNo() As String
    GetMaxBtNo = Format(GetMaxValue("Bt", "BtNo"), "0000000000")
End Function


Public Function GetBtOne(sBTID As String, nBTSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_sBtOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)

        Set GetBtOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetBtOne"), Err.Description
End Function

Public Function GetBtList(nChkDate As Integer, sSDate As String, sEDate As String, nChkSendDate As Integer, sEndDate As String, _
                    nChkCustom As Integer, sCustom As String, nChkArticle As Integer, sArticle As String, _
                    nChkPerson As Integer, sPersonID As String, nChkBTID As Integer, sBTID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_sBtList"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkSendDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustom)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticle)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkPerson)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkBTID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sBTID)

        Set GetBtList = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetBt"), Err.Description
End Function


Public Function GetBtSub(sBTID As String, nBTSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_sBtSub"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)

        Set GetBtSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetBtSub"), Err.Description
End Function



Public Function GetBtDyeAux(sBTID As String, nBTSeq As Integer, nColorSeq As Integer, nLot As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_sBtDyeAuxByLot"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nColorSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nLot)

        Set GetBtDyeAux = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetBtDyeAux"), Err.Description
End Function



Public Function UpdateBTConfirm(sBTID As String, nBTSeq As Integer, nColorSeq As Integer, nChkConfirm As String, _
                                        nConfirmLot As Integer, sDate As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Bt_uBtConfirm"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nColorSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, nChkConfirm)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nConfirmLot)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        
        
        .Execute
     
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.UpdateBTConfirm"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(0)
    
    sLog(0) = "xp_Bt_uBtConfirm '" & sBTID & "', " & nBTSeq & ", " & nConfirmLot & ", '" & sDate & "'"
    
    
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateBTConfirm = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateBTConfirm"), Err.Description
    Else
        Call LogService(sLog)

        UpdateBTConfirm = True
    End If
End Function


Public Function GetPrintBtList(sSDate As String, sEDate As String, nChkCustom As Integer, sCustom As String, _
    nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_pBtList"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustom)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrder)

        Set GetPrintBtList = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetBt"), Err.Description
End Function


' BT 삭제하기 전에 마지막 차수의 BT인지 확인
Public Function GetLastSeq(sBTID As String) As Integer
    Dim adoCmd As ADODB.Command
    Dim nSeq%

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_sBTLastSeq"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, nSeq)
                
        .Execute
        
        GetLastSeq = .Parameters(1).Value
        
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetLastSeq"), Err.Description
End Function



Public Function IsExistBt(sBTID As String) As Boolean
    IsExistBt = IsExistRecord("Bt", "BTID = '" & sBTID & "'")
End Function


Public Function AddNewBt(tBtList As tBt, tBtListSub() As TBtSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim sNewBTID$, sWhere$, nMaxValue
    Dim i%, iSQL%
    Dim NewBTSeq%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        

        ' Insert Inspect
        .CommandText = "xp_Bt_iBt"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sBTID)
        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 8, sNewBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, tBtList.nBTSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tBtList.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tBtList.sBTNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tBtList.sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBtList.nColorCnt)   ' 색상수
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sRecpDate)       ' 접수일
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sPersonID)       ' 실험자
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.RecpPersonID)    ' 접수 등록자
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, tBtList.Remark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sSendDate)       ' 접수일
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sSendPersonID)       ' 실험자
        
        .Execute
        
        tBtList.sBTID = .Parameters(1).Value
        tBtList.nBTSeq = .Parameters(2).Value
                
        iSQL = iSQL + 1

        .CommandText = "xp_Bt_iBtSub"
        .Prepared = True

        For i = 0 To UBound(tBtListSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sBTID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBtList.nBTSeq)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tBtListSub(i).nColorSeq)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tBtListSub(i).sColor)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.AddNewBt"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tBtListSub) + 1)
    With tBtList
        sLog(0) = "xp_Bt_iBt '" & .sBTID & "', '" & .sBTID & "', " & .nBTSeq & ", '" & .sCustomID & ", '" & _
                    .sBTNo & "', '" & .sArticleID & "', " & .nColorCnt & ", '" & .sRecpDate & "', '" & _
                    .sPersonID & "', '" & .RecpPersonID & "', '" & .RecpDTime & "', '" & .Remark & "','" & _
                    .sSendDate & "', '" & .sSendPersonID & "'"
    End With
    For i = 0 To UBound(tBtListSub)
        With tBtListSub(i)
            sLog(i + 1) = "xp_Bt_iBtSub '" & tBtList.sBTID & "', " & tBtList.nBTSeq & ", " & .nColorSeq & ", '" & .sColor & "'"
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewBt = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewBt"), Err.Description
    Else
        Call LogService(sLog)

        AddNewBt = True
    End If
End Function

Public Function UpdateBt(tBtList As tBt, tBtListSub() As TBtSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Bt_uBt"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBtList.nBTSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tBtList.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tBtList.sBTNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tBtList.sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBtList.nColorCnt)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sRecpDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sPersonID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.RecpPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, tBtList.Remark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sSendDate)       ' 접수일
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sSendPersonID)       ' 실험자
        
        .Execute
        iSQL = iSQL + 1

        .CommandText = "xp_Bt_iBtSub"
        .Prepared = True

        For i = 0 To UBound(tBtListSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBtList.sBTID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBtList.nBTSeq)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tBtListSub(i).nColorSeq)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tBtListSub(i).sColor)
            .Execute
            iSQL = iSQL + 1
        Next i
        
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.UpdateBt"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tBtListSub) + 1)
    
    With tBtList
        sLog(0) = "xp_Bt_iBt '" & .sBTID & "', '" & .sBTID & "', " & .nBTSeq & ", '" & .sCustomID & ", '" & _
                    .sBTNo & "', '" & .sArticleID & "', " & .nColorCnt & ", '" & .sRecpDate & "', '" & _
                    .sPersonID & "', '" & .RecpPersonID & "', '" & .Remark & "','" & _
                    .sSendDate & "', '" & .sSendPersonID & "'"
    End With
    
    For i = 0 To UBound(tBtListSub)
        With tBtListSub(i)
            sLog(i + 1) = "xp_Bt_iBtSub '" & tBtList.sBTID & "', " & tBtList.nBTSeq & ", " & .nColorSeq & ", '" & .sColor & "'"
        End With
    Next i
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateBt = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateBt"), Err.Description
    Else
        Call LogService(sLog)

        UpdateBt = True
    End If
End Function


Public Function UpdateBtSend(sBTID As String, nBTSeq As Integer, sEndDate As String, SendPerson As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_uBTSend"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, SendPerson)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.UpdateBtSend"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Bt_uBTSend '" & sBTID & "', " & nBTSeq & ", '" & sEndDate & "', '" & SendPerson & "'"

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateBtSend = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateBtSend"), Err.Description
    Else
        Call LogService(sLog)

        UpdateBtSend = True
    End If
End Function



Public Function UpdateBtUnSend(sBTID As String, nBTSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_uBTUnSend"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)
    
        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.UpdateBtUnSend"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Bt_uBTUnSend '" & sBTID & "', " & nBTSeq

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateBtUnSend = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateBtUnSend"), Err.Description
    Else
        Call LogService(sLog)

        UpdateBtUnSend = True
    End If
End Function



Public Function DeleteBt(sBTID As String, nBTSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Bt_dBt"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.DeleteBt"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Bt_dBt '" & sBTID & "', " & nBTSeq

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteBt = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteBt"), Err.Description
    Else
        Call LogService(sLog)

        DeleteBt = True
    End If
End Function



Public Function DeleteBtDyeAux(sBTID As String, nBTSeq As Integer, nColorSeq As Integer, nLot As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Bt_dBtDyeAux"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBTSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nColorSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nLot)
        
        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.DeleteBtDyeAux"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Bt_dBtDyeAux '" & sBTID & "', " & nBTSeq & ", " & nColorSeq & ", " & nLot

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteBtDyeAux = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteBtDyeAux"), Err.Description
    Else
        Call LogService(sLog)

        DeleteBtDyeAux = True
    End If
End Function



Public Function AddNewBtDyeAux(tBt() As tBtDyeAux) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim sNewBTID$, sWhere$, nMaxValue
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_Bt_iBtDyeAux"
        .Prepared = True

        For i = 0 To UBound(tBt)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBt(i).sBTID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nBTSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nColorSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nLot)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nDyeAuxSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tBt(i).sDyeAuxID)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 7, tBt(i).nRecipeQty)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.AddNewBtDyeAux"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tBt))
    
    For i = 0 To UBound(tBt)
        With tBt(i)
            sLog(i) = "xp_Bt_iBtDyeAux '" & .sBTID & "', " & .nBTSeq & ", " & .nLot & ", " & _
                        .sDyeAuxID & ", " & .nRecipeQty
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewBtDyeAux = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewBtDyeAux"), Err.Description
    Else
        Call LogService(sLog)

        AddNewBtDyeAux = True
    End If
End Function



Public Function UpdateBTDyeAux(tBt() As tBtDyeAux) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim sNewBTID$, sWhere$, nMaxValue
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Bt_dBtDyeAux"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBt(0).sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(0).nBTSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(0).nColorSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(0).nLot)
        
        .Execute
        iSQL = iSQL + 1
                        
        .CommandText = "xp_Bt_iBtDyeAux"
        .Prepared = True

        For i = 0 To UBound(tBt)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tBt(i).sBTID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nBTSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nColorSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nLot)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tBt(i).nDyeAuxSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tBt(i).sDyeAuxID)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 7, tBt(i).nRecipeQty)

            .Execute
            iSQL = iSQL + 1
        Next i

    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CBt.UpdateBTDyeAux"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tBt) + 1)
    With tBt(0)
        sLog(0) = "xp_Bt_iBtDyeAux '" & .sBTID & "', " & .nBTSeq & ", " & .nLot
    End With
    For i = 0 To UBound(tBt)
        With tBt(i)
            sLog(i + 1) = "xp_Bt_iBtDyeAux '" & .sBTID & "', " & .nBTSeq & ", " & .nLot & ", " & _
                        .sDyeAuxID & ", " & .nRecipeQty
        End With
    Next i
    
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateBTDyeAux = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateBTDyeAux"), Err.Description
    Else
        Call LogService(sLog)

        UpdateBTDyeAux = True
    End If
End Function

