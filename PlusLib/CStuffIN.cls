VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStuffIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' 입고 레코드 구조체
Type TStuffIN
    sJobFlag        As String
    sStuffDate      As String  '[3] 입고 일자
    sStuffClss      As String   '입고구분
    nStuffSeq       As Integer  '입고순번
    sCustomID       As String   '발주처코드
    sCustom         As String   '원단입고처
    sOrderID        As String   '작지번호(OrderID - 관리번호)
    sOrderNo        As String   'Order NO
    sArticleID      As String   '품명
    sWorkID         As String   '가공구분코드
    sExManNo        As String   '외주가공구분코드
    sUnitClss       As String   '입고단위
    nTotRoll        As Integer  '입고 총 절수
    nTotQty         As Long     '입고 총 수량
    sRemark         As String   '비고사항
    sThreadName     As String   '사종
    ADDClss         As String   '추가분
    sOrderFlag      As String   '거래구분 0:A건, 1:B건, 2:시가공, 3:샘플
    nChkDateUpd     As Integer  '1: 입고일자수정, 0: 입고일자 비 수정
    sOLDDate        As String
    sOLDClss        As String
    nOLDSeq         As Integer
    sSubulWidthID   As String   '원단폭
End Type


' 입고 세부 레코드 구조체
Type TStuffINSub
    sStuffDate      As String  '[3] 입고 일자
    sStuffClss      As String   '입고구분
    nStuffSeq       As Integer  '입고순번
    nRollNo         As Integer  '절번호
    sColorID        As String   '색상코드
    sColor          As String   '색상명
    nQty            As Integer  '입고수량
End Type

' 수불 정산처리 CloseClss
Type TSubulClose
    sSubulDate As String
    sSubulClss As String
    sIOClss As String
    nSeq  As Integer
    sCloseClss As String
End Type


Type TAssign
    JobFlag       As String
    StuffDate     As String
    StuffClss     As String
    StuffSeq      As Integer
    OrderID       As String
    AssignSeq     As Integer
    Qty           As Long
    Roll          As Integer
    AssignDate    As String
End Type

' 반품 세부 레코드 구조체
Type TStuffINReturn
    sStuffDate      As String  '[3] 반품일자
    sStuffClss      As String   '입고구분
    nStuffSeq       As Integer  '입고순번
    nStuffNO        As Integer  '일련번호
    sStuffPart      As String   '1.정단, 2.난단 구분
    nRollNo         As Integer  '절번호
    nQty            As String   '반품수량

End Type



Private Const CLASS_NAME As String = "CStuffIN"


'*****************************************************************
' 반품 거래명세서용 업체
'
'*****************************************************************
Public Function GetCustomInfo(ByVal sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Stuffin_sCustomInfo"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        
        Set GetCustomInfo = .Execute
    End With

    Set adoCmd = Nothing

End Function

Public Function UpdateOutWareMemo(ByVal sOrderID As String _
                        , ByVal nOutSeq As Integer _
                        , ByVal sMemo As String) As Boolean
                        
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean

    On Error GoTo ErrHandler
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OutWare_uMemo"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4000, sMemo)
        
        .Execute
    End With
    bError = False
    GoTo LogMessage

ErrHandler:
    bError = True

LogMessage:
    Set adoCmd = Nothing
    If bError Then
        UpdateOutWareMemo = False
        Err.Raise Err.Number, CLASS_NAME, ".UpdateStuffINMemo", Err.Description
    Else
        UpdateOutWareMemo = True
    End If
    
End Function

Public Function UpdateStuffINMemo(ByVal sStuffDate As String _
                        , ByVal sStuffClss As String _
                        , ByVal nStuffSeq As Integer _
                        , ByVal sMemo As String) As Boolean
                        
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean

    On Error GoTo ErrHandler
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_uMemo"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sStuffDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nStuffSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4000, sMemo)
        
        .Execute
    End With
    bError = False
    GoTo LogMessage

ErrHandler:
    bError = True

LogMessage:
    Set adoCmd = Nothing
    If bError Then
        UpdateStuffINMemo = False
        Err.Raise Err.Number, CLASS_NAME, ".UpdateStuffINMemo", Err.Description
    Else
        UpdateStuffINMemo = True
    End If
    
End Function


Public Function GetSubulMemo(ByVal IOClss As String _
                        , ByVal sStuffDate As String, ByVal sStuffClss As String, ByVal nStuffSeq As Integer _
                        , ByVal sOrderID As String, ByVal nOutSeq As Integer _
                        , ByRef sMemo As String) As Boolean
                        
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean

    On Error GoTo ErrHandler
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sMemo"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, IOClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sStuffDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nStuffSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamOutput, 4000, sMemo)
        
        .Execute
        sMemo = .Parameters(6).Value
        
    End With
    bError = False
    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CStuffIN.GetSubulMemo"
LogMessage:
    Set adoCmd = Nothing
    If bError Then
        GetSubulMemo = False
''        Err.Raise Err.Number, CLASS_NAME, ".GetSubulMemo", Err.Description
    Else
        GetSubulMemo = True
    End If
    
End Function

Public Function DoSubulAdjust(TSubulClose() As TSubulClose) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%
    Dim nRows As Integer
    Dim nStuffSeq As Integer

  '  On Error GoTo ErrHandler

    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    With adoCmd
        .CommandText = "xp_subul_uAdjust"
        .Prepared = True
        
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_subul_uAdjust"
        
        For i = 0 To UBound(TSubulClose)
            Call ClearParameter(adoCmd)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TSubulClose(i).sSubulDate)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, TSubulClose(i).sSubulClss)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, TSubulClose(i).sIOClss)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, TSubulClose(i).nSeq)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, TSubulClose(i).sCloseClss)
            
            .Execute
        Next i
    End With
    
    g_adoCon.CommitTrans
    DoSubulAdjust = True
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    DoSubulAdjust = False
    Set adoCmd = Nothing
End Function


Public Function GetSubulAdjust(ByVal sSDate As String _
                        , ByVal sEDate As String _
                        , ByVal nChkCustomID As Integer _
                        , ByVal sCustomID As String _
                        , ByVal nCheckOrderId As Integer _
                        , ByVal sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_subul_sAdjust"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nCheckOrderId)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        
        Set GetSubulAdjust = .Execute
    End With

    Set adoCmd = Nothing
End Function


Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub
'----------------------------------------------------
' 입고에 관련된 Order 레코드 1개를 가져온다.
'
'-----------------------------------------------------
Public Function GetStuffINByOrder(ByVal sOrder As String)
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
   
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sStuffINByOrder"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrder)
        
        Set GetStuffINByOrder = .Execute
    End With

    Set adoCmd = Nothing
End Function

' 입고건에 대한 수주할당 내역

Public Function GetStuffInAssignIN(ByVal StuffDate As String, ByVal StuffClss As String, ByVal StuffSeq As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
   
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sAssignIN"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, StuffDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, StuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, StuffSeq)
        
        Set GetStuffInAssignIN = .Execute
    End With

    Set adoCmd = Nothing

End Function
Public Function GetStuffINOne(ByVal sStuffDate As String, ByVal sStuffClss As String, ByVal sStuffSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sStuffINOne"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sStuffDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, , sStuffSeq)
        
        Set GetStuffINOne = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetStuffINSubONE(ByVal sStuffDate As String, ByVal sStuffClss As String, ByVal nStuffSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sStuffINSubOne"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sStuffDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, , nStuffSeq)
        
        Set GetStuffINSubONE = .Execute
    End With

    Set adoCmd = Nothing
End Function
'********************************************************
'-- 생지입고 명세서
'********************************************************
Public Function GetStuffINList(ByVal nChkDate As Integer _
                            , ByVal sDate As String _
                            , ByVal eDate As String _
                            , ByVal nChkCustom As Integer _
                            , ByVal sCustomID As String _
                            , ByVal nChkArticleID As Integer _
                            , ByVal sArticleID As String _
                            , ByVal nChkStuffClss As Integer _
                            , ByVal sStuffClss As String _
                            , ByVal nChkOrderID As Integer _
                            , ByVal sOrderID As String _
                            , ByVal nChkOrderNon As Integer _
                            , ByVal nOrderFlag As Integer) As ADODB.Recordset
                                
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sStuffINList"
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkOrderNon)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nOrderFlag)
        
        Set GetStuffINList = .Execute
    End With

    Set adoCmd = Nothing

End Function
Public Function GetStuffINByCustom(nChkDate As Integer _
                        , sSDate As String _
                        , sEDate As String _
                        , nChkCustomID As Integer _
                        , sCustomID As String _
                        , nChkArticleID As Integer _
                        , sArticleID As String _
                        , nChkStuffClss As Integer _
                        , sStuffClss As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sStuffINByCustom"
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sStuffClss)
        
        Set GetStuffINByCustom = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetStuffIN(ByVal xpName As String _
                        , nChkDate As Integer, sSDate As String, sEDate As String _
                        , nChkCustomID As Integer, sCustomID As String _
                        , nChkArticleID As Integer, sArticleID As String _
                        , nChkStuffClss As Integer, sStuffClss As String _
                        , nCheckOrderId As Integer, sOrderID As String _
                        , nCheckNon As Integer _
                        , nAddCheck As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = xpName
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nCheckOrderId)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nCheckNon)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nAddCheck)
        
        Set GetStuffIN = .Execute
    End With

    Set adoCmd = Nothing
End Function

'''Public Function AddNewStuffIN(StuffIN As TStuffIN, StuffINSub() As TStuffINSub, nSeq As Integer) As Boolean
'''    Dim adoCmd As ADODB.Command
'''    Dim sLog() As String, bError As Boolean
'''    Dim i%, iSQL%
'''    Dim nRows As Integer
'''    Dim nStuffSeq As Integer
'''
'''    On Error GoTo ErrHandler
'''
'''    g_adoCon.BeginTrans
'''
'''    nSeq = StuffIN.nStuffSeq
'''
'''    Set adoCmd = New ADODB.Command
'''    With adoCmd
'''        .ActiveConnection = g_adoCon
'''        .CommandType = adCmdStoredProc
'''        .CommandText = "xp_StuffIN_iuStuffIN"
'''
'''        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, nRows)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffIN.sJobFlag)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffIN.sStuffDate)
'''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.sStuffClss)
'''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInputOutput, 4, nSeq)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, StuffIN.sCustomID)
'''        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, StuffIN.sCustom)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 40, StuffIN.sUnitClss)
'''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.nTotRoll)
'''        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, StuffIN.nTotQty)
'''        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, StuffIN.sRemark)
'''        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, StuffIN.sThreadName)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, StuffIN.sOrderID)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, StuffIN.sArticleID)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, StuffIN.sWorkID)
'''        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, StuffIN.sOrderNo)
'''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffIN.ADDClss)
'''
'''        .Execute
'''
'''''        '----- StuffINSub 데이터 등록 Stored Procedure Call
''''
'''        nRows = .Parameters(0).Value
'''        nSeq = .Parameters(4).Value
'''
'''        .CommandText = "xp_StuffIN_iuStuffINSub"
'''        .Prepared = True
'''
'''        For i = 0 To UBound(StuffINSub)
'''            Call ClearParameter(adoCmd)
'''
'''            .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, nRows)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffIN.sStuffDate)
'''            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.sStuffClss)
'''            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 4, nSeq)
'''            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, i + 1)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 5, StuffINSub(i).sColorID)
'''            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, StuffINSub(i).sColor)
'''            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, StuffINSub(i).nQty)
'''
'''            .Execute
'''            iSQL = iSQL + 1
'''        Next i
'''    End With
'''
'''    g_adoCon.CommitTrans
'''    bError = False
'''
'''    GoTo LogMessage
'''
'''ErrHandler:
'''    g_adoCon.RollbackTrans
'''    bError = True
'''
'''LogMessage:
'''    Set adoCmd = Nothing
'''    If bError Then
'''        AddNewStuffIN = False
'''        Err.Raise Err.Number, CLASS_NAME, ".AddNewStuffIN", Err.Description
'''    Else
'''        AddNewStuffIN = True
'''    End If
'''End Function
'********************************************

Public Function AddNewStuffIN(ByVal StuffClss As String _
                            , ByVal nCount As Integer _
                            , ByRef StuffIN As TStuffIN _
                            , ByRef StuffINSub() As TStuffINSub _
                            , ByRef StuffInReturn() As TStuffINReturn _
                            , ByRef nSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%
    Dim nRows As Integer
    Dim nStuffSeq As Integer

   ' On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    nSeq = StuffIN.nStuffSeq
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        
        .CommandText = "xp_StuffIN_iuStuffIN"
        
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, nRows)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffIN.sJobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffIN.sStuffDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.sStuffClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInputOutput, 4, nSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, StuffIN.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, StuffIN.sCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 40, StuffIN.sUnitClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.nTotRoll)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, StuffIN.nTotQty)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, StuffIN.sRemark)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, StuffIN.sThreadName)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, StuffIN.sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, StuffIN.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, StuffIN.sWorkID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, StuffIN.sOrderNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffIN.ADDClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffIN.sOrderFlag)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.nChkDateUpd)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffIN.sOLDDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffIN.sOLDClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 4, StuffIN.nOLDSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, StuffIN.sSubulWidthID)
        
        
        
        .Execute
        nSeq = .Parameters(4).Value
        
        If StuffClss = "1" Then
            '----- StuffINSub 데이터 등록 Stored Procedure Call
    
            If nCount > 0 Then
                nRows = .Parameters(0).Value
                nSeq = .Parameters(4).Value
        
                
                .CommandText = "xp_StuffIN_iuStuffINSub"
                .Prepared = True
        
                For i = 0 To UBound(StuffINSub)
                    If Val(StuffINSub(i).nQty) <> 0 Then
                        Call ClearParameter(adoCmd)
            
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, nRows)
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffIN.sStuffDate)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.sStuffClss)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 4, nSeq)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, i + 1)
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 5, StuffINSub(i).sColorID)
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, StuffINSub(i).sColor)
                        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, StuffINSub(i).nQty)
            
                        .Execute
                        iSQL = iSQL + 1
                    End If
                Next i
            End If
        Else
            '----- StuffInReturn 데이터 등록 Stored Procedure Call
            If nCount > 0 Then
                .CommandText = "xp_StuffIN_iuReturn"
                .Prepared = True
                
                For i = 0 To UBound(StuffInReturn)
                    If Len(StuffInReturn(i).nQty) <> 0 Then
                        Call ClearParameter(adoCmd)
            
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffIN.sStuffDate)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffIN.sStuffClss)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nSeq)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, StuffInReturn(i).nStuffNO)
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffInReturn(i).sStuffPart)
                        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, StuffInReturn(i).nRollNo)
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, StuffInReturn(i).nQty)
            
                        .Execute
                        iSQL = iSQL + 1
                    End If
                Next i
            End If
        End If
    End With

    g_adoCon.CommitTrans
    bError = False
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CStuffIN.AddNewStuffIN"
LogMessage:
    Set adoCmd = Nothing
    If bError Then
        AddNewStuffIN = False
''        Err.Raise Err.Number, CLASS_NAME, ".AddNewStuffIN", Err.Description
    Else
        AddNewStuffIN = True
    End If
End Function

Public Function DeleteStuffIN(ByVal StuffDate As String, ByVal StuffClss As String, ByVal StuffSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    
    On Error GoTo ErrHandler

    bError = False
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_dStuffIN"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, StuffSeq)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CStuffIN.DeleteStuffIN"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
 '   sLog(0) = "xp_StuffIN_dStuffIN '" & sOrderID & "', " & CStr(nStuffID)
    
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)

        DeleteStuffIN = False
''        Err.Raise Err.Number, CLASS_NAME & ".DeleteStuffIN", Err.Description
    Else
        Call LogService(sLog)

        DeleteStuffIN = True
    End If
End Function


''Public Function PrintStuffIN(nChkDate As Integer, sSDate As String, sEDate As String, _
''                            nChkCustomID As Integer, sCustomID As String, _
''                            nChkArticleID As Integer, sArticleID As String) As ADODB.Recordset
''
''    Dim adoCmd As ADODB.Command
''
''    Set adoCmd = New ADODB.Command
''    With adoCmd
''        .ActiveConnection = g_adoCon
''        .CommandType = adCmdStoredProc
''        .CommandText = "xp_StuffIN_pStuffIN"
''
''        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
''        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
''        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
''
''        Set PrintStuffIN = .Execute
''    End With
''
''    Set adoCmd = Nothing
''End Function


'Public Function GetStuffInNotOrder(nChkDate As Integer, sSDate As String, sEDate As String, _
'                    nChkCustomID As Integer, sCustomID As String, _
'                    nChkArticleID As Integer, sArticleID As String, _
'                    nChkStuffClss As Integer, sStuffClss As String) As Recordset
'    Dim adoCmd As ADODB.Command
'
'    Set adoCmd = New ADODB.Command
'    With adoCmd
'        .ActiveConnection = g_adoCon
'        .CommandType = adCmdStoredProc
'        .CommandText = "xp_StuffIn_sStuffInNotOrder"
'
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkStuffClss)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sStuffClss)
'
'        Set GetStuffInNotOrder = .Execute
'    End With
'
'    Set adoCmd = Nothing
'End Function
'------------ StuffAssign delete 처리 --------------------------
Public Function DeleteAssign(TItem As TAssign) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    bError = False

    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_dAssign"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.StuffDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.StuffClss)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, TItem.StuffSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.OrderID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, TItem.AssignSeq)
        .Execute
    End With

    g_adoCon.CommitTrans
    GoTo LogMessage

ErrHandler:
    bError = True
    g_adoCon.RollbackTrans
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CStuffIN.DeleteAssign"
LogMessage:
    Set adoCmd = Nothing

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)

        DeleteAssign = False
''        Err.Raise Err.Number, CLASS_NAME & ".DeleteAssign", Err.Description
    Else

        DeleteAssign = True
    End If
End Function
'******************************************************************
' 입고에 대한 수주할당 내역 보고
'
'
'
'******************************************************************
Public Function GetStuffINByAssignList(ByVal sSDate As String _
                                     , ByVal sStuffClss As String _
                                     , ByVal sStuffSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrorHandler
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sStuffINAssign"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, sStuffSeq)

        Set GetStuffINByAssignList = .Execute
    End With
    
    Set adoCmd = Nothing
    Exit Function
ErrorHandler:

End Function

Public Function GetStuffInNotOrder(nChkDate As Integer, sSDate As String, sEDate As String, _
                    nChkCustomID As Integer, sCustomID As String, _
                    nChkArticleID As Integer, sArticleID As String, _
                    nChkStuffClss As Integer, sStuffClss As String, _
                    sAssignClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIn_sStuffInNotOrder"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sAssignClss)

        Set GetStuffInNotOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function
Public Function GetStuffInOrder(sCustomID As String, sArticleID As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIn_sStuffInOrder"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)

        Set GetStuffInOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

''Public Function UpdateStuffINOrder(tItem As TStuffIN) As Boolean
''    Dim adoCmd As ADODB.Command
''    Dim sLog() As String, bError As Boolean
''
''    On Error GoTo ErrHandler
''
''    bError = False
''
''    g_adoCon.BeginTrans
''    Set adoCmd = New ADODB.Command
''    With adoCmd
''        .ActiveConnection = g_adoCon
''        .CommandType = adCmdStoredProc
''        .CommandText = "xp_StuffIN_uStuffINOrder"
''
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tItem.sStuffDate)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tItem.sStuffClss)
''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItem.nStuffSeq)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tItem.sOrderID)
''
''        .Execute
''    End With
''
''    g_adoCon.CommitTrans
''    GoTo LogMessage
''
''ErrHandler:
''    bError = True
''    g_adoCon.RollbackTrans
''LogMessage:
''    Set adoCmd = Nothing
''
''    ' Make Log Message
''    ReDim sLog(0)
''    With tItem
''        sLog(0) = "xp_StuffIN_uStuffINOrder '" & .sStuffDate & "', " & .sStuffClss & "', " & .nStuffSeq & ", '" & .sOrderID & "' "
''    End With
''
''    If bError Then
''        Call ErrLogService(sLog, Err.Number, Err.Description, 0)
''
''        UpdateStuffINOrder = False
''        Err.Raise Err.Number, CLASS_NAME & ".UpdateStuffINOrder", Err.Description
''    Else
''        Call LogService(sLog)
''
''        UpdateStuffINOrder = True
''    End If
''End Function

Public Function UpdateStuffINOrder(TItem As TAssign) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    bError = False

    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_iuStuffAssign"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.JobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.StuffDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.StuffClss)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, TItem.StuffSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.OrderID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, TItem.AssignSeq)
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.Qty)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.Roll)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.AssignDate)

        .Execute
    End With

    g_adoCon.CommitTrans
    GoTo LogMessage

ErrHandler:
    bError = True
    g_adoCon.RollbackTrans
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CStuffIN.UpdateStuffINOrder"
LogMessage:
    Set adoCmd = Nothing

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)

        UpdateStuffINOrder = False
''        Err.Raise Err.Number, CLASS_NAME & ".UpdateStuffINOrder", Err.Description
    Else

        UpdateStuffINOrder = True
    End If
End Function


Public Function GetStuffINDetail(ByVal rsData As Recordset, ByRef nRollVar())
    Dim AAA(), II%, nPreRoll%, nCnt%, JJ%, KK%
    Dim nRollCnt%, nRollQty%, sRollStr$
    
    nRollCnt = 0: nRollQty = 0: sRollStr = ""

    ReDim AAA(rsData.RecordCount)
    
    II = 0
    Do Until rsData.EOF
        AAA(II) = rsData!Qty
        II = II + 1
        rsData.MoveNext
    Loop
    
    II = 0: nPreRoll = 0: nCnt = 0: KK = 0
    
    Do While II < UBound(AAA)
    
        sRollStr = AAA(II)
        
        For JJ = II To UBound(AAA)
            If AAA(II) = AAA(JJ) Then
                nCnt = nCnt + 1
            Else
                
                Exit For
            End If
        Next JJ
        II = JJ
        ReDim Preserve nRollVar(KK)
        nRollVar(KK) = sRollStr & " * " & nCnt
        
        KK = KK + 1
        nCnt = 0
    Loop
    
    
    '--- 90 * 1 인 경우  * 1을 없애는 처리
    For II = 0 To UBound(nRollVar)
        If Val(Mid(nRollVar(II), InStr(nRollVar(II), "*") + 1)) = 1 Then
            nRollVar(II) = Left(nRollVar(II), InStr(nRollVar(II), "*") - 1)
        End If
    Next II

End Function

'********************************************************
'-- 생지반품 명세서 출력용 RecordSet
'********************************************************
Public Function GetStuffINReturnGoods(ByVal StuffDate As String _
                            , ByVal StuffClss As String _
                            , ByVal StuffSeq As Integer _
                            , ByRef RsHeader As Recordset) As Boolean
                                
    On Error GoTo ErrHandler
    
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sReturnGoods"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, StuffDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, StuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, StuffSeq)
        
        Set RsHeader = .Execute
        
        GetStuffINReturnGoods = True
        
''        .CommandText = "xp_StuffIN_sReturnGoodsDetail"
''        .Prepared = True
''
''        Call ClearParameter(adoCmd)
''
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, StuffDate)
''        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, StuffClss)
''        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, StuffSeq)
''        Set GetStuffINReturnGoods = .Execute
        
    End With

    Set adoCmd = Nothing
    Exit Function
ErrHandler:
    Set adoCmd = Nothing

    GetStuffINReturnGoods = False
    MsgBox ("CStuffIN.GetStuffINReturnGoods: " & Err.HelpContext)
End Function

'반품 절수 나타내기
Public Function GetStuffInReturn(ByVal sStuffDate As String, ByVal sStuffClss As String, ByVal nStuffSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_StuffIN_sReturn"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sStuffDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sStuffClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, , nStuffSeq)
        
        Set GetStuffInReturn = .Execute
    End With

    Set adoCmd = Nothing
End Function




