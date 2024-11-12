VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' 변경이력
'------------------------------------------------------------------------------
'
'요청ID : S_201203_삼우_02
'요청일자 : 2012.03.05
'요청내용 : 오더별 명세 출력되게
'변경내용 : cmdOrderPrint 버튼 추가
'
'******************************************************************************

Type TSubul
    SubulDate As String
    SubulClss As String
    IOClss As String
    Seq As Integer
    OrderNo As String
    OrderSeq As Integer
    CustomID As String
    ArticleID As String
    QtyUnit As String
    Custom As String
    WorkID As String
    Cnt As Integer
    Qty As Double
    Price As Double
    OutRealQty As Double
    StuffDate As String
    StuffClss As String
    StuffSeq As Integer
    OrderID As String
    OutSeq As Integer
End Type

Type TSbStock
    sBasisDate           As String
    sCustomID            As String
    sArticleID           As String
    sSubulWdithID        As String
    sTaxClss             As String
    nStockQty            As Long
    sStockUnitClss       As String
    
End Type


Private Const CLASS_NAME As String = "CSubul"
' 선언부
Type TProcCost
    YearMon As String
    CustomID As String
    FromDate As String
    ToDate As String
    AdjustClss As String
    TaxClss As String
    PrevYearMon As String
    ExchRate   As Single   '환율
End Type

' 계상처리 관련 선언2
Type TMonthCustom
    sYearMon As String
    sCustomID As String
End Type

' 계상처리 관련 선언3
Type TProcCostDet
    sBasisYearMon As String
    sCustomID   As String
    nProcSeq    As Integer
    sOrderNo    As String
    sArticleID  As String
    sSubulWidthID   As String
    sWorkID     As String
    nWorkUnitPrice As Single
    nTempUnitPrice As Single
    sOrderID    As String
    sAdjustClss As String
    sUnitClss   As String
    nSumQty     As Long
    nSumQtyY    As Long
    nOrderQty   As Long
    nOutQty     As Long
    sTaxClss    As String
    nPrevMonSumQty As Long
    sDealClss   As String
    sPriceClss  As String
    nExchRate   As Single
    nAmountWon  As Long
    nAmountDollar   As Single
    nTax As Long
    nTotalPrice As Long
    sTaxSeq As String
    sOrderFlag  As String
End Type


Type TOutWareRec
    OrderID     As String
    OutSeq      As Integer
    OutClss     As String
    WorkID      As String
    ExchRate    As Single
    UnitPrice   As Single
    OutCustom   As String
    LossRate    As Single
    LossQty     As Single
    OutRoll     As Integer
    OutQty      As Long
    OutRealQty  As Long
    OutDate     As String
    ResultDate  As String
    OutTime     As String
    BoOutClss   As String
    BoConfirmClss   As String
    BoConfirmDate   As String
    LoadTime    As String
    TranNo      As String
    TranSeq     As Integer
    TelNo       As String
    Remark      As String
    OutType     As String
End Type

Type TOutWareSubRec
    OrderID   As String
    OutSeq    As Integer
    OutSubSeq   As Integer
    OrderSeq    As Integer
    RollSeq     As Integer
    BoxNo       As Integer
    LotNo       As String
    OutQty      As Long
End Type



Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

''Public Function GetSubul(sDate As String, _
''            nChkCustomID As Integer, sCustomID As String, _
''            nChkArticleID As Integer, sArticleID As String, _
''            nChkOrder As Integer, sOrder As String) As Recordset
''    Dim adoCmd As ADODB.Command
''
''    On Error GoTo ErrHandler
''
''    Set adoCmd = New ADODB.Command
''    With adoCmd
''        .ActiveConnection = g_adoCon
''        .CommandType = adCmdStoredProc
''        .CommandText = "xp_Subul_sSubul"
''
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
''        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
''
''        Set GetSubul = .Execute
''    End With
''    Set adoCmd = Nothing
''
''    Exit Function
''
''ErrHandler:
''    Set adoCmd = Nothing
''
''    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetSubul"), Err.Description
''End Function

'''Public Function UpdateSubulSet(tItem() As TSubul) As Boolean
'''    Dim adoCmd As ADODB.Command
'''    Dim sLog() As String, bError As Boolean
'''    Dim i%, j%, iSQL%
'''
'''    On Error GoTo ErrHandler
'''
'''    g_adoCon.BeginTrans
'''
'''    Set adoCmd = New ADODB.Command
'''    With adoCmd
'''        .ActiveConnection = g_adoCon
'''        .CommandType = adCmdStoredProc
'''
'''        .CommandText = "xp_Subul_uSubulSet"
'''        .Prepared = True
'''
'''        For i = 0 To UBound(tItem)
'''            Call ClearParameter(adoCmd)
'''
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tItem(i).SubulDate)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tItem(i).SubulClss)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tItem(i).IOClss)
'''            .Parameters.Append .CreateParameter(, adInteger, adParamInputOutput, 2, tItem(i).Seq)
'''            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, tItem(i).OrderNo)
'''            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, tItem(i).OrderSeq)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tItem(i).CustomID)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tItem(i).ArticleID)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tItem(i).QtyUnit)
'''            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, tItem(i).Custom)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tItem(i).WorkID)
'''            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tItem(i).Cnt)
'''            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, tItem(i).Qty)
'''            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, tItem(i).Price)
'''            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, tItem(i).OutRealQty)
'''            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tItem(i).OrderID)
'''            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItem(i).OutSeq)
'''
'''            .Execute
'''
'''            tItem(i).Seq = tItem(0).Seq
'''        Next i
'''
'''    End With
'''
'''    g_adoCon.CommitTrans
'''
'''    GoTo LogMessage
'''
'''ErrHandler:
'''    g_adoCon.RollbackTrans
'''    bError = True
'''
'''LogMessage:
'''    Set adoCmd = Nothing
'''
'''    ReDim sLog(UBound(tItem))
'''    For i = 0 To UBound(tItem)
'''        With tItem(i)
'''            sLog(i) = "xp_Subul_uSubulSet '" & .SubulDate & "', '" & .SubulClss & "', '" & .IOClss & "', " & .Seq & ", '" & _
'''                .OrderNo & "', " & .OrderSeq & ", '" & .CustomID & "', '" & .ArticleID & "', '" & .QtyUnit & "', '" & _
'''                .Custom & "', '" & .WorkID & "', " & .Cnt & ", " & .Qty & ", " & .Price & ", " & .OutRealQty & ", '" & _
'''                .OrderID & "', " & .OutSeq & " "
'''        End With
'''    Next i
'''
'''    If bError Then
'''        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)
'''
'''        UpdateSubulSet = False
'''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateSubulSet"), Err.Description
'''    Else
'''        Call LogService(sLog)
'''
'''        UpdateSubulSet = True
'''    End If
'''End Function


'  정산서
Public Function GetBillByDate(sStartDate As String, sEndDate As String, _
                            sJWMon As String, nChkCustom As Integer, sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Account_sBillList"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sStartDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 6, sJWMon)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        
        Set GetBillByDate = .Execute
    End With
    
    Set adoCmd = Nothing
End Function



' 청구서
Public Function GetBillAccountByDate(sStartDate As String, sEndDate As String, _
                            sStartJW As String, sEndJW As String, nChkCustom As Integer, sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Account_sAccountList"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sStartDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sStartJW)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEndJW)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        
        Set GetBillAccountByDate = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetOutWareReport(ByVal sDate As String, ByVal eDate As String, ByVal nChkOrder As Integer, ByVal sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OutWare_sReport"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 35, sOrderID)
        
        Set GetOutWareReport = .Execute
    End With
    
    Set adoCmd = Nothing
'
End Function


'-- Order별 출고명세서 - 입고내역

Public Function GetOutWareReportStuffIN(ByVal sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OutWare_sReportStuff"
        
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 24, sOrderID)
        
        Set GetOutWareReportStuffIN = .Execute
    End With
    
    Set adoCmd = Nothing
'
End Function


'-- Order별 출고명세서 - Order내역

Public Function GetOutWareReportOrder(ByVal sDate As String, ByVal eDate As String, ByVal nChkOrder As Integer, ByVal sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OutWare_sReportOrder"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 30, sOrderID)
        
        Set GetOutWareReportOrder = .Execute
    End With
    
    Set adoCmd = Nothing
'
End Function


'***********************************************************************************
'*
'* Author : frogfinger
'* Date : 2004-01-02
'*
'* 오더별 수불 집계
'*
'***********************************************************************************
Public Function GetAccountByOrder(sDate As String, nChkCustom As Integer, sCustomID As String, _
                            nChkArticle As Integer, sArticleID As String, nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Account_sTotalByOrder"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticle)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 15, sOrder)
        
        Set GetAccountByOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


'***********************************************************************************
'*
'* Author : frogfinger
'* Date : 2004-01-02
'*
'* 거래처, 품명별 수불 집계
'*
'***********************************************************************************
Public Function GetAccountByCustom(sDate As String, nChkCustom As Integer, sCustomID As String, _
                                        nChkArticle As Integer, sArticleID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Account_sTotalByCustom"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticle)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
                
        Set GetAccountByCustom = .Execute
    End With
    
    Set adoCmd = Nothing
End Function
''---------------------------------------------------------------------------------------------------
''---------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------
' 유상수
Public Function GetProcCost(sYearMon As String, nChkCustomID As Integer, sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sProcCost"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, sYearMon)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        

        Set GetProcCost = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetProcCost"), Err.Description
End Function



'***************************************************************
' Order별 청구서
'***************************************************************
Public Function GetProcCostReport(ByVal sDate As String _
                                , ByVal nChkCustomID As Integer _
                                , ByVal sCustomID As String _
                                , ByVal sTaxClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_sbProcCost_sReport"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, sDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sTaxClss)

        Set GetProcCostReport = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetProcCostReport"), Err.Description

End Function


'***************************************************************
' Order별 청구서
' S_201203_삼우_02 에 의한 추가
'***************************************************************
Public Function GetProcCostReportOrder(ByVal sDate As String _
                                , ByVal nChkCustomID As Integer _
                                , ByVal sCustomID As String _
                                , ByVal sTaxClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_sbProcCost_sReportByOrder"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, sDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sTaxClss)

        Set GetProcCostReportOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetProcCostReportOrder"), Err.Description

End Function

'***************************************************************
' Order별 수불( 입고/ 출고 ) 레코드 읽어 오기
'***************************************************************
Public Function GetsubulsReportSub(ByVal OrderID As String _
                            , ByVal sDate As String _
                            , ByVal eDate As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sReportSub"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, OrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        Set GetsubulsReportSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetSubulReport"), Err.Description

End Function



' 수불명세서 입고,출고에 따라 해당 Order의 수량 집계
Public Function GetSubulOrderSub(ByVal IOClss As String _
                                , ByVal StuffDate As String _
                                , ByVal StuffClss As String _
                                , ByVal StuffSeq As Integer _
                                , ByVal OrderID As String _
                                , ByVal OutSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sOrderSub"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, IOClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, StuffDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, StuffClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, StuffSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, OutSeq)

        Set GetSubulOrderSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetSubulOrderSub"), Err.Description
End Function



'-----------------------------------------------------------------------
' 유상수
Public Function GetProcCostDetail(ByVal sYearMon As String, ByVal sCustomID As String, ByVal sTaxClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sProcCostDetail"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, sYearMon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sTaxClss)

        Set GetProcCostDetail = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetProcCostDetail"), Err.Description
End Function

Public Function AddNewProcCostDet(TItem As TProcCostDet) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_iProcCostDet"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TItem.sBasisYearMon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sOrderNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sSubulWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nWorkUnitPrice)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sOrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem.sAdjustClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sUnitClss)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nSumQty)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nSumQtyY)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nOrderQty)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nOutQty)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, TItem.sTaxClss)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.nPrevMonSumQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sDealClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPriceClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nExchRate)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nAmountWon)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 8, TItem.nAmountDollar)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nTax)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nTotalPrice)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderFlag)
            
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddNewProcCostDet = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    AddNewProcCostDet = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "입력처리중 에러"
End Function



'-----------------------------------------------------------------------
' 유상수
Public Function GetProcCostDetailData(sYearMon As String, sCustomID As String, nProcSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sProcCostDetailData"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, sYearMon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nProcSeq)

        Set GetProcCostDetailData = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetProcCostDetailData"), Err.Description
End Function


Public Function UpdateProcCostDet(TItem As TProcCostDet) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_uProcCostDet"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TItem.sBasisYearMon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nProcSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sOrderNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sSubulWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nWorkUnitPrice)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nTempUnitPrice)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nSumQty)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nSumQtyY)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, TItem.sTaxClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sDealClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TItem.sAdjustClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPriceClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nExchRate)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nAmountWon)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nAmountDollar)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nTax)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem.nTotalPrice)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderFlag)
            
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateProcCostDet = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    UpdateProcCostDet = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "수정처리중 에러"
End Function



'-----------------------------------------------------------------------
' 유상수
Public Function GetStockList(sStartDate As String, sEndDate As String, nChkCustom As Integer, _
                            sCustomID As String, nChkArticle As Integer, sArticleID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sStockList"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sStartDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)

        Set GetStockList = .Execute
    End With
    
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetStockList"), Err.Description
End Function




'-----------------------------------------------------------------------
' 유상수
Public Function GetStockDataOne(TSbStock As PlusLib2.TSbStock) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sStockDataOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TSbStock.sBasisDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TSbStock.sTaxClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TSbStock.sSubulWdithID)
        
        Set GetStockDataOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetStockDataOne"), Err.Description
End Function




'-----------------------------------------------------------------------
' 유상수
Public Function GetNotCloseOrder(nChkDate As Integer, sStartDate As String, sEndDate As String, nChkCustom As Integer, _
                                sCustomID As String, nChkArticle As Integer, sArticleID As String, _
                                nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sNotCloseOrder"
    
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sStartDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, sOrder)

        Set GetNotCloseOrder = .Execute
    End With
    
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetNotCloseOrder"), Err.Description
End Function



'-----------------------------------------------------------------------
' 계상처리
' 유상수
Public Function UpdateProcCost(TCostItem() As TProcCost) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_uProcCost"
        .Prepared = True

        For i = 0 To UBound(TCostItem) - 1
            Call ClearParameter(adoCmd)
        
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TCostItem(i).YearMon)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TCostItem(i).CustomID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TCostItem(i).FromDate)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TCostItem(i).ToDate)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TCostItem(i).AdjustClss)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, TCostItem(i).TaxClss)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TCostItem(i).PrevYearMon)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TCostItem(i).ExchRate)
            
            .Execute
        
        Next i
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateProcCost = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    UpdateProcCost = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "계상처리 작성중 에러"
End Function



'-----------------------------------------------------------------------
' 유상수
Public Function AddNewStock(TSbStock As PlusLib2.TSbStock) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_iStock"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TSbStock.sBasisDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, "1")
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, "0")
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TSbStock.nStockQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TSbStock.sStockUnitClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TSbStock.sTaxClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TSbStock.sSubulWdithID)
            
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddNewStock = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    AddNewStock = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "재고 입력중 에러"
End Function


' 함수
Public Function AddNewOutData(ow() As TOutWareRec, owSub() As TOutWareSubRec) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%, j%
    
    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_iOutware"

        For i = 0 To UBound(ow) - 1
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, ow(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, ow(i).OutSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow(i).OutClss)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow(i).WorkID)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow(i).ExchRate)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow(i).UnitPrice)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, ow(i).OutCustom)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow(i).LossRate)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow(i).LossQty)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, ow(i).OutRoll)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, ow(i).OutQty)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, ow(i).OutRealQty)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow(i).OutDate)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow(i).ResultDate)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow(i).BoOutClss)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow(i).OutTime)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow(i).LoadTime)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow(i).OutType)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, ow(i).Remark)
            
            .Execute
            
            ow(i).OutSeq = .Parameters(1).Value
            
            ' Insert Sub
            .CommandText = "xp_Outware_iOutwareSub"
            .Prepared = True
    
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, owSub(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, ow(i).OutSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OutSubSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).RollSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).BoxNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, owSub(i).LotNo)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 5, owSub(i).OutQty)

            .Execute
        Next i
    End With

    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddNewOutData = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    AddNewOutData = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "출고 입력중 에러"
End Function


'-----------------------------------------------------------------------
' 유상수
Public Function UpdateStock(TSbStock As PlusLib2.TSbStock) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_uStock"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TSbStock.sBasisDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, "1")
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, "0")
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TSbStock.nStockQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TSbStock.sStockUnitClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TSbStock.sTaxClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TSbStock.sSubulWdithID)
            
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateStock = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    UpdateStock = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "재고 수정중 에러"
End Function




'-----------------------------------------------------------------------
' 유상수
Public Function DeleteStock(TSbStock As PlusLib2.TSbStock) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_dStock"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TSbStock.sBasisDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TSbStock.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TSbStock.sTaxClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TSbStock.sSubulWdithID)
            
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    DeleteStock = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    DeleteStock = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "재고 삭제중 에러"
End Function

'***********************************************************************************
'*
'* Author : wzchs
'* Date : 2004-01-02
'*
'* 납품사실증명원
'*
'***********************************************************************************
Public Function GetDeliveryReport(ByVal sMonth As String _
                                , ByVal nChkCustom As Integer _
                                , ByVal sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_sbProcCostDet_sDraft"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 6, sMonth)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
                
        Set GetDeliveryReport = .Execute
    End With
    
    Set adoCmd = Nothing
End Function



'***********************************************************************************
'*
'* Author : wzchs
'* Date : 2004-01-02
'*
'* 수출용 원자재 매도 확약서( Offer Sheet )
'*
'***********************************************************************************
Public Function GetDeliverySaleReport(ByVal sMonth As String _
                                , ByVal nChkCustom As Integer _
                                , ByVal sCustomID As String _
                                , ByVal nChkClss As Integer _
                                , ByRef nTotQtyYds As Long _
                                , ByRef nTotPirce As Long) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DeliverySale_Report"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 6, sMonth)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkClss)
        .Parameters.Append .CreateParameter(.CommandText, adDouble, adParamOutput, 4, nTotQtyYds)
        .Parameters.Append .CreateParameter(.CommandText, adDouble, adParamOutput, 4, nTotPirce)
                
        Set GetDeliverySaleReport = .Execute
        nTotQtyYds = .Parameters(4).Value
        nTotPirce = .Parameters(5).Value
        
    End With
    
    Set adoCmd = Nothing
End Function

'***************************************************************
' 가공료 집계표
'***************************************************************
Public Function GetProcCostCustom(ByVal sDate As String, ByVal eDate As String, ByVal sTaxClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_sbProcCost_sCustom"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, eDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sTaxClss)

        Set GetProcCostCustom = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetProcCostCustom"), Err.Description

End Function




'***************************************************************
' 재고 명세서
'
'***************************************************************
Public Function GetStockReport(sDate As String, eDate As String _
                            , nChkCustomID As Integer _
                            , sCustomID As String _
                            , sTaxClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_sbStock_StockReport"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sTaxClss)
        Set GetStockReport = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetStockReport"), Err.Description

End Function

'***************************************************************
' 수불명세서
'***************************************************************
Public Function GetSubulReport(ByVal sDate As String, ByVal eDate As String _
                            , ByVal nChkCustomID As Integer, ByVal sCustomID As String _
                            , ByVal nChkArticleID As Integer, ByVal sArticleID As String, ByVal sSubulWidthID As String _
                            , ByVal sOrderFlag As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sReport"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sSubulWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sOrderFlag)
        Set GetSubulReport = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetSubulReport"), Err.Description

End Function

'***************************************************************
' Order별 수불명세서
'***************************************************************
''
''              @nChkOrder         SMALLINT= 0
''            , @sOrderID          VARCHAR(30) = ''
''            , @nChkCustom        SMALLINT = 0
''            , @CustomID          CHAR(4)  = ''
''            , @nChkArticle       SMALLINT = 0
''            , @ArticleID         CHAR(4)  = ''
''            , @TaxClss           CHAR(1)  = '9'      -- '0' :A /  '1' : b / '9': 전체

Public Function GetSubulOrder(ByVal nChkDate As Integer, ByVal sDate As String, ByVal eDate As String _
                            , ByVal nChkOrde As Integer, ByVal sOrderID As String _
                            , ByVal nChkCustomID As Integer, ByVal sCustomID As String _
                            , ByVal nChkArticleID As Integer, ByVal sArticleID As String _
                            , ByVal sOrderFlag As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sOrderSubul"
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrde)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sOrderFlag)
        Set GetSubulOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetSubulOrder"), Err.Description

End Function


'-----------------------------------------------------------------------
' 계상완료처리 및 수주 마감
' 유상수
Public Function UpdateComplete(TItem() As TMonthCustom, Index As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_uProcCostComplete"

        .Prepared = True

        For i = 0 To UBound(TItem) - 1
            Call ClearParameter(adoCmd)
        
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TItem(i).sYearMon)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem(i).sCustomID)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 2, Index)
            
            .Execute
        Next i
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateComplete = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    UpdateComplete = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "계상완료 처리중 에러"
End Function






Public Function GetSubul(sDate As String, _
            nChkCustomID As Integer, sCustomID As String, _
            nChkArticleID As Integer, sArticleID As String, _
            nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sSubul"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)

        Set GetSubul = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetSubul"), Err.Description
End Function

Public Function UpdateSubulSet(TItem() As TSubul) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, j%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_Subul_uSubulSet"
        .Prepared = True

        For i = 0 To UBound(TItem)
            Call ClearParameter(adoCmd)
                
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem(i).SubulDate)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem(i).SubulClss)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem(i).IOClss)
            .Parameters.Append .CreateParameter(, adInteger, adParamInputOutput, 2, TItem(i).Seq)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, TItem(i).OrderNo)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TItem(i).OrderSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem(i).CustomID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem(i).ArticleID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem(i).QtyUnit)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, TItem(i).Custom)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem(i).WorkID)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem(i).Cnt)
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem(i).Qty)
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem(i).Price)
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, TItem(i).OutRealQty)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem(i).OutSeq)
            
            .Execute
            
            TItem(i).Seq = TItem(0).Seq
        Next i
    
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CSubul.UpdateSubulSet"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(UBound(TItem))
    For i = 0 To UBound(TItem)
        With TItem(i)
            sLog(i) = "xp_Subul_uSubulSet '" & .SubulDate & "', '" & .SubulClss & "', '" & .IOClss & "', " & .Seq & ", '" & _
                .OrderNo & "', " & .OrderSeq & ", '" & .CustomID & "', '" & .ArticleID & "', '" & .QtyUnit & "', '" & _
                .Custom & "', '" & .WorkID & "', " & .Cnt & ", " & .Qty & ", " & .Price & ", " & .OutRealQty & ", '" & _
                .OrderID & "', " & .OutSeq & " "
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateSubulSet = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateSubulSet"), Err.Description
    Else
        Call LogService(sLog)

        UpdateSubulSet = True
    End If
End Function


Public Function GetTax(sTaxSeq As String, sPrnDate As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Subul_uTaxPrintDate"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sTaxSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sPrnDate)
        
        .Execute
        
        .CommandText = "xp_Subul_sTax"
        
        Call ClearParameter(adoCmd)
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sTaxSeq)
        
        Set GetTax = .Execute
    End With
    
    Set adoCmd = Nothing
    
End Function

Public Function GetTaxTotal(sTaxSeq As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Subul_sTaxTotal"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sTaxSeq)
        
        .Execute
        
        Set GetTaxTotal = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetTaxSeqMax(sBasisYearMon As String) As String
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Subul_sTaxMax"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sBasisYearMon)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamOutput, 8, 0)
        
        .Execute
        
        GetTaxSeqMax = .Parameters(1).Value
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetTaxSeq(sBasisYearMon As String, sTaxSeq As String) As Boolean
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Subul_sTaxSeq"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sBasisYearMon)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sTaxSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamOutput, 8, sTaxSeq)
        
        .Execute
        
        GetTaxSeq = IIf(IsNull(.Parameters(2).Value), False, True)
    End With
    
    Set adoCmd = Nothing
End Function

Public Function ModifyTaxSeq(TItem As TProcCostDet) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_uTaxSeq"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TItem.sBasisYearMon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nProcSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sTaxSeq)
        
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    ModifyTaxSeq = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    ModifyTaxSeq = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "입력처리중 에러"
End Function

Public Function DeleteProcCostDet(TItem As TProcCostDet) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_dProcCostDet"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, TItem.sBasisYearMon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nProcSeq)
        
        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    DeleteProcCostDet = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    DeleteProcCostDet = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "입력처리중 에러"
End Function

Public Function GetSetAccountByArticle(sDate As String, nChkCustomID As Integer, sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sSetAccountByArticle"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        
        Set GetSetAccountByArticle = .Execute
    End With
    
    Set adoCmd = Nothing
    
End Function

Public Function GetOrderOne(sOrderID As String, sSDate As String, sEDate As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Subul_sOrderOne"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        
        Set GetOrderOne = .Execute
    End With
    
    Set adoCmd = Nothing
    
End Function

Public Function GetTaxList(nChkDate As Integer, sSDate As String, sEDate As String, nChkCustom As String, sCustomID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Subul_sTaxList"
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 6, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 6, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        
        Set GetTaxList = .Execute
    End With
    
    Set adoCmd = Nothing
    
End Function




