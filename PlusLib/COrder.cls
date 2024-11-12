VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************
' 변경이력
' 요청 ID : S_201211_삼우_01
' 요청내용 : 삼우 신규 태그 추가
' 변경일자 : 2012.11.12
' 변경및처리내용 : TagArticle 16에서 22로 변경
'*******************************************************************************
Option Explicit

Public Type TOrder
    sOrderID    As String   '관리번호
    sCustomID   As String   '거래처코드
    sOrderNo    As String   '오더번호
    sPoNO       As String   'PONo
    sChemClss   As String   '원단구분
    sOrderForm  As String   '주문형태
    sOrderClss  As String   '주문구분
    sAcptDate   As String   '접수일자
    sDvlyDate   As String   '납기일자
    sArticleID  As String   '품명코드
    sDvlyPlace  As String   '납품장소
    sWorkID     As String   '가공구분
    sPriceClss  As String   '화폐구분
    nExchRate   As Single   '환율
    nOrderQty   As Long   '수주수량
    sUnitClss   As String   '수량단위
    nColorCnt   As Integer  '색상수
    sStuffWidth As String   '생지폭
    nStuffWeight    As Integer  '생지중량
    nCutQty As Integer  '필장
    sWorkWidth  As String   '가공폭
    nWorkWeight As Integer  '가공중량
    nWorkDensity    As Integer  '가공밀도
    nChunkRate  As Single   '축율
    nLossRate   As Single   'Loss율
    nReduceRate As Single   '감량율
    sTagClss    As String   'Tag 구분
    sLabelID    As String   'Labe 구분
    sBandID     As String   'Band 구분
    sEndClss    As String   'End Mark 구분
    sMadeClss   As String   'Made IN Korea 구분
    sSurfaceClss    As String   '사용면구분
    sShipClss   As String   'Ship Sample
    sAdvnClss   As String   'Advnced Sample
    sLotClss    As String   'Lot Sample
    sEndMark    As String   'End Mark 내용
    sTagArticle As String   'Tag 품명
    sTagArticle2 As String  'Tag 품명2  'S_201905_삼우_01 에 의한 추가
    sTagOrderNo As String   'Tag 오더번호
    sTagOrderNo2 As String   'Tag 오더번호2 'S_201905_삼우_01 에 의한 추가
    sTagRemark  As String   'Tag 비고사항
    sTagRemark2 As String   'Tag 비고사항2
    sTag        As String   'Tag 내용
    sBasisID    As String   '검사기준
    sBasisUnit  As String   '검사기준단위
    sSpendingClss   As String   '소요량 정산 구분
    sDyeingID   As String   '염색기구분
    sWorkingClss    As String   '가공료 정산 구분
    sPatternID  As String   '공정패턴
    sAccountClss    As String   '세금계산서 발행 구분
    sModifyClss As String   '정정구분
    sModifyRemark   As String   '정정사유
    sCancelRemark   As String   '취소사유
    sRemark     As String   '비고사항
    sActiveClss As String   '확정구분
    sCloseClss  As String   '종결구분
    sModifyDate As String   '정정일자
    sBTID   As String       'BT ID
    nBTIDSeq       As Integer
    sOrderFlag     As String    '오더구분
    sItem          As String    'ITEM
    sSubulWidthID  As String    '재고폭
    OutTelNO       As String    '출고처 전화번호
    
    '---------------------------------------------------
    'S_201901_삼우_01 에 의한 추가 : 자라Tag 내용 추가
    sTagQuality     As String   'Quality
    sTagDestition   As String   'Destition
    sUnitPriceClss  As String   '단가 단위
    nWeightPerYard  As Single   '야드당 KG
    '---------------------------------------------------
    
End Type

Public Type TOrderSub
    sOrderID        As String   '관리번호
    nOrderSeq       As Integer  '색상순번
    sColorID        As String   '색상코드
    sColor          As String   '색상명
    sDesignNO       As String   '디자인 번호
    nColorQty       As Long     '수주수량
    nUnitPrice      As Single   '단가
    sPoNO           As String   'P/o NO
    sFlag           As String   '추가 수정 삭제 구분
End Type


Public Type TOrderInst
    sInstDate       As String
    nInstSeq        As Integer
    sOrderID        As String
    
End Type

Public Type TOrderInstDet
    sInstDate       As String
    nInstSeq        As Integer
    nProcSeq        As Integer
    sProcessID      As String
End Type

'---------------------------------------------------------------------------------------------
Private Const CLASS_NAME As String = "COrder"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetExistOrder(sOrderNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sExistOrderNo"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 35, sOrderNo)

        Set GetExistOrder = .Execute
    End With

    Set adoCmd = Nothing

End Function

Public Function AddNewOrder(TItem As TOrder, nOrderSub As Integer, tItemSub() As TOrderSub, TInst As TOrderInst, TInstDET() As TOrderInstDet) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        
        .CommandText = "xp_Order_iOrder"

        .Parameters.Append .CreateParameter(, adChar, adParamInputOutput, 10, TItem.sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, TItem.sOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, TItem.sPoNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderForm)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sAcptDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sDvlyDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, TItem.sDvlyPlace)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPriceClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nExchRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nOrderQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sUnitClss)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 4, TItem.nColorCnt)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sStuffWidth)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nStuffWeight)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.nCutQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sWorkWidth)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nWorkWeight)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nWorkDensity)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nChunkRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nLossRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nReduceRate)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sTagClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sLabelID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sBandID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sEndClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sMadeClss)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sSurfaceClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 20, TItem.sShipClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 20, TItem.sAdvnClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 20, TItem.sLotClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 250, TItem.sEndMark)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagArticle)
        'S_201905_삼우_01 에 의한 추가
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagArticle2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagOrderNo)
        'S_201905_삼우_01 에 의한 추가
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagOrderNo2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagRemark)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagRemark2)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 250, TItem.sTag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sBasisID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sBasisUnit)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sSpendingClss)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sDyeingID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sWorkingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sAccountClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 250, TItem.sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sActiveClss)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nBTIDSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sChemClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sPatternID)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sItem)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sSubulWidthID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.OutTelNO)
        '--------------------------------------------------------------------------------------------------------------
        'S_201901_삼우_01 에 의한 추가 : 자라Tag 내용 추가
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, TItem.sTagQuality)      'Tag Quality
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, TItem.sTagDestition)    'Tag Destition
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sUnitPriceClss)        '단가단위
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nWeightPerYard)      '야드당 g 수
        '--------------------------------------------------------------------------------------------------------------
        
        .Execute
        iSQL = iSQL + 1

        TItem.sOrderID = .Parameters(0).Value

        .CommandText = "xp_Order_iOrderSub"
        .Prepared = True

        For i = 0 To nOrderSub
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sOrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItemSub(i).nOrderSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 5, tItemSub(i).sColorID)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tItemSub(i).sColor)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tItemSub(i).sDesignNO)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nColorQty)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nUnitPrice)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, tItemSub(i).sPoNO)
            
            .Execute
            iSQL = iSQL + 1
        Next i
        
        Dim nInstSeq As Integer
        If UBound(TInstDET) > 0 Then
            ' 생지 투입계획
            .CommandText = "xp_Order_iOrderInst"
            .Prepared = True
            Call ClearParameter(adoCmd)
            .Parameters.Append .CreateParameter(, adChar, adParamInputOutput, 8, TInst.sInstDate)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, TInst.nInstSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TInst.sOrderID)
            .Execute
            
            TInst.sInstDate = .Parameters(0).Value
            nInstSeq = .Parameters(1).Value
            
            iSQL = iSQL + 1
             
             
            .CommandText = "xp_Order_iOrderInstDet"
            .Prepared = True
    
            For i = 0 To UBound(TInstDET) - 1
                Call ClearParameter(adoCmd)
    
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TInst.sInstDate)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nInstSeq)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TInstDET(i).nProcSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TInstDET(i).sProcessID)
                
                .Execute
                iSQL = iSQL + 1
            Next i
            
        End If
        
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COrder.AddNewOrder"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(nOrderSub + 1)
    With TItem
        sLog(0) = "xp_Order_iOrder '" & .sOrderID & "', '" & .sCustomID & "', '" & .sOrderNo & "', '" & .sPoNO & "', '" & .sOrderForm & "', '" & _
                    .sOrderClss & "', '" & .sAcptDate & "', '" & .sDvlyDate & "', '" & .sArticleID & "', '" & .sDvlyPlace & "', '" & _
                    .sWorkID & "', '" & .sPriceClss & "', " & .nExchRate & ", " & .nOrderQty & ", '" & .sUnitClss & "', " & _
                    .nColorCnt & ", '" & .sStuffWidth & "', " & .nStuffWeight & ", " & .nCutQty & ", '" & .sWorkWidth & "', " & .nWorkWeight & ", " & _
                    .nWorkDensity & ", " & .nChunkRate & ", " & .nLossRate & ", " & .nReduceRate & ", '" & .sTagClss & "', '" & _
                    .sLabelID & "', '" & .sBandID & "', '" & .sEndClss & "', '" & .sMadeClss & "', '" & .sSurfaceClss & "', '" & .sShipClss & "', '" & _
                    .sAdvnClss & "', '" & .sLotClss & "', '" & .sEndMark & "', '" & .sTagArticle & "', '" & .sTagArticle2 & "', '" & .sTagOrderNo & "', '" & _
                    .sTagOrderNo2 & "', '" & .sTagRemark & "', '" & .sTag & "', '" & .sTagRemark2 & "', '" & _
                    .sBasisID & "', '" & .sBasisUnit & "', '" & .sSpendingClss & "', '" & .sDyeingID & "', '" & .sWorkingClss & "', '" & .sAccountClss & "', '" & _
                    .sRemark & "' " & .sActiveClss & "', '" & .sBTID & "', " & .nBTIDSeq & ", '" & .sChemClss & "', '" & .sOrderFlag & "' "
    End With
    For i = 0 To nOrderSub
        With tItemSub(i)
            sLog(i + 1) = "xp_Order_iOrderSub '" & TItem.sOrderID & "', " & .nOrderSeq & ", '" & .sColorID & "', '" & .sColor & "', '" & _
                .sDesignNO & "', " & .nColorQty & ", " & .nUnitPrice & " "
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewOrder = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewOrder"), Err.Description
    Else
        Call LogService(sLog)

        AddNewOrder = True
    End If
End Function

Public Function UpdateOrder(TItem As TOrder, nOrderSub As Integer, tItemSub() As TOrderSub, TInst As TOrderInst, TInstDET() As TOrderInstDet) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%, nInstSeq As Integer

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        
        .CommandText = "xp_Order_uOrder"

        .Parameters.Append .CreateParameter(, adChar, adParamInputOutput, 10, TItem.sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, TItem.sOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, TItem.sPoNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderForm)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sAcptDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sDvlyDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, TItem.sDvlyPlace)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPriceClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nExchRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nOrderQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sUnitClss)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 4, TItem.nColorCnt)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sStuffWidth)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nStuffWeight)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TItem.nCutQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sWorkWidth)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nWorkWeight)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nWorkDensity)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nChunkRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nLossRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nReduceRate)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sTagClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sLabelID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sBandID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sEndClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sMadeClss)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sSurfaceClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 20, TItem.sShipClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 20, TItem.sAdvnClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 20, TItem.sLotClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 250, TItem.sEndMark)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagArticle)
        'S_201905_삼우_01 에 의한 추가
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagArticle2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagOrderNo)
        'S_201905_삼우_01 에 의한 추가
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagOrderNo2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagRemark)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sTagRemark2)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 250, TItem.sTag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sBasisID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sBasisUnit)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sSpendingClss)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sDyeingID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sWorkingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sAccountClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 250, TItem.sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sActiveClss)
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TItem.sBTID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TItem.nBTIDSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sChemClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sOrderFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sPatternID)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sItem)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sSubulWidthID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.OutTelNO)
        '--------------------------------------------------------------------------------------------------------------
        'S_201901_삼우_01 에 의한 추가 : 자라Tag 내용 추가
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, TItem.sTagQuality)      'Tag Quality
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, TItem.sTagDestition)    'Tag Destition
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sUnitPriceClss)        '단가단위
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nWeightPerYard)      '야드당 g 수
        '--------------------------------------------------------------------------------------------------------------
        
        .Execute
        iSQL = iSQL + 1

        For i = 0 To nOrderSub
            If tItemSub(i).sFlag = "A" Then
                .CommandText = "xp_Order_iOrderSub"
                .Prepared = True

                Call ClearParameter(adoCmd)

                .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tItemSub(i).sOrderID)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItemSub(i).nOrderSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 5, tItemSub(i).sColorID)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tItemSub(i).sColor)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tItemSub(i).sDesignNO)
                .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nColorQty)
                .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nUnitPrice)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, tItemSub(i).sPoNO)
                
            ElseIf tItemSub(i).sFlag = "U" Then
                .CommandText = "xp_Order_uOrderSub"
                .Prepared = True

                Call ClearParameter(adoCmd)

                .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tItemSub(i).sOrderID)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItemSub(i).nOrderSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 5, tItemSub(i).sColorID)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tItemSub(i).sColor)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 30, tItemSub(i).sDesignNO)
                .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nColorQty)
                .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tItemSub(i).nUnitPrice)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, tItemSub(i).sPoNO)
            
            ElseIf tItemSub(i).sFlag = "D" Then
                .CommandText = "xp_Order_dOrderSub"
                .Prepared = True

                Call ClearParameter(adoCmd)

                .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tItemSub(i).sOrderID)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tItemSub(i).nOrderSeq)
            End If

            .Execute
            iSQL = iSQL + 1
        Next i
        
        If UBound(TInstDET) > 0 Then
            ' 생지 투입계획
            .CommandText = "xp_Order_iOrderInst"
            .Prepared = True
            Call ClearParameter(adoCmd)
            .Parameters.Append .CreateParameter(, adChar, adParamInputOutput, 8, TInst.sInstDate)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, TInst.nInstSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TInst.sOrderID)
            .Execute
            
            TInst.sInstDate = .Parameters(0).Value
            nInstSeq = .Parameters(1).Value
            iSQL = iSQL + 1
             
             
            .CommandText = "xp_Order_iOrderInstDet"
            .Prepared = True
    
            For i = 0 To UBound(TInstDET) - 1
                Call ClearParameter(adoCmd)
    
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TInst.sInstDate)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nInstSeq)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TInstDET(i).nProcSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TInstDET(i).sProcessID)
                
                .Execute
                iSQL = iSQL + 1
            Next i
            
        End If
        
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COrder.UpdateOrder"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(nOrderSub + 1)
    With TItem
        sLog(0) = "xp_Order_uOrder '" & .sOrderID & "', '" & .sCustomID & "', '" & .sOrderNo & "', '" & .sPoNO & "', '" & .sOrderForm & "', '" & _
                    .sOrderClss & "', '" & .sAcptDate & "', '" & .sDvlyDate & "', '" & .sArticleID & "', '" & .sDvlyPlace & "', '" & _
                    .sWorkID & "', '" & .sPriceClss & "', " & .nExchRate & ", " & .nOrderQty & ", '" & .sUnitClss & "', " & _
                    .nColorCnt & ", '" & .sStuffWidth & "', " & .nStuffWeight & ", " & .nCutQty & ", '" & .sWorkWidth & "', " & .nWorkWeight & ", " & _
                    .nWorkDensity & ", " & .nChunkRate & ", " & .nLossRate & ", " & .nReduceRate & ", '" & .sTagClss & "', '" & _
                    .sLabelID & "', '" & .sBandID & "', '" & .sEndClss & "', '" & .sMadeClss & "', '" & .sSurfaceClss & "', '" & .sShipClss & "', '" & _
                    .sAdvnClss & "', '" & .sLotClss & "', '" & .sEndMark & "', '" & .sTagArticle & "', '" & .sTagArticle2 & "', '" & .sTagOrderNo & "', '" & _
                    .sTagRemark & "', '" & .sTag & "', '" & .sTagRemark2 & "', '" & _
                    .sTagOrderNo2 & "', '" & .sBasisID & "', '" & .sBasisUnit & "', '" & .sSpendingClss & "', '" & .sDyeingID & "', '" & .sWorkingClss & "', '" & .sAccountClss & "', '" & _
                    .sRemark & "' " & .sActiveClss & "', '" & .sBTID & "', " & .nBTIDSeq & ", '" & .sChemClss & "', '" & .sOrderFlag & "' "
    End With
    For i = 0 To nOrderSub
        With tItemSub(i)
            If .sFlag = "A" Then
            sLog(i + 1) = "xp_Order_iOrderSub '" & TItem.sOrderID & "', " & .nOrderSeq & ", '" & .sColorID & "', '" & .sColor & "', '" & _
                .sDesignNO & "', " & .nColorQty & ", " & .nUnitPrice & " "
            ElseIf .sFlag = "U" Then
            sLog(i + 1) = "xp_Order_iOrderSub '" & TItem.sOrderID & "', " & .nOrderSeq & ", '" & .sColorID & "', '" & .sColor & "', '" & _
                .sDesignNO & "', " & .nColorQty & ", " & .nUnitPrice & " "
            ElseIf .sFlag = "D" Then
                sLog(i + 1) = "xp_Order_dOrderSub '" & TItem.sOrderID & "', " & .nOrderSeq & " "
            End If
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateOrder = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateOrder"), Err.Description
    Else
        Call LogService(sLog)

        UpdateOrder = True
    End If
End Function

Public Function GetCustomPrice(TItem As PlusLib2.TCustomUnit) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '거래처 단가 조회
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_sUnitPrice"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sStuffWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        
        Set GetCustomPrice = .Execute
        
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
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        Err.Raise Err.Number, "CCustom.DeleteCustomUnit", Err.Description
    Else   ' 로그
        Call LogService(sLog)

    End If

End Function



Public Function DeleteOrder(sOrderID As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_Order_dOrder"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)

        .Execute

    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COrder.DeleteOrder"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(0)
    sLog(0) = "xp_Order_dOrder '" & sOrderID & "' "

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        DeleteOrder = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewOrder"), Err.Description
    Else
        Call LogService(sLog)

        DeleteOrder = True
    End If
End Function

Public Function GetDraftOrder(nChkDate As Integer, sFromDate As String, sToDate As String, _
                            nChkCustom As Integer, sCustomID As String, _
                            nChkArticle As Integer, sArticleID As String, _
                            nChkOrder As Integer, sOrder As String, _
                            nChkChemClss As Integer, sChemClss As String, _
                            nChkOrderClss As Integer, sChkPattern As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sDraftOrder"
       
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sFromDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sToDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkChemClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sChemClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrderClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sChkPattern)
        
        Set GetDraftOrder = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetOrderOne(sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sOrderOne"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)

        Set GetOrderOne = .Execute
    End With

    Set adoCmd = Nothing

End Function

Public Function GetOrderSub(sOrderID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sOrderSub"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)

        Set GetOrderSub = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetCustomData(sCustomID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sCustomData"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)

        Set GetCustomData = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetArticleData(sArticleID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sArticleData"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)

        Set GetArticleData = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetUnitPrice(sCustomID As String, sArticleID As String, sColorID As String, sWorkID As String, _
                            sPriceClss As String, sUnitClss As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sUnitPrice"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 5, sColorID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sWorkID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sPriceClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sUnitClss)

        Set GetUnitPrice = .Execute

    End With

    Set adoCmd = Nothing
End Function

Public Function GetOrderTotal(nChkDate As Integer, sSDate As String, sEDate As String, _
                                nChkCustom As Integer, sCustomID As String, nChkArticle As Integer, sArticleID As String, _
                                nChkOrder As Integer, sOrder As String, _
                                Optional nChkClose As Integer = 0, Optional nChkOrderClss As Integer = 0, _
                                Optional nChkFiber As Integer = 0) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sOrderTotal"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticle)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkClose)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrderClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkFiber)
        
        Set GetOrderTotal = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetOrderSubTotal(sOrderID As String, Optional nChkOrderSeq As Integer = 0, Optional nOrderSeq As Integer = 0) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sOrderSubTotal"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)

        Set GetOrderSubTotal = .Execute
    End With

    Set adoCmd = Nothing
End Function


Public Function PrintOrderDetail(nChkDate As Integer, sSDate As String, sEDate As String, _
                            nChkCustom As Integer, sCustomID As String, _
                            nChkArticle As Integer, sArticleID As String, _
                            nChkOrder As Integer, sOrder As String, _
                            nChkChemClss As Integer, sChemClss As String, _
                            nChkOrderClss As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_pOrderDetail"
        
       
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkChemClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sChemClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrderClss)
        
        Set PrintOrderDetail = .Execute
    End With

    Set adoCmd = Nothing
End Function


Public Function PrintOrderClose(nChkDate As Integer, sSDate As String, sEDate As String, _
                                nChkCustom As Integer, sCustomID As String, nChkArticle As Integer, sArticleID As String, _
                                nChkOrder As Integer, sOrder As String, _
                                Optional nChkClose As Integer = 0, Optional nChkOrderClss As Integer = 0, _
                                Optional nChkFiber As Integer = 0) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_pOrderClose"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticle)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkClose)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrderClss)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkFiber)
        
        Set PrintOrderClose = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function UpdateOrderClose(sOrderID() As String, NewValue As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%, iSQL%, sDate$
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    sDate = Format(Now, "YYYYMMDD")
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_uOrderClose"
        .Prepared = True


        For i = 0 To UBound(sOrderID)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, , NewValue)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID(i))

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
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COrder.UpdateOrderClose"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(sOrderID))

    For i = 0 To UBound(sOrderID)
        sLog(i) = "xp_Order_uOrderclose " & NewValue & ", '" & sOrderID(i) & "' "
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateOrderClose = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateOrderClose"), Err.Description
    Else
        Call LogService(sLog)

        UpdateOrderClose = True
    End If

End Function

Public Function GetResultDayByProcess(ByVal sSDate As String, ByVal sEDate As String, ByVal nExchange As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sResultDayByProcess"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nExchange)

        Set GetResultDayByProcess = .Execute
    End With

    Set adoCmd = Nothing
End Function

'일일예상매출표
Public Function GetResultSaleExpect(ByVal sDate As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Result_SaleExpect"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)

        Set GetResultSaleExpect = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetResultSaleExpectDetail(ByVal sDate As String, ByVal sCustomID As String _
                                        , ByVal sArticleID As String, ByVal sWorkID As String, ByVal nUnitPrice As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Result_SaleExpect_Detail"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sWorkID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nUnitPrice)

        Set GetResultSaleExpectDetail = .Execute
    End With

    Set adoCmd = Nothing

End Function

'--- 일자별 Order접수 명세서
Public Function GetOrderAcptView(ByVal sDate As String _
                        , ByVal eDate As String _
                        , ByVal sFromDate As String _
                        , ByVal sToDate As String _
                        , ByVal nChkOrder As Integer, ByVal sOrderID As String _
                        , ByVal nChkCustom As Integer, ByVal sCustomID As String _
                        , ByVal nChkArticle As Integer, ByVal sArticleID As String _
                        , ByVal nExchRate As Single, ByVal nOrderFlag As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sAcptView"
       
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sFromDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sToDate)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sOrderID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, nExchRate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderFlag)
        
        Set GetOrderAcptView = .Execute
    End With

    Set adoCmd = Nothing
End Function

