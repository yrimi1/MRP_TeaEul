VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' 변경이력
'-------------------------------------------------------------------------------------
' 요청ID: S_201301_삼우_03
' 요청자 : 노과장
' 요청일자: 2013.01.22
' 요청내용 : 검사실적 조회에서 오더전체 자료 나오지 않고 검색조건에 입력된 날짜에 검색된 자료만 나오게 요청
'**********************************************************************************

Option Explicit

Type TInspect
    OrderID     As String  ' 관리 번호
    RollSeq     As Integer ' 일련 순위
    OrderSeq    As Integer  ' 색상 번호
    RollNo      As Integer ' Roll No
    ExamNO      As String  ' 검사 호기
    ExamDate    As String  ' 검사 일자
    ExamTime    As String  ' 검사 시간
    TeamID      As String  ' 검사 조
    PersonID    As String  ' 검사자 코드
    StuffQty    As Integer ' 투입원단 수량
    RealQty     As Single  ' 실제검사 수량
    CtrlQty     As Integer ' 조정검사 수량
    SampleQty   As Single  ' 견본 수량
    LossQty     As Single  ' 보상 수량
    CutQty      As Single  ' 난단 수량
    UnitClss     As String   ' 수량 단위
    StuffWeight      As Single  ' 원단 중량
    StuffWeightUnit   As Integer ' 단위당 중량
    StuffWidth       As Single  ' 원단 폭
    Density     As Integer ' 원단 밀도
    GradeID     As String  ' 등급
    LotNo       As String  ' Lot No
    DefectQty   As Integer ' 불량 갯수
    DefectPoint As Integer ' 불량 점수
    DefectID    As String  ' 대표불량 코드
    DefectClss      As String  ' 대표불량 불량 구분
    CutDefectID As String  ' 난단 대표불량 코드
    CutDefectClss   As String  ' 난단 대표불량 불량 구분
End Type

Type TInspectSub
    OrderID   As String    ' 관리 번호
    RollSeq   As Integer   ' 일련 순위
    DefectSeq As Integer   ' 불량 순위
    DefectID  As String    ' 불량 코드
    YPos      As Integer   ' 수직위치 * 100
    Demerit   As Integer   ' 감점
End Type

Private Const CLASS_NAME As String = "CInspect"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

'------------------------------------------
Public Function GetOrder(nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkCustomID As Integer, sCustomID As String, _
    nChkArticleID As Integer, sArticleID As String, _
    nChkOrder As Integer, sOrder As String, nChkClose As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sOrder"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkClose)

        Set GetOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOrder"), Err.Description
End Function

'-----------------------------
'S_201301_삼우_03  에 의한 수정 - 날짜 조건 추가
Public Function GetOrderSub(nChkDate As Integer, _
                            sSDate As String, _
                            sEDate As String, _
                            sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sOrderSub"
        
        'S_201301_삼우_03  에 의한 추가------------------------------------------------------
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        '-----------------------------------------------------------------------------------
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)

        Set GetOrderSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOrderSub"), Err.Description
End Function

'-----------------------------
'S_201301_삼우_03  에 의한 수정 - 날짜 조건 추가
Public Function GetInspect(nChkDate As Integer, _
                            sSDate As String, _
                            sEDate As String, _
                            sOrderID As String, _
                            nOrderSeq As Integer, _
                            Optional nChkRollSeq As Integer = 0, _
                            Optional nRollSeq As Integer = 0) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspect"

        'S_201301_삼우_03  에 의한 추가------------------------------------------------------
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        '-----------------------------------------------------------------------------------
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkRollSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nRollSeq)

        Set GetInspect = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspect"), Err.Description
End Function
'-------------------------------
Public Function GetInspectSub(sOrderID As String, nRollID As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspectSub"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nRollID)

        Set GetInspectSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectSub"), Err.Description
End Function

'-------------------------------
Public Function GetInspectRollNo(sOrderID As String, sColorID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkQCLot As Integer, sQCLot As String, nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspectRollNo"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sColorID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkQCLot)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sQCLot)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set GetInspectRollNo = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectRollNo"), Err.Description
End Function

'------------------------------------------
Public Function GetDefect(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sDefect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)

        Set GetDefect = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefect"), Err.Description
End Function
'---------------------------------------------------------
Public Function GetDefectTotal(nChkDate As Integer, sSDate As String, sEDate As String, _
                        nChkCustomID As Integer, sCustomID As String, _
                        nChkArticleID As Integer, sArticleID As String, _
                        nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sDefectTotal"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrder)

        Set GetDefectTotal = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefectTotal"), Err.Description
End Function
'----------------------------
Public Function GetResultByDate(nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkOrder As Integer, sOrder As String, nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sResultByDate"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, sExamNo)

        Set GetResultByDate = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetResultByDate"), Err.Description
End Function

Public Function GetResultByLot(nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkOrder As Integer, sOrder As String, nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sResultByLot"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, sExamNo)

        Set GetResultByLot = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetResultByLot"), Err.Description
End Function


'-------------------------------
Public Function GetResultByOrder(nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkOrder As Integer, sOrder As String, nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sResultByOrder"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, sExamNo)
        
        Set GetResultByOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetResultByDate"), Err.Description
End Function

'------------------------------
Public Function GetGradeQtyByOrder(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sGradeQtyByOrder"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)
        
        Set GetGradeQtyByOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetGradeQtyByOrder"), Err.Description
End Function

'------------------------------
Public Function GetGradeQtyByLot(sOrderID As String, sOrderSeq As Integer, sLotNo As String, _
                                 nChkDate As Integer, sSDate As String, sEDate As String, _
                                 nChkExamNo As Integer, sExamNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sGradeQtyByLot"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, sOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
'        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkLotNo)
'        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)
        
        Set GetGradeQtyByLot = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetGradeQtyByOrder"), Err.Description
End Function

'---------------------------
Public Function GetGradeQtyByColor(sOrderID As String, nOrderSeq As Integer, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sGradeQtyByColor"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set GetGradeQtyByColor = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetGradeQtyByColor"), Err.Description
End Function
'----------------------------------
Public Function PrintInspect(bByLot As Boolean, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkCustomID As Integer, sCustomID As String, nChkArticleID As Integer, sArticleID As String, _
    nChkOrder As Integer, sOrder As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = IIf(bByLot, "xp_Inspect_pInspectByLot", "xp_Inspect_pInspect")

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrder)

        Set PrintInspect = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintInspect"), Err.Description
End Function
'-----------------------------------
Public Function PrintResultByOrder(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_pResultByOrder"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set PrintResultByOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintResultByOrder"), Err.Description
End Function
'오더별 검사결과표 엑셀 인쇄양식
Public Function PrintResultByOrderExcel(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_pResultByOrder_Excel"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set PrintResultByOrderExcel = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintResultByOrderExcel"), Err.Description
End Function


'LOT별 검사결과표 엑셀 인쇄양식
Public Function PrintResultByLotExcel(sOrderID As String, sOrderSeq As Integer, sLotNo As String, _
    nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_pResultByLot_Excel"
                        

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, sOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
'        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkLotNo)
'        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set PrintResultByLotExcel = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintResultByOrderExcel"), Err.Description
End Function
Public Function PrintResultByOrder_eng(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_pResultByOrder_ENG"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set PrintResultByOrder_eng = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintResultByOrder"), Err.Description
End Function
'-------------------------------
Public Function PrintResultByColor(sOrderID As String, nOrderSeq As Integer, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_pResultByColor"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set PrintResultByColor = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintResultByColor"), Err.Description
End Function
Public Function PrintResultByColor_eng(sOrderID As String, nOrderSeq As Integer, nChkDate As Integer, sSDate As String, sEDate As String, _
    nChkLotNo As Integer, sLotNo As String, nChkExamNo As Integer, sExamNo As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_pResultByColor_eng"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)

        Set PrintResultByColor_eng = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintResultByColor"), Err.Description
End Function


'---------------------------------------------------
Public Function GetExistRollNo(nRollNoClss As Integer, sOrderID As String, nOrderSeq As String, nRollNo As Integer, sExamNo As String, sLotNo As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim nRoll%
    
    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sExistRollNo"
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nRollNoClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nRollNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 5, sLotNo)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sExamNo)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamOutput, 2, nRoll)
        
        .Execute
        
        GetExistRollNo = IIf(IsNull(.Parameters(6)), False, True)
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetExistRollNo"), Err.Description
End Function

Public Function GetDefectByLang(nChkLang As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sDefectByLang"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkLang)

        Set GetDefectByLang = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefectByLang"), Err.Description
End Function



Public Function AddNewInspect(tIns As TInspect, nInsSub As Integer, tInsSub() As TInspectSub) As Boolean
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
        .CommandText = "xp_Inspect_iInspect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tIns.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, tIns.RollSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.OrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInputOutput, 2, tIns.RollNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, tIns.ExamNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tIns.ExamDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tIns.ExamTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, tIns.TeamID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tIns.PersonID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.StuffQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.RealQty)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.CtrlQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.SampleQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.LossQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.CutQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.UnitClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.StuffWeight)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 5, tIns.StuffWeightUnit)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.Density)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.StuffWidth)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.GradeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, tIns.LotNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tIns.DefectID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.DefectClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tIns.CutDefectID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.CutDefectClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.DefectQty)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.DefectPoint)
         
        .Execute
        iSQL = iSQL + 1

        tIns.RollSeq = .Parameters(1).Value
        tIns.RollNo = .Parameters(3).Value

        ' Insert InspectSub
        .CommandText = "xp_Inspect_iInspectSub"
        .Prepared = True

        For i = 0 To nInsSub
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tIns.OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.RollSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).DefectSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tInsSub(i).DefectID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).YPos)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).Demerit)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CInspect.AddNewInspect"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(nInsSub + 1)
    With tIns
        sLog(0) = "xp_Inspect_iInspect '" & .OrderID & "', " & .RollSeq & ", '" & .OrderSeq & "', " & .RollNo & ", '" & _
            .ExamNO & "', '" & .ExamDate & "', '" & .ExamTime & "', '" & .TeamID & "', '" & .PersonID & "', " & _
            .StuffQty & ", " & .RealQty & ", " & .CtrlQty & ", " & .SampleQty & ", " & .LossQty & ", " & _
            .CutQty & ", '" & .UnitClss & "', " & .StuffWeight & ", " & .StuffWeightUnit & ", " & .Density & ", " & .StuffWidth & ", '" & _
            .GradeID & "', '" & .LotNo & "', '" & .DefectID & "', '" & .DefectClss & "', '" & _
            .CutDefectID & "', '" & .CutDefectClss & "', " & .DefectQty & ", " & .DefectPoint & " "

    End With
    For i = 0 To nInsSub
        With tInsSub(i)
            sLog(i + 1) = "xp_Inspect_iInspectSub '" & .OrderID & "', " & .RollSeq & ", " & .DefectSeq & ", '" & _
                .DefectID & "', " & .YPos & ", " & .Demerit
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewInspect = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewInspect"), Err.Description
    Else
        Call LogService(sLog)

        AddNewInspect = True
    End If
End Function

Public Function UpdateInspect(tIns As TInspect, nInsSub As Integer, tInsSub() As TInspectSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Update Inspect
        .CommandText = "xp_Inspect_uInspect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tIns.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.RollSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.OrderSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, tIns.ExamNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tIns.ExamDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tIns.ExamTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, tIns.TeamID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tIns.PersonID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.StuffQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.RealQty)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.CtrlQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.SampleQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.LossQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.CutQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.UnitClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.StuffWeight)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 5, tIns.StuffWeightUnit)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.Density)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tIns.StuffWidth)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.GradeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, tIns.LotNo)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tIns.DefectID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.DefectClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tIns.CutDefectID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tIns.CutDefectClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.DefectQty)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tIns.DefectPoint)
        
        .Execute
        iSQL = iSQL + 1

        ' Insert InspectSub
        .CommandText = "xp_Inspect_iInspectSub"
        .Prepared = True

        For i = 0 To nInsSub
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tInsSub(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).RollSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).DefectSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tInsSub(i).DefectID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).YPos)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tInsSub(i).Demerit)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CInspect.UpdateInspect"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(nInsSub + 1)
    With tIns
        sLog(0) = "xp_Inspect_iInspect '" & .OrderID & "', " & .RollSeq & ", '" & .OrderSeq & "', " & .RollNo & ", '" & _
            .ExamNO & "', '" & .ExamDate & "', '" & .ExamTime & "', '" & .TeamID & "', '" & .PersonID & "', " & _
            .StuffQty & ", " & .RealQty & ", " & .CtrlQty & ", " & .SampleQty & ", " & .LossQty & ", " & _
            .CutQty & ", '" & .UnitClss & "', " & .StuffWeight & ", " & .StuffWeightUnit & ", " & .Density & ", " & .StuffWidth & ", '" & _
            .GradeID & "', '" & .LotNo & "', '" & .DefectID & "', '" & .DefectClss & "', '" & _
            .CutDefectID & "', '" & .CutDefectClss & "', " & .DefectQty & ", " & .DefectPoint & " "
    End With
    For i = 0 To nInsSub
        With tInsSub(i)
            sLog(i + 1) = "xp_Inspect_iInspectSub '" & .OrderID & "', " & .RollSeq & ", " & .DefectSeq & ", '" & _
                .DefectID & "', " & .YPos & ", " & .Demerit
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateInspect = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateInspect"), Err.Description
    Else
        Call LogService(sLog)

        UpdateInspect = True
    End If
End Function

Public Function DeleteInspect(sOrderID As String, nRollSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_dInspect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nRollSeq)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CInspect.DeleteInspect"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Inspect_dInspect '" & sOrderID & "', " & CStr(nRollSeq)

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteInspect = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteInspect"), Err.Description
    Else
        Call LogService(sLog)

        DeleteInspect = True
    End If
End Function

'-------------------------------
Public Function GetInspectTotal(nChkDate As Integer, sSDate As String, sEDate As String, nChkCustomID As Integer, sCustomID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspectTotal"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)

        Set GetInspectTotal = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectTotal"), Err.Description
End Function

'-------------------------------
Public Function GetInspectTotalMonth(sSDate As String, sEDate As String, nChkCustomID As Integer, sCustomID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspectTotal_Month"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        
        Set GetInspectTotalMonth = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectTotalMonth"), Err.Description
End Function

'-------------------------------
Public Function PrintInspectRollDetail(sOrderID As String, sSDate As String, sEDate As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_mInspectRollDetail"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sEDate)

        .Execute
        
        Call ClearParameter(adoCmd)
        
        .CommandText = "xp_Inspect_sInspectRollDetail"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        
        Set PrintInspectRollDetail = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "PrintInspectRollDetail"), Err.Description
End Function

Public Function GetDefectArray(sSDate As String, sEDate As String, sOrderID As String, nOrderSeq As Integer, sLotNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sDefectArray"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, sLotNo)
        
        Set GetDefectArray = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefectArray"), Err.Description
End Function

Public Function GetInspectByLot(sSDate As String, sEDate As String, nClss As Integer, _
                                nChkCustom As Integer, sCustomID As String, _
                                nChkArticle As Integer, sArticleID As String, _
                                nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspectByLot"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nClss)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sOrder)
        ' nClss     0: 전체, 1:제직불량만, 2:가공불량만
        Set GetInspectByLot = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectByLot"), Err.Description
End Function
Public Function GetInspectByLotPerMonth(sSDate As String, sEDate As String, nClss As Integer, _
                                nChkCustom As Integer, sCustomID As String, _
                                nChkArticle As Integer, sArticleID As String, _
                                nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Inspect_sInspectByLotPerMonth"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nClss)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sOrder)
        ' nClss     0: 전체, 1:제직불량만, 2:가공불량만
        
        Set GetInspectByLotPerMonth = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectByLotPerMonth"), Err.Description
End Function

