Attribute VB_Name = "chs_Module"
'********************************************************************************************
'변경이력
'-------------------------------------------------------------------------------------------
' 요청 ID : S_201303_태을염직_01
' 요청자 : 김대진 대리
' 요청내용 : 수량이 10만이상시 오류
' 변경일자 : 2013.03.19
' 변경내용 : integer에서 long으로 변경
'
'변경이력
'
'2013.12.12   자체    오승욱   S_201312_태을염직_99   지번주소에서 도로명 주소로 입력가능하게,거래처 주소 도로명 주소 Select
'**************************************************************************************************
'
Option Explicit

Public Const PLANCPB = "4000"
Public Const PLANRAPID = "4300"
Public Const AllStr = "(전체)"

' --- 인쇄시 설정
Public Const PRNRowHeight = 400              '인쇄시 RowHeight
Public Const PRNHeaderColor = &HB4B4B4       '&HAAAAAA    '&H9F9F9F    '&H8F8F8F               '인쇄시 Header Title Color
Public Const FrozenColor = &H8000000F

Public g_sysDate As String, g_sysTime As String

Public Enum ESHRINK
    ES_EXPAND = 0
    ES_REDUCE = 1
End Enum

Public Enum eDate
    ED_CUR = 0
    ED_PRE = 1
    ED_NEXT = 2
End Enum


Public Enum CDEPTH
    ED1_DEPTH = &HE9E9E9
    ED2_DEPTH = &HE5E5E5
    ED3_DEPTH = &HE0E0E0
    ED4_DEPTH = &HC0C0C0
End Enum

Public Sub SetStuffWidth(ByVal cboName As Object)
    Dim oCode As PlusLib2.CCode
    Dim rs    As ADODB.Recordset
    Dim II%
    
    On Error GoTo ErrHandler

    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon

    Set rs = oCode.GetStuffWidth
    Set oCode = Nothing
    II = 0
    cboName.Clear
    If Not rs Is Nothing Then
        If Not rs.BOF Then
           rs.MoveFirst
           Do Until rs.EOF
            cboName.AddItem Trim$(rs(0))
            cboName.ItemData(II) = val(rs(1))
            
            Debug.Print II; rs(0); cboName.ItemData(II); rs(1)
            II = II + 1
            rs.MoveNext
           Loop
        End If
    End If

    rs.Close
    Set rs = Nothing
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCode = Nothing

    Err.Raise Err.Number, "chs_Module.SetStuffWidth", Err.Description, Err.HelpFile, Err.HelpContext

End Sub


Public Sub SetPrintMode(ByVal oFlex As VSFlexGrid _
                , ByVal nHeaderRows As Integer _
                , ByVal bMode As Boolean _
                , Optional nOrientation As Integer = 1)
                
    Dim II As Long
    Dim nFromRow As Long, nToRow As Long
    Dim nColWidth As Long
    
    ' 인쇄방향 기본 세로인쇄
    nOrientation = 1
    
    With oFlex
        nFromRow = .FixedRows - nHeaderRows
        nToRow = .FixedRows - 1
        
        .Cell(flexcpBackColor, 0, 1, .Rows - 1, .Cols - 1) = vbWhite
        .Cell(flexcpBackColor, 0, 0, nFromRow, .Cols - 1) = vbWhite
        
        .SheetBorder = vbBlack
        '.GridLineWidth =
        
        If bMode = True Then
            .GridLines = flexGridInset
            
            
            .GridLinesFixed = flexGridFlat
            .GridColorFixed = vbWhite

            .Cell(flexcpFontBold, nFromRow, 0, nToRow, .Cols - 1) = True                  'Header Title Bold처리
            
            .Select nFromRow, 0, nToRow, .Cols - 1
            .CellBorder RGB(0, 0, 0), 1, 1, 1, 1, 1, 1
            
 '          .CellBorder RGB(0, 0, 0), 2, 3, 2, 2, 1, 1
 '           .Cell(flexcpBackColor, nFromRow, 0, nToRow, .Cols - 1) = PRNHeaderColor       'Header Title Backcolor 설정
            
            .RowHeight(0) = 1200
            
            ' Header Title나타낼 부분 살려내기
            For II = .FixedRows To .Rows - 1
                .RowHeight(II) = PRNRowHeight
            Next II
            
            For II = 0 To nFromRow - 1
                .RowHidden(II) = False
            Next II
            
            ' 인쇄폭이 10100 이 넘으면 가로인쇄로 자동 변경
            For II = 0 To .Cols - 1
                If .ColHidden(II) = False Then
                    nColWidth = nColWidth + .ColWidth(II)
                End If
            Next II
            
            If nColWidth > 10100 Then
                nOrientation = 2
            End If
        Else
            .GridLines = flexGridFlat
            .GridLinesFixed = flexGridInset
            .Cell(flexcpFontBold, nFromRow, 0, nToRow, .Cols - 1) = False                  'Header Title Bold처리
            .Cell(flexcpBackColor, nFromRow, 0, nToRow, .Cols - 1) = FrozenColor            'Header Title Backcolor 설정
            For II = .FixedRows To .Rows - 1
                .RowHeight(II) = 225
            Next II
            
            ' Header Title나타낼 부분 Hidden처리
            For II = 0 To nFromRow - 1
                .RowHidden(II) = True
            Next II
        
        End If
        
    End With

End Sub



Public Sub SetCboMonth(ByVal CboBox As Object)
    Dim FromDate As Date, dDate As Date, II As Integer, JJ As Integer
    
    JJ = 24
    
    FromDate = DateAdd("m", JJ * -1, Date)
    CboBox.Clear
    
    For II = 1 To JJ
        dDate = DateAdd("m", II, FromDate)
        CboBox.AddItem Format(dDate, "yyyy년 mm월")
    Next II
    
  '  format( date,  "yyyy년 mm월")
End Sub
'******************************************************************************
' 함   수  명 : GetMonth
' 기       능 : 2004년 03월 -> 200403으로 설정
' 사 용 방 법 : GetMonth( "2004 03" )
' 사 용 인 수 : arg1 -> VsFlexGrid Object
'               arg2 -> subTotal로 설정할 행
'               arg3 -> group 단계
' 리   턴  값 : 없음
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Function GetMonth(ByVal pDate As String) As String
    Dim AAA As Variant
    
    AAA = Split(pDate, " ")
    
    GetMonth = Left(AAA(0), 4) & Left(AAA(1), 2) & "01"

End Function

'******************************************************************************
' 프로시져 명 : GetLastMonthDay
' 기       능 : 현재달의 마지막 일자를 구한다.
' 사 용 방 법 : arg1=getlastmonthday(arg2,arg3)
' 사 용 인 수 : arg1 -> 리턴값 마지막 일자
'               arg2 -> 기준일자
'               arg3 -> 기준일자의 마지막일자 0:기준일자 마지막일자, 1:다음달의 마지막일자
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************

Public Function GetLastDayMonth(ByVal BasDate$, Optional ByVal viOffset As eDate) As String
    Dim MvDate As Date
    Dim iOffSet As Integer
    
    MvDate = CDate(Format$(BasDate, "####-##-##"))
    
'    If Not IsMissing(vdatBase) Then
'        DatBase = CDate(vdatBase)
'    Else
'        DatBase = Date
 '   End If
 '
    Select Case viOffset
        Case ED_CUR:   iOffSet = 0
        Case ED_PRE:   iOffSet = -1
        Case ED_NEXT:  iOffSet = 1
    End Select
    
'    If Not IsMissing(viOffset) Then
'        iOffSet = CInt(viOffset)
'    Else
'        iOffSet = 0
'    End If
    
    GetLastDayMonth = Format$(DateSerial(Year(MvDate), Month(MvDate) + iOffSet + 1, 0), "yyyymmdd")
End Function
Function ALP_TO_STR(TOTAL As Double) As String
    Dim II As Integer, KUM As String, TMP As String * 11
    Dim W_UNIT(11), W_NUM(9), n As String, W_MINUS$
    Dim W_TOTAL As Double, STR_TOTAL  As String
    
    W_TOTAL = TOTAL
    
    If W_TOTAL < 0 Then
         W_TOTAL = W_TOTAL * -1
         W_MINUS = "-"
    Else
         W_MINUS = ""
    End If
    
    KUM = ""
    
    W_UNIT(1) = "百"   '백
    W_UNIT(2) = "拾"   '십
    W_UNIT(3) = "億"   '억
    W_UNIT(4) = "阡"   '천
    W_UNIT(5) = "百"   '백
    W_UNIT(6) = "拾"   '십
    W_UNIT(7) = "萬"   '만
    W_UNIT(8) = "阡"   '천
    W_UNIT(9) = "百"   '백
    W_UNIT(10) = "拾"  '십
    W_UNIT(11) = ""
    
    W_NUM(1) = "壹"  '일
    W_NUM(2) = "貳"  '이
    W_NUM(3) = "蔘"  '태을염직
    W_NUM(4) = "四"  '사
    W_NUM(5) = "五"  '오
    W_NUM(6) = "六"  '육
    W_NUM(7) = "七"  '칠
    W_NUM(8) = "八"  '팔
    W_NUM(9) = "九"  '구
    
    STR_TOTAL = W_TOTAL
    TMP = Space(11 - Len(STR_TOTAL)) & W_TOTAL
    
    For II = 1 To 11
        n = Trim$(Mid$(TMP, II, 1))
        If n <> "0" And n <> "" Then
            KUM$ = KUM$ & W_NUM(val(n)) & W_UNIT(II)
        End If
        If (II = 3 And KUM$ <> "" And Right(KUM$, 1) <> "億") Then
            KUM$ = KUM$ & "億"
        End If
        If (II = 7 And KUM$ <> "" And Right(KUM$, 1) <> "萬" And _
                      Right(KUM$, 1) <> "億") Then
            KUM$ = KUM$ & "萬"
        End If
    Next II
    'ALP_TO_STR = KUM$
    'If W_MINUS < 0 Then
    ALP_TO_STR = Trim$(W_MINUS$ & KUM$)
    'End If
End Function

Public Sub SetGrdColor(ByVal oFlex As VSFlexGrid, ByVal sDepth As String _
                , ByVal nRow1 As Integer, ByVal nCol1 As Integer _
                , ByVal nRow2 As Integer, ByVal nCol2 As Integer)
    Dim nColorVal As Long
    With oFlex
        Select Case sDepth
            Case "1": nColorVal = ED1_DEPTH
            Case "2": nColorVal = ED2_DEPTH
            Case "3": nColorVal = ED3_DEPTH
            Case "4": nColorVal = ED4_DEPTH
            Case Else: nColorVal = vbWhite
        End Select
        .Cell(flexcpBackColor, nRow1, nCol1, nRow2, nCol2) = nColorVal
    End With
End Sub


Public Sub ColResize(ByVal oFlex As VSFlexGrid, ByVal nType As ESHRINK, Optional nPercent As Integer = 10)
    Dim II%
    
    If nType = ES_EXPAND Then
        With oFlex
            For II = .FixedCols To .Cols - 1
                .Redraw = flexRDBuffered
                .ColWidth(II) = Int(.ColWidth(II) * (1 + nPercent / 100))
            Next II
        End With
    Else
        With oFlex
            For II = .FixedCols To .Cols - 1
                .Redraw = flexRDBuffered
                .ColWidth(II) = Int(.ColWidth(II) / (1 + nPercent / 100))
            Next II
        End With
    End If
End Sub

'******************************************************************************
' 함   수  명 : GridCollapse
' 기       능 : 해당행을 Shrink한다
' 사 용 방 법 : DoFlexGridGroup(arg1,arg2,arg3)
' 사 용 인 수 : arg1 -> VsFlexGrid Object
'               arg2 -> subTotal로 설정할 행
'               arg3 -> group 단계
' 리   턴  값 : 없음
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************

Public Sub GridCollapse(ByVal oFlex As VSFlexGrid, ByVal Row As Integer)
    With oFlex
        If Row < .FixedRows Then Exit Sub

        If .IsCollapsed(Row) = flexOutlineCollapsed Then
            .IsCollapsed(Row) = flexOutlineExpanded
        Else
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub

'******************************************************************************
' 함   수  명 : DoFlexGridGroup
' 기       능 : 해당행을 subTotal 행으로 설정한다.
' 사 용 방 법 : DoFlexGridGroup(arg1,arg2,arg3)
' 사 용 인 수 : arg1 -> VsFlexGrid Object
'               arg2 -> subTotal로 설정할 행
'               arg3 -> group 단계
' 리   턴  값 : 없음
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Sub DoFlexGridGroup(ByVal oFlex As VSFlexGrid, ByVal irow As Integer, ByVal iLvl As Integer _
                        , Optional iBackColor As Long = &HE0E0E0, Optional iForeColor As Long = &H80000012)
    With oFlex
        '----  iRow을 subTotal Group으로 설정
        .IsSubtotal(irow) = True
        
        '----  iRow행을 subTotal Group의 level설정
        .RowOutlineLevel(irow) = iLvl

        '----  iRow행을 subTotal Group의 level설정
        .Cell(flexcpBackColor, irow, 0, irow, .Cols - 1) = iBackColor
        .Cell(flexcpForeColor, irow, 0, irow, .Cols - 1) = iForeColor
    End With
End Sub

'******************************************************************************
' 함   수  명 : SetGrdShrink
' 기       능 : VsFlexGrid의 그룹의 +/- 처리
' 사 용 방 법 : SetGrdShrink(arg1,arg2)
' 사 용 인 수 : arg1 -> VsFlexGrid Object
'               arg2 -> EORDERMAKE 구조체 값
' 리   턴  값 : 없음
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Sub SetGrdShrink(ByVal oFlex As VSFlexGrid, nType As EORDERMAKE)
    Dim II As Integer
    Dim nRows As String, sRows_var As Variant
    
''    OM_EXPAND = 0   ''확장
''    OM_REDUCE = 1   ''축소
    
    nRows = ""
    With oFlex
        If .Rows < .FixedRows Then Exit Sub
        Select Case nType
            Case 0
                For II = .FixedRows To .Rows - 1
                    If .IsCollapsed(II) = flexOutlineCollapsed Then
                        nRows = nRows & "," & II
                    End If
                Next II
            Case 1
                For II = .Rows - 1 To .FixedRows Step -1
                    If .IsCollapsed(II) = flexOutlineExpanded And .IsSubtotal(II) Then
                        nRows = nRows & "," & II
                    End If
                Next II
        End Select
        
        nRows = Mid(nRows, 2)
        
        sRows_var = Split(nRows, ",")
    
        For II = 0 To UBound(sRows_var)
            If .IsCollapsed(sRows_var(II)) = flexOutlineCollapsed Then
                .IsCollapsed(sRows_var(II)) = flexOutlineExpanded
            Else
                .IsCollapsed(sRows_var(II)) = flexOutlineCollapsed
            End If
        Next II
    End With
End Sub

Public Function GetPatternProc(ByVal PatternID As String) As String
    Dim oPattern As PlusLib2.CPattern
    Dim rs As ADODB.Recordset
    Dim iLoop%, i%
    Dim sProcess$
    
    On Error GoTo ErrHandler
    
    Set oPattern = New PlusLib2.CPattern
    oPattern.Connection = g_adoCon

    Set rs = oPattern.GetPatternSub(PatternID)
    Set oPattern = Nothing

    
    Do Until rs.EOF
        sProcess = sProcess & "→" & "[" & CheckNull(rs!Process) & "]"
        rs.MoveNext
    Loop
    
    GetPatternProc = Mid(sProcess, 2)
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrHandler:
    'MsgBox "[" & Err.Number & "]" & ":" & Err.Description, vbCritical
    Call ErrorBox(Err.Number, "frmPatternCode.ShowData", Err.Description)
    Set rs = Nothing
    Set oPattern = Nothing

End Function

Public Sub GetNowDate(ByRef oDate As String, ByRef oTime As String)

    Dim oCLogin As PlusLib2.CLogin
    Dim dDateTime As Variant
    
    
    
    Screen.MousePointer = vbHourglass
    
'    On Error GoTo ErrHandler

    Set oCLogin = New PlusLib2.CLogin
    oCLogin.Connection = g_adoCon
    
    '-----------------------------------------------------------------------------------------
    
    dDateTime = oCLogin.GetNow
    
    
    Set oCLogin = Nothing
    
    oDate = Format(dDateTime, "yyyymmdd")

    oTime = Format(dDateTime, "hhmm")
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

End Sub
'******************************************************************************
' 프로시져 명 : FillComboBox
' 기       능 : 주어진 SQL을 실행하여 그 결과를 ComboBox에 채워넣는다
' 사 용 방 법 : Call FillComboBox(arg1,arg2,arg3)
' 사 용 인 수 : arg1 -> ComboBox
'               arg2 -> SQL 문
'               arg3 -> ComboBox의 모든 Item을 나타내는 문자(ALL,* 등)
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Sub FillComboBox(CboBox As Object, ByVal dRS As ADODB.Recordset, Optional AllStr As String = "")
    
    On Error GoTo Process_Fail
    
    CboBox.Clear
    If AllStr <> "" Then CboBox.AddItem AllStr
    
    
    If Not dRS Is Nothing Then
        If Not dRS.BOF Then
           dRS.MoveFirst
           Do Until dRS.EOF
            CboBox.AddItem Trim$(dRS(0))

            dRS.MoveNext
           Loop
        End If
    End If
    If CboBox.ListCount > 0 Then
        CboBox.ListIndex = 0
    Else
        CboBox.ListIndex = -1
    End If
    Exit Sub
Process_Fail:
    CboBox.AddItem "Error!"
End Sub

Public Function GetOrderSeq(ByVal OrderID As String, ByVal ColorName As String) As Integer
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    dSql_str = ""
    dSql_str = dSql_str & vbCr & "   SELECT OrderSeq  "
    dSql_str = dSql_str & vbCr & "     FROM [OrderColor] "
    dSql_str = dSql_str & vbCr & "    WHERE OrderID = '" & Trim(OrderID) & "' "
    dSql_str = dSql_str & vbCr & "      AND Color = '" & Trim(ColorName) & "' "
    
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
         GetOrderSeq = dRS(0)
    End If
    dRS.Close
    Set dRS = Nothing
    
End Function

'/*******************************************************************************
' * Name        : DeletePlanCPB
' * Description : ComboBox에 ProcessName | ProcessID 을 나타낸다.
' * 기       능 : TPlanCPB 구조체에 값 입력 후 처리
' *             : arg1 -> '4000', '4300' 이런식으로 넣는다.
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    날 짜        작성자    버전                   변경사항
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     최현숙    1.0     작성
' ******************************************************************************/
Public Sub SetProcessID(ByVal cboProcessID As ComboBox, ByVal pCodes As String)
    Dim adoCmd As ADODB.Command
    Dim dRS As New ADODB.Recordset
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '품명 삭제. DelClss를 1로 업데이트
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Process_sList"
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, pCodes)
        Set dRS = .Execute()
    
    End With
    
    If dRS.RecordCount > 0 Then
        Call FillComboBox(cboProcessID, dRS)
        cboProcessID.ListIndex = 0
    Else
        cboProcessID.Clear
    End If
    
    GoTo LogMessage
    
ErrHandler:
    bError = True
    
LogMessage:
    Set adoCmd = Nothing
    ReDim sLog(0)
    If bError Then  ' 에러 로그
        Err.Raise Err.Number, "chs_Module.SetProcID", Err.Description
    End If
End Sub

'/*******************************************************************************
' * Name        : DeletePlanCPB
' * Description : ComboBox에 ProcessName | ProcessID 을 나타낸다.
' * 기       능 : TPlanCPB 구조체에 값 입력 후 처리
' *             : arg1 -> '4000', '4300' 이런식으로 넣는다.
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    날 짜        작성자    버전                   변경사항
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     최현숙    1.0     작성
' ******************************************************************************/
Public Sub SetComboProcss(ByVal objCboBox As ComboBox, Optional pAllStr As String = "")
    Dim dRS As New ADODB.Recordset
    Dim bError As Boolean
    Dim dSql_str As String
    
    On Error GoTo ErrHandler
    
    dSql_str = "    SELECT Process" & vbCrLf & _
               "      FROM [mt_process] " & vbCrLf & _
               "     WHERE right(processid, 2) <> '00' " & vbCrLf & _
               "  ORDER BY ProcessID "
               
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
        Call FillComboBox(objCboBox, dRS, pAllStr)
        objCboBox.ListIndex = 0
    Else
        objCboBox.Clear
    End If
    
    dRS.Close
    Set dRS = Nothing
    
    GoTo LogMessage
    
ErrHandler:
    bError = True
    
LogMessage:
    ReDim sLog(0)
    If bError Then  ' 에러 로그
        Err.Raise Err.Number, "chs_Module.SetComboProcss", Err.Description
    End If
    
End Sub


'******************************************************************************
' 프로시져 명 : GetRollQty
' 기       능 : 수량 * 절 의 string을  전체 수량과 절로 되돌려 준다.
' 사 용 방 법 : Call GetRollQty(arg1, arg2, arg3, arg4)
' 사 용 인 수 : arg1 -> 수량 * 절의 문자식
'               arg2 -> 절수
'               arg3 -> 절 단위 수량
'               arg4 -> 전체수량
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
' 'S_201303_태을염직_01 에 의한 수정 (Integer -> Long)
'******************************************************************************
Public Function GetRollQty(ByVal iRollQty_str As String _
                        , ByRef oRoll_int As Integer _
                        , ByRef oQty_int As Long _
                        , ByRef oTotQty_int As Long)
    Dim nPosition%
    Dim nRoll As Long, nQty As Long, nTotQty As Long
        
    nPosition = InStr(iRollQty_str, "*")
    
    If nPosition > 0 Then
        '절수
        nRoll = CheckNum(Mid(iRollQty_str, nPosition + 1))
        nQty = CheckNum(Left(iRollQty_str, nPosition - 1))
        nTotQty = nRoll * nQty
    Else
        nRoll = 1
        nQty = CheckNum(iRollQty_str)
        nTotQty = CheckNum(iRollQty_str)
    End If
    
    oRoll_int = nRoll * IIf(nQty < 0, -1, 1)
    oQty_int = nQty
    oTotQty_int = nTotQty
    
End Function
Public Sub SetInstDefct(ByVal ComboBox As Object, ByVal sProcessID As String)
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_sDefect"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sProcessID)

        Set dRS = .Execute
    End With
    Set adoCmd = Nothing
    
    ComboBox.Clear
    
    If dRS.RecordCount > 0 Then
           Do Until dRS.EOF
            ComboBox.AddItem Trim$(dRS(0))
            dRS.MoveNext
           Loop
    End If
    
    dRS.Close
    Set dRS = Nothing
End Sub

Public Function MakeStuffKey(ByVal iKey As String, ByRef StuffDate As String, ByRef StuffClss As String, ByRef StuffSeq As Integer)
    Dim aa As Variant, II%
    aa = Split(iKey, "-", -1, 1)
    StuffDate = aa(0)
    StuffClss = aa(1)
    StuffSeq = aa(2)
End Function

Public Function GetColor() As Recordset
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_GetColorName"

    '    .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, m_ProcID)

    End With
    Set GetColor = adoCmd.Execute
    Set adoCmd = Nothing
End Function

'******************************************************************************
' 프로시져 명 : GetMachineProcID
' 기       능 : 설비(장비)명에 따른 공정 레코드를 가져온다.
' 사 용 방 법 : GetProcDetail(ByVal dProcID As String) As Recordset
' 사 용 인 수 : dProcID -> 공정코드
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Function GetMachineProcID(ByVal dMachineName As String) As ADODB.Recordset
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    dSql_str = " "
    dSql_str = dSql_str & vbCr & "  SELECT DISTINCT BB.PROCESS "
    dSql_str = dSql_str & vbCr & "    FROM MT_MACHINE AA"
    dSql_str = dSql_str & vbCr & "        ,MT_PROCESS BB "
    dSql_str = dSql_str & vbCr & "   WHERE AA.MACHINE LIKE '%" & dMachineName & "%'  "
    dSql_str = dSql_str & vbCr & "     AND AA.PROCESSID     =  BB.PROCESSID"
    
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
        Set GetMachineProcID = dRS.Clone
    End If
    dRS.Close
    Set dRS = Nothing
    
End Function


'******************************************************************************
' 프로시져 명 : GetProcDetail
' 기       능 : detail 공정 코드 가져오기
' 사 용 방 법 : GetProcDetail(ByVal dProcID As String) As Recordset
' 사 용 인 수 : dProcID -> 공정코드
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Function GetProcDetail(ByVal dProcIDGrp As String) As Recordset
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    dSql_str = ""
    dSql_str = dSql_str & vbCr & "   SELECT CNAME = PROCESS "
 '   dSql_str = dSql_str & vbCr & "       , CODE = PROCESSID "
    dSql_str = dSql_str & vbCr & "     FROM MT_PROCESS"
    dSql_str = dSql_str & vbCr & "    WHERE LEFT(PROCESSID ,2)   =  '" & Left(dProcIDGrp, 2) & "'  "
    dSql_str = dSql_str & vbCr & "      AND PROCESSID    <>  '" & dProcIDGrp & "' "
    dSql_str = dSql_str & vbCr & " ORDER BY PROCESSID "
    
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
        Set GetProcDetail = dRS.Clone
    End If
    dRS.Close
    Set dRS = Nothing
    
End Function

Public Function GetProcessID(ByVal dProcName As String) As String
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    'SELECT dbo.fn_getProcID( '1차정련')
    
    dSql_str = " SELECT dbo.fn_getProcID( '" & RTrim(dProcName) & "' ) "
    
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
        GetProcessID = CheckNull(dRS(0))
    End If
    dRS.Close
    Set dRS = Nothing
    
End Function

'******************************************************************************
' 프로시져 명 : SetGridToggleChecked
' 기       능 : Grid의 check 값 선택 / 해제 Toggle 처리 하는 Procedure
' 사 용 방 법 : Call SetGridToggleChecked( arg1,arg2)
' 사 용 인 수 : arg1 -> VSFlexGrid object
'               arg2 -> [0] 선택선택   [1]: 선택해제SQL 문
'               arg3 -> ComboBox의 모든 Item을 나타내는 문자(ALL,* 등)
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Sub SetGridToggleChecked(ByRef oGrid As VSFlexGrid, Index As Integer, Optional nCol As Integer = 1)
    Dim SetValue, i%
    
    If Index = 0 Then   '[0] 전체선택
        SetValue = flexChecked
    Else                '[1] 선택 해제
        SetValue = flexUnchecked
    End If

    With oGrid
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, nCol) = SetValue
        Next i
    End With
End Sub

Public Function GetPerson(ByVal sFlag As String, ByVal sSearch As String, ByRef oCode As String, ByRef oName As String) As Boolean
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    dSql_str = " SELECT PersonId, Name  " & vbCr & _
               "   FROM [mt_Person]  "
    Select Case UCase(sFlag)
        Case "C"
            dSql_str = dSql_str & "     WHERE PersonID = '" & Trim$(sSearch) & "' "
        Case "R"
            dSql_str = dSql_str & "     WHERE Name = '" & Trim$(sSearch) & "' "
    End Select
                   
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount = 1 Then
        oCode = dRS(0)
        oName = dRS(1)
    End If
               
    dRS.Close
    Set dRS = Nothing
    
    Exit Function
    
ErrHandler:
    If Not dRS Is Nothing Then
        Set dRS = Nothing
    End If
    
    Call ErrorBox(Err.Number, "frmPlanCPB.FillPlanCPB", Err.Description)
    
End Function

'******************************************************************************
' 프로시져 명 : ClearScreen
' 기       능 : Form의 각 Field를 Clear 시킴(TextBox, ComboBox, MaskEdBox)
' 사 용 방 법 : Call ClearScreen(arg1)
' 사 용 인 수 : arg1 -> Form Name
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Sub ClearScreen(pForm As Form, Optional pContainer As String = "")
    
    Dim i As Integer
    Dim TmpMsk As String
    Dim dControl As Object
    
    For i = 0 To pForm.Controls.Count - 1

        If pContainer <> "" And pForm.Controls(i).Container.Name <> pContainer Then
                GoTo for_loop
        End If
        
        If TypeOf pForm.Controls(i) Is TextBox Or TypeOf pForm.Controls(i) Is WizText Then
            pForm.Controls(i) = ""
        ElseIf TypeOf pForm.Controls(i) Is ComboBox Then
                If pForm.Controls(i).Style = 0 Then
                    pForm.Controls(i) = ""
                    pForm.Controls(i).ListIndex = IIf(pForm.Controls(i).ListCount > 0, 0, -1)
                Else 'Dropdown List
                    pForm.Controls(i).ListIndex = IIf(pForm.Controls(i).ListCount > 0, 0, -1)
                End If
        ElseIf TypeOf pForm.Controls(i) Is MaskEdBox Then
                TmpMsk = pForm.Controls(i).Mask
                pForm.Controls(i).Mask = ""
                pForm.Controls(i).Text = ""
                pForm.Controls(i).Mask = TmpMsk
        ElseIf TypeOf pForm.Controls(i) Is CheckBox Then
                pForm.Controls(i).Value = 0
        End If
for_loop:
        
    Next
End Sub

Public Sub FixedColAlignMentSetting(vsGrid As VSFlexGrid)
    Dim iCount As Integer
    For iCount = 0 To vsGrid.Cols - 1
        vsGrid.FixedAlignment(iCount) = flexAlignCenterCenter
    Next iCount
End Sub
'******************************************************************************
' 함   수  명 : FindItem
' 기       능 : ComboBox내에서 주어진 값을 찾는다
' 사 용 방 법 : i = FindItem(arg1,arg2,arg3)
' 사 용 인 수 : arg1 -> ComboBox Name
'               arg2 -> 찾고자하는 값
'               arg3 -> 비교할 문자열 길이(생략하면 전체비교)
' 리   턴  값 : i = 찾았을경우 ComboBox의 Listindex를 리턴
'               i = 찾지못했을경우 -1 리턴
'******************************************************************************
'    날 짜        작성자    버전                   변경사항
'--------------  --------  ------  --------------------------------------------
'
'******************************************************************************
Public Function FindItem(CboBox As Control, Value$, Optional CompareLen) As Integer
    
    Dim i As Integer
    
    If IsMissing(CompareLen) Then
        For i = 0 To CboBox.ListCount - 1
            If (Value = CboBox.List(i) Or InStr(CboBox.List(i), Value$) > 0) Then
                FindItem = i
                Exit Function
            End If
        Next
    Else
        For i = 0 To CboBox.ListCount - 1
            If Left(Value, CompareLen) = Left(CboBox.List(i), CompareLen) Then
                FindItem = i
                Exit Function
            End If
        Next
    End If
    
    FindItem = -1 'Not Found
    
End Function


'****************************************************************************************'
''Public Sub SetStuffINReturnGoods(ByVal sStuffDate As String, ByVal sStuffClss As String, ByVal nStuffSeq As Integer)
''    Dim oStuffIn As PlusLib2.CStuffIN
''    Dim RsHeader As ADODB.Recordset, RsDetail As ADODB.Recordset
''    Dim rsData As ADODB.Recordset
''    Dim nRollvar(), II%
''
'''    On Error GoTo ErrHandler
''
''    Set oStuffIn = New PlusLib2.CStuffIN
''    oStuffIn.Connection = g_adoCon
''    oStuffIn.UserName = g_sUserName
''
''    Set RsDetail = oStuffIn.GetStuffINReturnGoods(sStuffDate, sStuffClss, nStuffSeq, RsHeader)
''
''
''    ''''' StuffINSub 의 데이터를 읽어 온다
''    Set rsData = oStuffIn.GetStuffINSubONE(sStuffDate, sStuffClss, nStuffSeq)
''
''    ReDim nRollvar(rsData.RecordCount)
''    II = 0
''    Do Until rsData.EOF
''        nRollvar(II) = rsData!Qty
''        rsData.MoveNext
''        II = II + 1
''    Loop
''
''    rsData.Close
''    Set rsData = Nothing
''
''    Call SetPrint(RsHeader, nRollvar)
''
''    Set oStuffIn = Nothing
''    Exit Sub
''ErrHandler:
''    MsgBox ("반품 명세서 출력 중 오류 발생 ")
''End Sub

Public Sub SetPrnHeader(ByVal RsHeader As Recordset, ByVal nPage As Integer)
    Printer.Orientation = vbPRORPortrait
    Printer.ScaleMode = vbMillimeters
    Printer.PaperSize = vbPRPSLetter
'    Printer.Width = 242
'    Printer.Height = 140
'    Printer.ScaleHeight = 150
'    Printer.ScaleWidth = 250
    

    'Page
    Call PrnData(175, 30, "PAGE: " & CStr(nPage))
    
    '발행일자
    Call PrnData(30, 47, MakeDate(DF_FULL, Now))
    
    '일련번호
    Call PrnData(82, 47, RsHeader!StuffDate & "-" & RsHeader!StuffClss & "-" & RsHeader!StuffSeq)
    
    ' 등록번호
    Call PrnData(126, 55, Format(RsHeader!CustomNo, "###-##-#####"))
    
    ' 상호  &  대표자성명
    Call PrnData(126, 65, RsHeader!kCustom)
    Call PrnData(164, 65, RsHeader!Cheif)
    
    ' 주소
''    'S_201312_태을염직_99 에 의한 수정-OLD소스
''    Call PrnData(126, 68, Trim(RsHeader!Address1))
''    Call PrnData(126, 68, Trim(RsHeader!Address2))
        If Trim(RsHeader!Address1) <> "" Then            '도로명 주소값 있을경우
        Call PrnData(126, 68, Trim(RsHeader!Address1))
        Call PrnData(126, 68, Trim(RsHeader!Address2))
    Else
        Call PrnData(126, 68, Trim(RsHeader!AddressJiBun1))
        Call PrnData(126, 68, Trim(RsHeader!AddressJiBun2))
    End If
    
    
    ' 업태 &  종목
    Call PrnData(126, 73, Trim(RsHeader!Condition))
    Call PrnData(164, 73, Trim(RsHeader!Category))
    
    '  품명 & 가공구분 & 규격 & OrderNO & Order수량 & 절수 & 비고
    Call PrnData(20, 89, RsHeader!Article)
    Call PrnData(55, 89, RsHeader!WorkName)
    Call PrnData(82, 89, RsHeader!StuffWidth)
    Call PrnData(95, 89, RsHeader!OrderNo)
    Call PrnData(122, 89, SetCurrency(ChkNullValue(RsHeader!OrderQty), 0) & ChkNullValue(RsHeader!UnitClss))
    Call PrnData(145, 89, SetCurrency(RsHeader!TotRoll, 0))
    Call PrnData(155, 89, SetCurrency(RsHeader!TotQty, 0))

End Sub

Public Sub SetPrint(ByVal RsHeader As Recordset, ByRef nRollvar())
    Dim intBlank$, dRoll_str As String, II%, nLinePos As Long, xPos As Long
    Dim PrnDate As String, nPrnLines As Integer, nPage As Integer
    
    
    nPage = 1
    Call SetPrnHeader(RsHeader, nPage)

    nLinePos = 105  '첫번째 라인 출력 위치  5씩 증가,, 10라인 출력 후 다음 페이지 인쇄
    dRoll_str = ""
    nPrnLines = 1
    For II = 0 To UBound(nRollvar)
        Select Case (II + 1) Mod 10
            Case 1: xPos = 88
            Case 2: xPos = 98
            Case 3: xPos = 108
            Case 4: xPos = 118
            Case 5: xPos = 128
            Case 6: xPos = 138
            Case 7: xPos = 148
            Case 8: xPos = 158
            Case 9: xPos = 168
            Case 0: xPos = 178
        End Select
        
        Call PrnData(xPos, nLinePos, nRollvar(II))
        If (II + 1) Mod 10 = 0 Then
            nLinePos = nLinePos + 6
            nPrnLines = nPrnLines + 1
        End If
        
        If nPrnLines > 23 Then
            Printer.NewPage
            nPrnLines = 1
            nPage = nPage + 1
            Call SetPrnHeader(RsHeader, nPage)
        End If
    Next II
    
    Printer.EndDoc
End Sub

Sub PrnData(ByVal xPos As Long, ByVal yPos As Long, ByVal dStr As String)
    Printer.CurrentX = xPos
    Printer.CurrentY = yPos
    Printer.Print Trim$(dStr)
End Sub



Public Function ChkNullValue(ChkCol As Field) As Variant

    
    If IsNull(ChkCol) Then
        If ChkCol.Type = adChar Or ChkCol.Type = adVarChar Then
            ChkNullValue = ""
        Else
            ChkNullValue = 0
        End If
    Else
        If ChkCol.Type = adChar Or ChkCol.Type = adVarChar Then
            ChkNullValue = Trim$(ChkCol)
        Else
            ChkNullValue = ChkCol
        End If
    End If
End Function


Public Sub ColResize_ColWidth(ByVal oFlex As VSFlexGrid, ByVal nType As ESHRINK, Optional nPercent As Integer = 10, Optional vColWidth)
    Dim II%


    '--- 확장
    If nType = ES_EXPAND Then
        With oFlex
            For II = .FixedCols To .Cols - 1
                .Redraw = flexRDBuffered
                .ColWidth(II) = vColWidth(II)
            Next II
        End With
    Else
        '--- 축소
        With oFlex
            For II = .FixedCols To .Cols - 1
                .Redraw = flexRDBuffered
                vColWidth(II) = .ColWidth(II)
                .ColWidth(II) = Int(.ColWidth(II) / (1 + nPercent / 100))
            Next II
        End With
    End If
End Sub


Private Function PrintDot(nXPos As Integer, nYPos As Integer, sStr As String, Optional nFont As Integer = 10)
    With Printer
        .CurrentX = nXPos
        .CurrentY = nYPos
        .Font.Size = nFont
    End With
    Printer.Print sStr
End Function

'****************************************************************************************'
