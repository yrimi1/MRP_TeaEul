Attribute VB_Name = "mod_TmsComm"
'**************************************************************************************************
'** System 명 : MRRPLUS2
'** Author    : Wizard
'** 작성자    : 이경미
'** 내용      : 섬유관리시스템에서 공통으로 사용하는 Function, Sub 의 모음이다.
'** 생성일자  : 2010.05.27
'**------------------------------------------------------------------------------------------------
'** 변경일자    변경자  변경내용
'**************************************************************************************************
Option Explicit

Public Const PLANCPB = "4000"
Public Const PLANRAPID = "4300"

Public Const PRNRowHeight = 400              '인쇄시 RowHeight
'''Public Const PRNHeaderColor = &HB4B4B4       '&HAAAAAA    '&H9F9F9F    '&H8F8F8F               '인쇄시 Header Title Color
Public Const FrozenColor = &H8000000F
Public Function SetFloat2(nNumber As Variant) As String
    Dim nTmp#, nTmp1%, nTmp2%, iCount%
    Dim sBaseFmt As String
    
    sBaseFmt = IIf(nNumber = 0 Or nNumber = Null, "0", "#,##0.000")
    SetFloat2 = Format(nNumber, sBaseFmt)
End Function
Public Function MakeDyeAuxClss(sDyeAuxClss As String) As String
    If sDyeAuxClss = "0" Then
        MakeDyeAuxClss = " "
    ElseIf sDyeAuxClss = "1" Then
        MakeDyeAuxClss = "고압분산염료"
    ElseIf sDyeAuxClss = "2" Then
        MakeDyeAuxClss = "반응성염료"
    ElseIf sDyeAuxClss = "3" Then
        MakeDyeAuxClss = "산성염료"
    ElseIf sDyeAuxClss = "4" Then
        MakeDyeAuxClss = "직접염료"
    ElseIf sDyeAuxClss = "5" Then
        MakeDyeAuxClss = "카치온염료"
    ElseIf sDyeAuxClss = "6" Then
        MakeDyeAuxClss = "아세테이트"
    ElseIf sDyeAuxClss = "8" Then
        MakeDyeAuxClss = "NaOH"
    End If
End Function
'Public g_sysDate As String, g_sysTime As String
 
'//TODO: 한영혼합문장의 start 위치에서부터 length 만큼 문자를 읽어온다.
Function MidH(s, start, length)
    Dim f, CharAt, VBLength, VBn1, VBn2, BLength, AddByte
    VBn2 = length
    VBLength = Len(s)
    BLength = 0
    For f = 1 To VBLength
        CharAt = Mid(s, f, 1)
        If Asc(CharAt) > 0 And Asc(CharAt) < 255 Then
            BLength = BLength + 1
        Else
            BLength = BLength + 2
        End If
        If BLength >= start Then
            Exit For
        End If
    Next
 
    VBn1 = f
    If VBn1 < 1 Then VBn1 = 1
 
    BLength = 0
    For f = VBn1 To VBLength
        CharAt = Mid(s, f, 1)
        If Asc(CharAt) > 0 And Asc(CharAt) < 255 Then
            BLength = BLength + 1
        Else
            BLength = BLength + 2
        End If
        If BLength = length Then
            VBn2 = f + 1
            Exit For
        ElseIf BLength > length Then
            VBn2 = f
            Exit For
        End If
    Next
    MidH = Mid(s, VBn1, VBn2 - VBn1)
End Function
Public Function Gf_Excel_CopySheet(oExcel As Excel.Application, psFromSheet As String, psToSheet As String, nPage As Integer, EXCEL_ROLL_ROW As Integer)
    Dim i%, nBaseRow%

    nBaseRow = GF_Excel_GetBaseRow(nPage, EXCEL_ROLL_ROW)
    With oExcel
        .Sheets(psFromSheet).Select

        .Rows("1:" & CStr(EXCEL_ROLL_ROW)).Select
        .Selection.Copy

        .Sheets(psToSheet).Select
        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
        .Selection.Insert Shift:=xlDown
        
       ' .Cells(nBaseRow + 2, 50) = nPage
    End With
End Function

'************************************************************************
'--- 해당월의 N번주차 from -> to Date 찾아 내기
'--- 200411월 1주차: 일요일->월요일 까지 ( 20041031-> 20041106 )
'************************************************************************
Public Function GetWeekTerm(ByVal pDate As String, ByRef vWeek As Variant) As Boolean
    
    '해당월의 1일이 무슨 요일인지 확인
    Dim nDate As Integer '요일
    Dim II%, JJ%
    Dim sDate As String, eDate As String, mDate As Date
    
    JJ = 0
    
    ReDim vWeek(5)
    
    For II = 1 To 31 Step 7
        mDate = Format(Left(pDate, 6) & Format(II, "0#"), "####-##-##")
        
        sDate = DateAdd("d", (Weekday(mDate) - 1) * -1, mDate)
        eDate = DateAdd("d", 6, DateAdd("d", (Weekday(mDate) - 1) * -1, mDate))
        
        vWeek(JJ) = sDate & "  ~  " & eDate
        
        JJ = JJ + 1
        
    Next II
    
End Function

Public Function GetDayWeek(ByVal sDate As String) As String
    Dim vDayWeek As Variant
    
    vDayWeek = Split(",일,월,화,수,목,금,토", ",")
    
    GetDayWeek = vDayWeek(Weekday(Format(sDate, "####-##-##")))
    
End Function

Public Function GF_Excel_GetBaseRow(ByVal nPage As Integer, ByVal nRow As Integer)
    GF_Excel_GetBaseRow = (nPage - 1) * nRow
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

Function ALP_TO_HAN(TOTAL As Double) As String
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
    
    W_UNIT(1) = "백"   '백
    W_UNIT(2) = "십"   '십
    W_UNIT(3) = "억"   '억
    W_UNIT(4) = "천"   '천
    W_UNIT(5) = "백"   '백
    W_UNIT(6) = "십"   '십
    W_UNIT(7) = "만"   '만
    W_UNIT(8) = "천"   '천
    W_UNIT(9) = "백"   '백
    W_UNIT(10) = "십"  '십
    W_UNIT(11) = ""
    
    W_NUM(1) = "일"  '일
    W_NUM(2) = "이"  '이
    W_NUM(3) = "태을염직"  '태을염직
    W_NUM(4) = "사"  '사
    W_NUM(5) = "오"  '오
    W_NUM(6) = "육"  '육
    W_NUM(7) = "칠"  '칠
    W_NUM(8) = "팔"  '팔
    W_NUM(9) = "구"  '구
    
    STR_TOTAL = W_TOTAL
    TMP = Space(11 - Len(STR_TOTAL)) & W_TOTAL
    
    For II = 1 To 11
        n = Trim$(Mid$(TMP, II, 1))
        If n <> "0" And n <> "" Then
            KUM$ = KUM$ & W_NUM(val(n)) & W_UNIT(II)
        End If
        If (II = 3 And KUM$ <> "" And Right(KUM$, 1) <> "억") Then
            KUM$ = KUM$ & "억"
        End If
        If (II = 7 And KUM$ <> "" And Right(KUM$, 1) <> "만" And _
                      Right(KUM$, 1) <> "억") Then
            KUM$ = KUM$ & "만"
        End If
    Next II
    'ALP_TO_STR = KUM$
    'If W_MINUS < 0 Then
    ALP_TO_HAN = Trim$(W_MINUS$ & KUM$)
    'End If
End Function



''''''''Public Sub SetGrdColor(ByVal oFlex As VSFlexGrid, ByVal sDepth As String _
''''''''                , ByVal nRow1 As Integer, ByVal nCol1 As Integer _
''''''''                , ByVal nRow2 As Integer, ByVal nCol2 As Integer)
''''''''    Dim nColorVal As Long
''''''''    With oFlex
''''''''        Select Case sDepth
''''''''            Case "1": nColorVal = ED1_DEPTH
''''''''            Case "2": nColorVal = ED2_DEPTH
''''''''            Case "3": nColorVal = ED3_DEPTH
''''''''            Case "4": nColorVal = ED4_DEPTH
''''''''            Case Else: nColorVal = vbWhite
''''''''        End Select
''''''''        .Cell(flexcpBackColor, nRow1, nCol1, nRow2, nCol2) = nColorVal
''''''''    End With
''''''''End Sub


''Public Sub ColResize(ByVal oFlex As VSFlexGrid, ByVal nType As ESHRINK, Optional nPercent As Integer = 10)
''    Dim II%
''
''    If nType = ES_EXPAND Then
''        With oFlex
''            For II = .FixedCols To .Cols - 1
''                .Redraw = flexRDBuffered
''                .ColWidth(II) = Int(.ColWidth(II) * (1 + nPercent / 100))
''            Next II
''        End With
''    Else
''        With oFlex
''            For II = .FixedCols To .Cols - 1
''                .Redraw = flexRDBuffered
''                .ColWidth(II) = Int(.ColWidth(II) / (1 + nPercent / 100))
''            Next II
''        End With
''    End If
''End Sub

Public Sub SetDateTerm(ByVal Interval As Integer, pFromDate As DTPicker, pToDate As DTPicker)
    Dim dFromDate As Date, dToDate As Date
    
    Select Case Interval
        Case -1
            dToDate = Date
            dFromDate = DateAdd("d", -30, dToDate)
        Case -2
            dToDate = Date
            dFromDate = DateAdd("d", -60, dToDate)
    End Select
    
    pFromDate = dFromDate
    pToDate = dToDate

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
''
''Public Sub GridCollapse(ByVal oFlex As VSFlexGrid, ByVal Row As Integer)
''    With oFlex
''        If Row < .FixedRows Then Exit Sub
''
''        If .IsCollapsed(Row) = flexOutlineCollapsed Then
''            .IsCollapsed(Row) = flexOutlineExpanded
''        Else
''            .IsCollapsed(Row) = flexOutlineCollapsed
''        End If
''    End With
''End Sub

Public Function GetKorWeekDay(psdate As Date) As String
Dim lsRtnKorWeekDay As String

    On Error GoTo Err_Rtn
    
    Weekday (Date)
    
    Select Case Weekday(Date)
        Case "1"
            lsRtnKorWeekDay = "일"
        Case "2"
            lsRtnKorWeekDay = "월"
        Case "3"
            lsRtnKorWeekDay = "화"
        Case "4"
            lsRtnKorWeekDay = "수"
        Case "5"
            lsRtnKorWeekDay = "목"
        Case "6"
            lsRtnKorWeekDay = "금"
        Case "7"
            lsRtnKorWeekDay = "토"
    End Select
    
    GetKorWeekDay = lsRtnKorWeekDay
    
Err_Rtn:
    

End Function




Public Function GF_Chk_Keypress(pKeyAscii As Integer, psGbn As String) As Integer
'****************************************************************
'입력한 값 Check
'일자   : 2010.07.09
'작업자 : 이경미
'9.     : 숫자, . 만 가능 그외 '' 로 return
'A      : 문자. A~Z 만 가능
'a      : 문자. a~z 만 가능
'Ex) 9.A 이면 숫자, . 영문 대문자만 가능.
'****************************************************************
    If pKeyAscii = 22 Then  'Ctrl+v
        GF_Chk_Keypress = pKeyAscii
         Exit Function
    End If
    If InStr(1, psGbn, "9") > 0 Then  '9가 있으면
        If pKeyAscii >= 48 And pKeyAscii <= 57 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
        
        
    If InStr(1, psGbn, "-") > 0 Then  '-가 있으면
        If pKeyAscii = 45 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
    
    If InStr(1, psGbn, ".") > 0 Then  '.가 있으면
        If pKeyAscii = 46 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
        
    If InStr(1, psGbn, "A") > 0 Then  'A 가 있으면
        If pKeyAscii >= 65 And pKeyAscii <= 90 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
        
    If InStr(1, psGbn, "a") > 0 Then  'a 가 있으면
        If pKeyAscii >= 97 And pKeyAscii <= 122 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
    
    If pKeyAscii = 32 Then       'Space
         GF_Chk_Keypress = pKeyAscii
         Exit Function
     End If
        
    If pKeyAscii = 46 Then       'Del
         GF_Chk_Keypress = pKeyAscii
         Exit Function
     End If
        
    If pKeyAscii = 8 Then       'Back Space
         GF_Chk_Keypress = pKeyAscii
         Exit Function
     End If
        
        
    GF_Chk_Keypress = 0

End Function

Public Function GF_Set_GridPosition(psGridObj As VSFlexGrid, psKeyValue() As String, psaKeyCol() As String) As Boolean
'****************************************************************
'Parameter 로 넘어오는 Grid 를 Key 값 위치로 Position 시킴  함
'일자   : 2010.07.09
'작업자 : 이경미
'****************************************************************
Dim lnrows                                  As Long
Dim liKeyValue                              As Integer
Dim i                                       As Integer
Dim lsKeyString                             As String
Dim lsGridString                            As String

    On Error GoTo Err_Rtn
    
    liKeyValue = UBound(psKeyValue)
    
    'Key String
    For i = LBound(psKeyValue) To UBound(psKeyValue)
        If i = LBound(psKeyValue) Then
            lsKeyString = psKeyValue(i)
        Else
            lsKeyString = lsKeyString & " | " & psKeyValue(i)
        End If
    Next i
    
    With psGridObj
        For lnrows = psGridObj.FixedRows To psGridObj.Rows - 1
        
            'Grid String
            For i = LBound(psKeyValue) To UBound(psKeyValue)
                If i = LBound(psKeyValue) Then
                    lsGridString = .TextMatrix(lnrows, psaKeyCol(i))
                Else
                    lsGridString = lsGridString & " | " & .TextMatrix(lnrows, psaKeyCol(i))
                End If
            Next i
                
            'Key String 과 Grid string 이 같으면
        
            If lsKeyString = lsGridString Then
                .Row = lnrows       'All 값이 Matching 될 때
                .TopRow = lnrows
                Exit For
            End If
Next_rows:
        Next lnrows
    
    End With
    Exit Function
Err_Rtn:
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[GF_Set_GridPosition]"
    End If
End Function

'****************************************************************
'*Author : 위저드 정보 시스템
'*작성자 : 이경미
'*Description:
'*  INI 파일에서 해당 Section의 Key에 해당하는 값을 읽어온다.
'*
'****************************************************************
Public Function GetValue(NewSection As String, NewKey As String, Optional NewDefault) As String
    Dim ReturnLength As Long, ReturnValue As String
    
    ReturnValue = String$(255, &H0)
    If GetPrivateProfileString(NewSection, NewKey, "", ReturnValue, Len(ReturnValue), m_sAppFile) = 0 Then
        If IsMissing(NewDefault) Then '주어진 Default값이 없을 경우
            GetValue = ""
        Else '주어진 Default값이 있을 경우
            GetValue = NewDefault
        End If
    Else
        GetValue = Left(ReturnValue, InStr(ReturnValue, Chr(0)) - 1)
    End If
End Function

Public Function GF_Excel_CA(plnNum As Integer) As String
'****************************************************************
'*Author : 위저드 정보 시스템
'*작성자 : 이경미
'*작성일 : 2010.12.30
'*Description:
'*  Parameter 로 받은 숫자를 기준으로 Excel 에서 사용하는 Alphabet으로 변경한다.
'*
'****************************************************************
Dim lsRtnVal        As String
Dim liMok           As String
Dim liRemain        As String
Dim stemp           As String
Dim i As Integer

    
    If plnNum > 256 Then
        MsgBox "256이상의 숫자는 Excel Column 명 변경이 불가능 합니다. 현재수:" & plnNum, vbInformation, "[확인]"
        GoTo Err_Rtn
    End If

    '넘어오는 값으로 26로 나누어 몫은 앞부분, 나머지는 뒤에 숫자로 나열하여 Return 함
    Select Case plnNum   ' 숫자를 평가합니다.
        Case 1 To 26
                    GF_Excel_CA = Chr(plnNum + 64)
        Case 27 To 52
                    GF_Excel_CA = "A" & Chr(plnNum - 26 + 64)
        Case 53 To 78
                    GF_Excel_CA = "B" & Chr(plnNum - 52 + 64)
        Case 79 To 104
                    GF_Excel_CA = "C" & Chr(plnNum - 78 + 64)
        Case 105 To 130
                    GF_Excel_CA = "D" & Chr(plnNum - 104 + 64)
        Case 131 To 156
                    GF_Excel_CA = "E" & Chr(plnNum - 130 + 64)
        Case 157 To 182
                    GF_Excel_CA = "F" & Chr(plnNum - 156 + 64)
        Case 183 To 208
                    GF_Excel_CA = "G" & Chr(plnNum - 182 + 64)
        Case 209 To 234
                    GF_Excel_CA = "H" & Chr(plnNum - 208 + 64)
        Case 235 To 256
                    GF_Excel_CA = "I" & Chr(plnNum - 234 + 64)
        Case Else
             GF_Excel_CA = "ZZ"
    End Select
    
    Exit Function
    
Err_Rtn:

    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[GF_Excel_CA]"
    End If

End Function
 
'****************************************************************
'*Author: Shaikan
'*
'*Description:
'*  INI 파일에서 해당 Section의 Key에 해당하는 값을 읽어온다.
'*
'****************************************************************
Public Sub SetValue(sSection As String, sKey As String, sValue As String, sFileName As String)
    Call WritePrivateProfileString(sSection, sKey, sValue, sFileName)
End Sub


Public Function GF_Get_UnitName(psUnitID As String) As String
        
    If psUnitID = "0" Then
        GF_Get_UnitName = "Yd"       '야드
    
    ElseIf psUnitID = 1 Then
        GF_Get_UnitName = "Mt"       '미터
    
    ElseIf psUnitID = "2" Then
    
        GF_Get_UnitName = "Kg"
        
    Else
        GF_Get_UnitName = psUnitID
    End If

End Function




Public Sub FlexTOExcel(Grid As Object)
On Error GoTo errTrap
  Dim Temp01 As String
  Dim xl As Object
  Dim Wb As Object
  Dim ws As Object
  Dim i As Long, j As Long, k As Long
  
  Screen.MousePointer = vbHourglass
  Set xl = GetObject(, "Excel.application")
  
  xl.Visible = True
  Set Wb = xl.Workbooks.Add()
  Set ws = Wb.Worksheets.Add
  
  For i = 0 To Grid.Rows - 1
     Grid.Row = i
     k = i
     
     For j = 0 To Grid.Cols - 1
         If j <= 25 Then          'A
            Temp01 = Chr(65 + j)
         ElseIf j <= 52 Then
            Temp01 = Chr(65) + Chr(65 + (j - 26)) 'AA
         ElseIf j <= 78 Then
            Temp01 = Chr(66) + Chr(65 + (j - 52)) 'ba
         End If
     
         xl.Range(Temp01 & CStr(i + 1)).Value = Grid.TextMatrix(i, j)
  '      xl.Range(Chr(65 + J) & CStr(i + 1)).Value = Grid.TextMatrix(i, J)
      Next
     
  Next
  

  
  Set ws = Nothing
  Set Wb = Nothing
  Set xl = Nothing
  
  Screen.MousePointer = vbDefault
  Exit Sub
  
errTrap:
  
  If Err = 432 Or Err = 429 Then
     Set xl = CreateObject("Excel.Application")
     Resume Next
  Else
     Screen.MousePointer = vbDefault
     MsgBox "Error..........    ", vbCritical
  End If
  
  
  
  
End Sub


