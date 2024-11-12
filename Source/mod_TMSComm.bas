Attribute VB_Name = "mod_TmsComm"
'**************************************************************************************************
'** System �� : MRRPLUS2
'** Author    : Wizard
'** �ۼ���    : �̰��
'** ����      : ���������ý��ۿ��� �������� ����ϴ� Function, Sub �� �����̴�.
'** ��������  : 2010.05.27
'**------------------------------------------------------------------------------------------------
'** ��������    ������  ���泻��
'**************************************************************************************************
Option Explicit

Public Const PLANCPB = "4000"
Public Const PLANRAPID = "4300"

Public Const PRNRowHeight = 400              '�μ�� RowHeight
'''Public Const PRNHeaderColor = &HB4B4B4       '&HAAAAAA    '&H9F9F9F    '&H8F8F8F               '�μ�� Header Title Color
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
        MakeDyeAuxClss = "��кл꿰��"
    ElseIf sDyeAuxClss = "2" Then
        MakeDyeAuxClss = "����������"
    ElseIf sDyeAuxClss = "3" Then
        MakeDyeAuxClss = "�꼺����"
    ElseIf sDyeAuxClss = "4" Then
        MakeDyeAuxClss = "��������"
    ElseIf sDyeAuxClss = "5" Then
        MakeDyeAuxClss = "īġ�¿���"
    ElseIf sDyeAuxClss = "6" Then
        MakeDyeAuxClss = "�Ƽ�����Ʈ"
    ElseIf sDyeAuxClss = "8" Then
        MakeDyeAuxClss = "NaOH"
    End If
End Function
'Public g_sysDate As String, g_sysTime As String
 
'//TODO: �ѿ�ȥ�չ����� start ��ġ�������� length ��ŭ ���ڸ� �о�´�.
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
'--- �ش���� N������ from -> to Date ã�� ����
'--- 200411�� 1����: �Ͽ���->������ ���� ( 20041031-> 20041106 )
'************************************************************************
Public Function GetWeekTerm(ByVal pDate As String, ByRef vWeek As Variant) As Boolean
    
    '�ش���� 1���� ���� �������� Ȯ��
    Dim nDate As Integer '����
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
    
    vDayWeek = Split(",��,��,ȭ,��,��,��,��", ",")
    
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
    
    W_UNIT(1) = "��"   '��
    W_UNIT(2) = "�"   '��
    W_UNIT(3) = "��"   '��
    W_UNIT(4) = "��"   'õ
    W_UNIT(5) = "��"   '��
    W_UNIT(6) = "�"   '��
    W_UNIT(7) = "ؿ"   '��
    W_UNIT(8) = "��"   'õ
    W_UNIT(9) = "��"   '��
    W_UNIT(10) = "�"  '��
    W_UNIT(11) = ""
    
    W_NUM(1) = "��"  '��
    W_NUM(2) = "��"  '��
    W_NUM(3) = "߸"  '��������
    W_NUM(4) = "��"  '��
    W_NUM(5) = "��"  '��
    W_NUM(6) = "�"  '��
    W_NUM(7) = "��"  'ĥ
    W_NUM(8) = "��"  '��
    W_NUM(9) = "��"  '��
    
    STR_TOTAL = W_TOTAL
    TMP = Space(11 - Len(STR_TOTAL)) & W_TOTAL
    
    For II = 1 To 11
        n = Trim$(Mid$(TMP, II, 1))
        If n <> "0" And n <> "" Then
            KUM$ = KUM$ & W_NUM(val(n)) & W_UNIT(II)
        End If
        If (II = 3 And KUM$ <> "" And Right(KUM$, 1) <> "��") Then
            KUM$ = KUM$ & "��"
        End If
        If (II = 7 And KUM$ <> "" And Right(KUM$, 1) <> "ؿ" And _
                      Right(KUM$, 1) <> "��") Then
            KUM$ = KUM$ & "ؿ"
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
    
    W_UNIT(1) = "��"   '��
    W_UNIT(2) = "��"   '��
    W_UNIT(3) = "��"   '��
    W_UNIT(4) = "õ"   'õ
    W_UNIT(5) = "��"   '��
    W_UNIT(6) = "��"   '��
    W_UNIT(7) = "��"   '��
    W_UNIT(8) = "õ"   'õ
    W_UNIT(9) = "��"   '��
    W_UNIT(10) = "��"  '��
    W_UNIT(11) = ""
    
    W_NUM(1) = "��"  '��
    W_NUM(2) = "��"  '��
    W_NUM(3) = "��������"  '��������
    W_NUM(4) = "��"  '��
    W_NUM(5) = "��"  '��
    W_NUM(6) = "��"  '��
    W_NUM(7) = "ĥ"  'ĥ
    W_NUM(8) = "��"  '��
    W_NUM(9) = "��"  '��
    
    STR_TOTAL = W_TOTAL
    TMP = Space(11 - Len(STR_TOTAL)) & W_TOTAL
    
    For II = 1 To 11
        n = Trim$(Mid$(TMP, II, 1))
        If n <> "0" And n <> "" Then
            KUM$ = KUM$ & W_NUM(val(n)) & W_UNIT(II)
        End If
        If (II = 3 And KUM$ <> "" And Right(KUM$, 1) <> "��") Then
            KUM$ = KUM$ & "��"
        End If
        If (II = 7 And KUM$ <> "" And Right(KUM$, 1) <> "��" And _
                      Right(KUM$, 1) <> "��") Then
            KUM$ = KUM$ & "��"
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
' ��   ��  �� : GridCollapse
' ��       �� : �ش����� Shrink�Ѵ�
' �� �� �� �� : DoFlexGridGroup(arg1,arg2,arg3)
' �� �� �� �� : arg1 -> VsFlexGrid Object
'               arg2 -> subTotal�� ������ ��
'               arg3 -> group �ܰ�
' ��   ��  �� : ����
'******************************************************************************
'    �� ¥        �ۼ���    ����                   �������
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
            lsRtnKorWeekDay = "��"
        Case "2"
            lsRtnKorWeekDay = "��"
        Case "3"
            lsRtnKorWeekDay = "ȭ"
        Case "4"
            lsRtnKorWeekDay = "��"
        Case "5"
            lsRtnKorWeekDay = "��"
        Case "6"
            lsRtnKorWeekDay = "��"
        Case "7"
            lsRtnKorWeekDay = "��"
    End Select
    
    GetKorWeekDay = lsRtnKorWeekDay
    
Err_Rtn:
    

End Function




Public Function GF_Chk_Keypress(pKeyAscii As Integer, psGbn As String) As Integer
'****************************************************************
'�Է��� �� Check
'����   : 2010.07.09
'�۾��� : �̰��
'9.     : ����, . �� ���� �׿� '' �� return
'A      : ����. A~Z �� ����
'a      : ����. a~z �� ����
'Ex) 9.A �̸� ����, . ���� �빮�ڸ� ����.
'****************************************************************
    If pKeyAscii = 22 Then  'Ctrl+v
        GF_Chk_Keypress = pKeyAscii
         Exit Function
    End If
    If InStr(1, psGbn, "9") > 0 Then  '9�� ������
        If pKeyAscii >= 48 And pKeyAscii <= 57 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
        
        
    If InStr(1, psGbn, "-") > 0 Then  '-�� ������
        If pKeyAscii = 45 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
    
    If InStr(1, psGbn, ".") > 0 Then  '.�� ������
        If pKeyAscii = 46 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
        
    If InStr(1, psGbn, "A") > 0 Then  'A �� ������
        If pKeyAscii >= 65 And pKeyAscii <= 90 Then
            GF_Chk_Keypress = pKeyAscii
            Exit Function
        End If
    End If
        
    If InStr(1, psGbn, "a") > 0 Then  'a �� ������
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
'Parameter �� �Ѿ���� Grid �� Key �� ��ġ�� Position ��Ŵ  ��
'����   : 2010.07.09
'�۾��� : �̰��
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
                
            'Key String �� Grid string �� ������
        
            If lsKeyString = lsGridString Then
                .Row = lnrows       'All ���� Matching �� ��
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
'*Author : ������ ���� �ý���
'*�ۼ��� : �̰��
'*Description:
'*  INI ���Ͽ��� �ش� Section�� Key�� �ش��ϴ� ���� �о�´�.
'*
'****************************************************************
Public Function GetValue(NewSection As String, NewKey As String, Optional NewDefault) As String
    Dim ReturnLength As Long, ReturnValue As String
    
    ReturnValue = String$(255, &H0)
    If GetPrivateProfileString(NewSection, NewKey, "", ReturnValue, Len(ReturnValue), m_sAppFile) = 0 Then
        If IsMissing(NewDefault) Then '�־��� Default���� ���� ���
            GetValue = ""
        Else '�־��� Default���� ���� ���
            GetValue = NewDefault
        End If
    Else
        GetValue = Left(ReturnValue, InStr(ReturnValue, Chr(0)) - 1)
    End If
End Function

Public Function GF_Excel_CA(plnNum As Integer) As String
'****************************************************************
'*Author : ������ ���� �ý���
'*�ۼ��� : �̰��
'*�ۼ��� : 2010.12.30
'*Description:
'*  Parameter �� ���� ���ڸ� �������� Excel ���� ����ϴ� Alphabet���� �����Ѵ�.
'*
'****************************************************************
Dim lsRtnVal        As String
Dim liMok           As String
Dim liRemain        As String
Dim stemp           As String
Dim i As Integer

    
    If plnNum > 256 Then
        MsgBox "256�̻��� ���ڴ� Excel Column �� ������ �Ұ��� �մϴ�. �����:" & plnNum, vbInformation, "[Ȯ��]"
        GoTo Err_Rtn
    End If

    '�Ѿ���� ������ 26�� ������ ���� �պκ�, �������� �ڿ� ���ڷ� �����Ͽ� Return ��
    Select Case plnNum   ' ���ڸ� ���մϴ�.
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
'*  INI ���Ͽ��� �ش� Section�� Key�� �ش��ϴ� ���� �о�´�.
'*
'****************************************************************
Public Sub SetValue(sSection As String, sKey As String, sValue As String, sFileName As String)
    Call WritePrivateProfileString(sSection, sKey, sValue, sFileName)
End Sub


Public Function GF_Get_UnitName(psUnitID As String) As String
        
    If psUnitID = "0" Then
        GF_Get_UnitName = "Yd"       '�ߵ�
    
    ElseIf psUnitID = 1 Then
        GF_Get_UnitName = "Mt"       '����
    
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


