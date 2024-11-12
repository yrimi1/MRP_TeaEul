Attribute VB_Name = "Library"
Option Explicit

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Enum EDATEFORMAT
    DF_LONG = 0
    DF_SHORT = 1
    DF_FULL = 2
End Enum

Public Enum EORDERFLAG
    OF_NONE = 0
    OF_ORDERID = 1
    OF_ORDERNO = 2
End Enum

Private Const LIB_NAME As String = "PlusLib"

Public Function GetComputer() As String
    Dim sBuffer As String
    Dim lLength As Long

    sBuffer = Space(255 + 1)
    lLength = Len(sBuffer)

    If CBool(GetComputerName(sBuffer, lLength)) Then
''        GetComputer = Left(sBuffer, lLength)
        '2012.03.06 ����-�ѱ���ǻ�͸� �߶󳻼� DB�Է½� ����
        GetComputer = MidH(sBuffer, 1, lLength)
    Else
        GetComputer = ""
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-21 (WEN)
'* UPDATE :
'*
'* ��¥ String �������� ������ iFormat�� �Ѿ�� ���� ���� �����Ѵ�.
'*   - iFormat : ���� �ε���
'*   - sDate   : ��¥ String ������
'*   = Return Value : String (����� �ؽ�Ʈ)
'********************************************************************************
Public Function MakeDate(ByVal iFormat As EDATEFORMAT, ByVal sDate As String) As String
    Dim sFmt As String

    If iFormat = DF_FULL Then
        sFmt = "YYYY�� MM�� DD��"
    ElseIf iFormat = DF_LONG Then
        sFmt = "YYYY-MM-DD"
    ElseIf iFormat = DF_SHORT Then
        sFmt = "YYYYMMDD"
    End If

    If IsDate(sDate) Then
        MakeDate = Format(sDate, sFmt)
    ElseIf Len(sDate) = 8 Then
        MakeDate = Format(Left(sDate, 4) & "-" & Mid(sDate, 5, 2) & "-" & Mid(sDate, 7), sFmt)
    Else
        MakeDate = ""
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-17 (SAT)
'* UPDATE :
'*
'* ���� String �������� ������ nCount�� �Ѿ�� ���� ���� �����Ѵ�.
'*   - sText  : �ؽ�Ʈ
'*   - nCount : �Ҽ������� �ڸ� ��
'*   = Return Value : ����� �ؽ�Ʈ
'********************************************************************************
Public Function SetCurrency(ByVal sText As String, Optional nCount As Integer = 0) As String
    Dim iCount As Integer
    Dim sBaseFmt As String

    sBaseFmt = "#,##0"
    If nCount > 0 Then
        sBaseFmt = "#,##0."
        For iCount = 0 To nCount - 1
            sBaseFmt = sBaseFmt & "0"
        Next iCount
    End If

    If IsNumeric(sText) Then
        SetCurrency = Format(sText, sBaseFmt)
    Else
        SetCurrency = "0"
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* ���ڷ� �Ѿ�� vValue�� ���� �˻��Ͽ� String���� ��ȯ�Ѵ�.
'*   - vValue : �˻��� ��
'*   = Return Value : �˻� �� ����� ��
'********************************************************************************
Public Function CheckNull(vValue As Variant) As String
    If Len(Trim(vValue)) = 0 Then
        CheckNull = "NULL"
    Else
        CheckNull = "'" & Trim(vValue) & "'"
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* ���ڷ� �Ѿ�� vValue�� ���� �������� �˻��Ͽ� ���ڷ� ��ȯ�Ѵ�.
'*   - vValue : �˻��� ��
'*   = Return Value : �˻� �� ����� ��
'********************************************************************************
Public Function CheckNum(vValue As Variant) As String
    If IsNumeric(vValue) Then
        CheckNum = CStr(vValue)
    Else
        CheckNum = "0"
    End If
End Function

Public Function ErrorSource(sClass As String, sFunction As String) As String
    ErrorSource = LIB_NAME & "." & sClass & "." & sFunction
End Function



'2012.03.06 �ű��߰�
'// TODO: �ѿ�ȥ�չ����� ���ʿ������� ������ŭ ���ڸ� �о�´�.
'// ***     MidH����
Function LeftH(str, strlen)
    Dim rValue, tmpStr, tmpASC, lenSum, f
 
  '  If isset(str) Then
        lenSum = 0
        rValue = ""
 
        For f = 1 To Len(str)
            tmpStr = Mid(str, f, 1)
            tmpASC = Asc(tmpStr)
            If tmpASC > 0 And tmpASC < 255 Then lenSum = lenSum + 1 Else lenSum = lenSum + 2
            rValue = rValue & tmpStr
            If (lenSum > strlen) Then Exit For
        Next
        LeftH = rValue
  '  End If
End Function

'2012.03.06 �ű��߰�
'//TODO: �ѿ�ȥ�չ����� �����ʿ������� ������ŭ ���ڸ� �о�´�.
'// ***     MidH����
Function RightH(str, strlen)
    Dim rValue, tmpStr, tmpASC, lenSum, f
 
   ' If isset(str) Then
        lenSum = 0
        rValue = ""
        str = StrReverse(str)
        For f = 1 To Len(str)
            tmpStr = Mid(str, f, 1)
            tmpASC = Asc(tmpStr)
            If tmpASC > 0 And tmpASC < 255 Then lenSum = lenSum + 1 Else lenSum = lenSum + 2
            rValue = rValue & tmpStr
            If (lenSum > strlen) Then Exit For
        Next
        RightH = StrReverse(rValue)
  '  End If
End Function

'2012.03.06 �ű��߰�
'//TODO: �ѿ�ȥ�չ����� start ��ġ�������� length ��ŭ ���ڸ� �о�´�.
Function MidH(s, Start, length)
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
        If BLength >= Start Then
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


