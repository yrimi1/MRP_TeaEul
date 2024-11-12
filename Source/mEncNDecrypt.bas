Attribute VB_Name = "mEncNDecrypt"
'*****************************************************************************************
' ��ȣó�� ���
'-----------------------------------------------------------------------------------------
' DESC : �ŷ�óDB�� DBConnInfo ���̺� SQL������ ���̵� : wizard�� ��ȣ ������ ���� �Է����� �ʱ� ����
'        PassAuthCode ������ ���� ��ȣȭ(Encrypt) �Ͽ� �־� �ΰ� �� ���� ������ ��ȣȭ(Decrypt) �ؼ�
'        ������ ������ �����ȣ DB�� ZIPDB�� �����ϰ� �Ѵ�.
'        �����ȣ ���� ���θ� �ּ� �˻��� ���� �� �ŷ�ó���� ��ȸ�� �ϱ� ����
'
'
'�����̷�
'-----------------------------------------------------------------------------------------
' 2013.12.12  ���¿�                 ���� �߰�
'*****************************************************************************************


Public arrEncCode As Variant               '��ȣȭ ������ �ڵ� �迭
Public arrDecCode As Variant               '��ȣȭ ����� �ڵ� �迭

'��ȣȭ XOR����� ������ �ʱ�ȭ
Public Sub SetXorData()
    arrEncCode = Array(1, 84, 62, 23, 59, 48, 66, 11, 43, 93, 37, 50, 43, 19, 77, 29, 5, 69, 49, 21)
    
        
''    '��ȣȭ ���� XOR ����� ������ �ʱ�ȭ
''    '���� ��ȣȭ �������� �� �迭��
''    ReDim arrDecCode(UBound(arrEncCode))
''    For i = UBound(arrEncCode) To 0 Step -1
''        arrDecCode(UBound(arrEncCode) - i) = arrEncCode(i)
''    Next i

''    '�迭Ȯ�ο�-----------------------------------
''    For i = 0 To UBound(arrEncCode)
''        Debug.Print i & " : " & "Enc - " & arrEncCode(i) & ", " & "Dec - " & arrDecCode(i)
''    Next i
''    '---------------------------------------------

End Sub

'��ȣȭ XOR����� ������ �迭 �缱��
Public Sub SetXorDataReDim(nLen As Integer)
    Dim i As Integer
    'Preserve ������� ���� �迭���� �����ϸ鼭 ���������� ����(���ڼ�)��ŭ�� �迭�� �����Ѵ�
    ReDim Preserve arrEncCode((nLen / 2) - 1)
    
    '��ȣ�� ���� XOR �������� arrEncCode ���̸�ŭ ��ȣȭ XOR�������� ���̵� �� �����Ѵ�
    ReDim arrDecCode(UBound(arrEncCode))
        
    '��ȣȭ XOR ������(arrEncCode)�� �� �迭������ ��ȣȭ XOR������(arrDecCode)�� �����Ѵ�
    For i = UBound(arrEncCode) To 0 Step -1
        arrDecCode(UBound(arrEncCode) - i) = arrEncCode(i)
    Next i
    
End Sub

            

'���� ��ȣȭ
Function enCode(ByVal strText As String) As String
    Dim arrByte() As Byte
    Dim i As Integer
    
    Dim nEncCode As Integer
    
''    nEncCode = &H43

    'Preserve ������� ���� �迭���� �����ϸ鼭 ���������� ����(���ڼ�)��ŭ�� �迭�� �����Ѵ�
    ReDim Preserve arrEncCode(Len(Trim(strText)) - 1)
    
    '��ȣ�� ���� XOR �������� arrEncCode ���̸�ŭ ��ȣȭ XOR�������� ���̵� �� �����Ѵ�
    ReDim arrDecCode(UBound(arrEncCode))
    
    '��ȣȭ XOR ������(arrEncCode)�� �� �迭������ ��ȣȭ XOR������(arrDecCode)�� �����Ѵ�
    For i = UBound(arrEncCode) To 0 Step -1
        arrDecCode(UBound(arrEncCode) - i) = arrEncCode(i)
    Next i
    
''    '�迭Ȯ�ο�-----------------------------------
''    For i = 0 To UBound(arrEncCode)
''        Debug.Print i & " : " & "Enc - " & arrEncCode(i) & ", " & "Dec - " & arrDecCode(i)
''    Next i
''    '-------------------------------------------------
    arrByte = StrConv(strText, vbFromUnicode)
    For i = 0 To UBound(arrByte)

        nEncCode = arrEncCode(i)
''        Debug.Print nEncCode
''        enCode = enCode & Right("0" & Hex(arrByte(i) Xor &H43), 2)
''        enCode = enCode & Right("0" & Hex(arrByte(i) Xor "&H" & nEncCode), 2)

        '���ڿ� �Ųٷ� �ٿ��ֱ�
        enCode = Right("0" & Hex(arrByte(i) Xor "&H" & nEncCode), 2) & enCode
    Next
End Function

'������ ��ȣȭ�� ���ڿ��� ��ȣȭ �ϴ� �Լ��Դϴ�.
'2�ڸ� ������ �߶�.. ���ڷ� ��ȯ�� �Ŀ� �ٽ� arrDecCode �迭�� ������ xor ������ �մϴ�.
'�׷����� ����Ʈ�� ���ڿ��� ��ȯ�մϴ�.

Function deCode(ByVal strText As String) As String
    Dim arrByte() As Byte
    Dim i As Integer
    Dim j As Integer
    Dim nDecCode As Integer
    
    ReDim arrByte((Len(strText) / 2) - 1)
    For i = 1 To Len(strText) Step 2

        nDecCode = arrDecCode(j)
''        arrByte(i \ 2) = Val("&H" & Mid(strText, i, 2)) Xor &H43
        arrByte(i \ 2) = val("&H" & Mid(strText, i, 2)) Xor "&H" & nDecCode

''        Debug.Print nDecCode
        j = j + 1
    Next
    deCode = StrConv(arrByte, vbUnicode)
    
    Dim sOneChar As String
    Dim sTempData As String
    
    sTempData = deCode
    deCode = ""

    '���ڿ� �Ųٷ� �߶� �ٿ� �ֱ�
    For i = 1 To Len(Trim(sTempData))
        deCode = Mid(sTempData, i, 1) & deCode
    Next i
    
End Function

