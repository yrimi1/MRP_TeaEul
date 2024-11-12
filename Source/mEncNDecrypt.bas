Attribute VB_Name = "mEncNDecrypt"
'*****************************************************************************************
' 암호처리 모듈
'-----------------------------------------------------------------------------------------
' DESC : 거래처DB의 DBConnInfo 테이블에 SQL인증의 아이디 : wizard의 암호 노출을 직접 입력하지 않기 위해
'        PassAuthCode 필으의 값을 암호화(Encrypt) 하여 넣어 두고 이 값을 가져와 복호화(Decrypt) 해서
'        위저드 서버의 우편번호 DB인 ZIPDB에 연결하게 한다.
'        우편번호 관련 도로명 주소 검색을 위한 각 거래처에서 조회를 하기 위함
'
'
'변경이력
'-----------------------------------------------------------------------------------------
' 2013.12.12  오승욱                 최초 추가
'*****************************************************************************************


Public arrEncCode As Variant               '암호화 연산자 코드 배열
Public arrDecCode As Variant               '복호화 연산사 코드 배열

'암호화 XOR연산용 데이터 초기화
Public Sub SetXorData()
    arrEncCode = Array(1, 84, 62, 23, 59, 48, 66, 11, 43, 93, 37, 50, 43, 19, 77, 29, 5, 69, 49, 21)
    
        
''    '복호화 위한 XOR 연산용 데이터 초기화
''    '실제 암호화 데이터의 역 배열순
''    ReDim arrDecCode(UBound(arrEncCode))
''    For i = UBound(arrEncCode) To 0 Step -1
''        arrDecCode(UBound(arrEncCode) - i) = arrEncCode(i)
''    Next i

''    '배열확인용-----------------------------------
''    For i = 0 To UBound(arrEncCode)
''        Debug.Print i & " : " & "Enc - " & arrEncCode(i) & ", " & "Dec - " & arrDecCode(i)
''    Next i
''    '---------------------------------------------

End Sub

'암호화 XOR연산용 데이터 배열 재선언
Public Sub SetXorDataReDim(nLen As Integer)
    Dim i As Integer
    'Preserve 기능으로 원래 배열값을 유지하면서 원본데이터 길이(글자수)만큼만 배열을 선언한다
    ReDim Preserve arrEncCode((nLen / 2) - 1)
    
    '암호와 위한 XOR 데이터인 arrEncCode 길이만큼 복호화 XOR데이터의 길이도 재 선언한다
    ReDim arrDecCode(UBound(arrEncCode))
        
    '암호화 XOR 데이터(arrEncCode)의 역 배열순으로 복호화 XOR데이터(arrDecCode)에 대입한다
    For i = UBound(arrEncCode) To 0 Step -1
        arrDecCode(UBound(arrEncCode) - i) = arrEncCode(i)
    Next i
    
End Sub

            

'원문 암호화
Function enCode(ByVal strText As String) As String
    Dim arrByte() As Byte
    Dim i As Integer
    
    Dim nEncCode As Integer
    
''    nEncCode = &H43

    'Preserve 기능으로 원래 배열값을 유지하면서 원본데이터 길이(글자수)만큼만 배열을 선언한다
    ReDim Preserve arrEncCode(Len(Trim(strText)) - 1)
    
    '암호와 위한 XOR 데이터인 arrEncCode 길이만큼 복호화 XOR데이터의 길이도 재 선언한다
    ReDim arrDecCode(UBound(arrEncCode))
    
    '암호화 XOR 데이터(arrEncCode)의 역 배열순으로 복호화 XOR데이터(arrDecCode)에 대입한다
    For i = UBound(arrEncCode) To 0 Step -1
        arrDecCode(UBound(arrEncCode) - i) = arrEncCode(i)
    Next i
    
''    '배열확인용-----------------------------------
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

        '문자열 거꾸로 붙여넣기
        enCode = Right("0" & Hex(arrByte(i) Xor "&H" & nEncCode), 2) & enCode
    Next
End Function

'다음은 암호화된 문자열을 복호화 하는 함수입니다.
'2자리 단위로 잘라서.. 숫자로 변환한 후에 다시 arrDecCode 배열의 순으로 xor 연산을 합니다.
'그런다음 바이트를 문자열로 변환합니다.

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

    '문자열 거꾸로 잘라서 붙여 넣기
    For i = 1 To Len(Trim(sTempData))
        deCode = Mid(sTempData, i, 1) & deCode
    Next i
    
End Function

