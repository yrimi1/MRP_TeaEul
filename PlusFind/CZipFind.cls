VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZipFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
'** System 명 : MRRPLUS2-PlusFind
'** Author    : Wizard
'** 작성자    :
'** 내용      : 주소 검색 클래스
'** 생성일자  :
'** 변경일자  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  요청사항 ID: S_201312_삼우_99
'  요청자:
'  변경날짜 : 2013.12.12
'  작업자   : 오승욱
'  요청내용 : 지번주소에서 도로명 주소로 입력가능하게
'  변경내용 : 도로명 주소 입력을 위한 추가 및 ListBox  대신 그리드 사용, 모든 컨트롤 배열로 작성함
'             index 0은 도로명 주소 관련 컨트롤, index 1 은 구 지번주소 검색용 컨트롤
'**************************************************************************************************
Private wZipCode As String * 7
Private wAddress As String
''Private sSearch As String

'S_201312_삼우_99 에 의한 추가----------------------------------------------------------
Private wOldNNewClss As String * 1          '도로명,구 지번 주소 구분
Private wGunMoolMngNo As String * 25        '건물 고유식별 번호
Private wAddressDetail As String            '건물 본번/부번
Private wAddressAssist As String            '보조주소
Private sSearch(1) As String                '기본주소
''Private sSearch2(1) As String           '상세주소
'-----------------------------------------------------------------------

'기본주소값 셋팅
'S_201312_삼우_99 에 의한 수정(OLD:Address1)
Public Property Let AddressJiBun1(NewValue As String)
    sSearch(1) = NewValue
End Property

'S_201312_삼우_99 에 의한 추가
Public Property Let Address1(NewValue As String)
    sSearch(0) = NewValue
End Property

'S_201312_삼우_99 에 의한 추가
Public Property Let Address2(NewValue As String)
    sSearch(0) = NewValue
End Property

'S_201312_삼우_99 에 의한 추가
'주소 신규 구분 0:구(지번주소) 1:신(도로명 주소)
Public Property Let OldNNewSet(NewValue As String)
    wOldNNewClss = NewValue
End Property

'S_201312_삼우_99 에 의한 추가
Public Property Get OldNNewClss() As Variant
    OldNNewClss = wOldNNewClss
End Property

'S_201312_삼우_99 에 의한 추가
'건물 고유식별 번호
Public Property Get GunMoolMngNo() As Variant
    GunMoolMngNo = wGunMoolMngNo
End Property

'우편번호 반환
Public Property Get ZipCode() As Variant
    ZipCode = wZipCode
End Property

'기본주소 반환
Public Property Get Address() As Variant
    Address = wAddress
End Property

'S_201312_삼우_99 에 의한 추가
'상세주소(건물 본/부번) 반환
Public Property Get AddressDetail() As Variant
    AddressDetail = wAddressDetail
End Property

'S_201312_삼우_99 에 의한 추가
'보조주소 반환
Public Property Get AddressAssist() As Variant
    AddressAssist = wAddressAssist
End Property


'wAddressAssist
''------------------------------------------------------
''2002-05-15 (WEN) Shaikan
''------------------------------------------------------
Public Property Let Connection(adoConnect As Object)

    ''S_201312_삼우_99 에 의한 수정
''    Set adoCon = adoConnect
    Set adoWizCon = adoConnect

End Property



Public Function Show() As Boolean
    Dim FindZip As frmZip
    
    'S_201312_삼우_99 에 의한 추가
    On Error GoTo Err_Rtn
    
    Set FindZip = New frmZip
    
    'S_201312_삼우_99 에 의한 추가-도로명,지번 주소구분
    FindZip.g_OldNNewClss = wOldNNewClss
    
    'S_201312_삼우_99 에 의한 수정(Index 부분 추가)
    FindZip.txtName(CInt(FindZip.g_OldNNewClss)) = sSearch(CInt(FindZip.g_OldNNewClss))
    FindZip.Show vbModal
    
    If FindZip.ReturnStatus Then
''        wZipCode = Mid(FindZip.lstZip.Text, 1, 7)
''        wAddress = Trim(Mid(FindZip.lstZip.Text, 11))

        If FindZip.g_OldNNewClss = "0" Then                             '도로명 주소
            wOldNNewClss = "0"                                          '도로명,지번 주소구분
            wZipCode = Mid(FindZip.txtZipCode(0).Text, 1, 7)            '우편번호
            wAddress = FindZip.txtAddr(0).Text                          '기본주소
            wAddressDetail = FindZip.txtAddr2(0).Text                   '건물 본/부번
            wAddressAssist = FindZip.txtAssist(0).Text            '보조주소
            wGunMoolMngNo = FindZip.txtGunMoolMngNo(0).Text                    '건물 고유 식별 번호
        Else

            wOldNNewClss = "1"                                          '도로명,지번 주소구분
''            wZipCode = Mid(FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("ZipCode")), 1, 7)
            wZipCode = Mid(FindZip.txtZipCode(1).Text, 1, 7)            '우편번호
''            wAddress = FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("City")) & " "
''            wAddress = wAddress & FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("Section")) & " "
''            wAddress = wAddress & FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("Village")) & " "
            wAddress = FindZip.txtAddr(1).Text                          '기존주소
''            wAddressDetail = FindZip.txtAddr2(1).Text                   '상세주소
            wAddressDetail = ""                                         '상세주소
            wAddressAssist = ""                                          '보조주소
            wGunMoolMngNo = FindZip.txtGunMoolMngNo(1).Text             '건물 고유 식별 번호
        End If
        
        Show = True
    Else
        Show = False
    End If
        
    Unload FindZip
    Set FindZip = Nothing
    
    Exit Function
    
Err_Rtn:
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CZipFind.Show"
End Function


