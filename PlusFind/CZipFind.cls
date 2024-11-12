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
'** System �� : MRRPLUS2-PlusFind
'** Author    : Wizard
'** �ۼ���    :
'** ����      : �ּ� �˻� Ŭ����
'** ��������  :
'** ��������  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  ��û���� ID: S_201312_���_99
'  ��û��:
'  ���泯¥ : 2013.12.12
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� : ���θ� �ּ� �Է��� ���� �߰� �� ListBox  ��� �׸��� ���, ��� ��Ʈ�� �迭�� �ۼ���
'             index 0�� ���θ� �ּ� ���� ��Ʈ��, index 1 �� �� �����ּ� �˻��� ��Ʈ��
'**************************************************************************************************
Private wZipCode As String * 7
Private wAddress As String
''Private sSearch As String

'S_201312_���_99 �� ���� �߰�----------------------------------------------------------
Private wOldNNewClss As String * 1          '���θ�,�� ���� �ּ� ����
Private wGunMoolMngNo As String * 25        '�ǹ� �����ĺ� ��ȣ
Private wAddressDetail As String            '�ǹ� ����/�ι�
Private wAddressAssist As String            '�����ּ�
Private sSearch(1) As String                '�⺻�ּ�
''Private sSearch2(1) As String           '���ּ�
'-----------------------------------------------------------------------

'�⺻�ּҰ� ����
'S_201312_���_99 �� ���� ����(OLD:Address1)
Public Property Let AddressJiBun1(NewValue As String)
    sSearch(1) = NewValue
End Property

'S_201312_���_99 �� ���� �߰�
Public Property Let Address1(NewValue As String)
    sSearch(0) = NewValue
End Property

'S_201312_���_99 �� ���� �߰�
Public Property Let Address2(NewValue As String)
    sSearch(0) = NewValue
End Property

'S_201312_���_99 �� ���� �߰�
'�ּ� �ű� ���� 0:��(�����ּ�) 1:��(���θ� �ּ�)
Public Property Let OldNNewSet(NewValue As String)
    wOldNNewClss = NewValue
End Property

'S_201312_���_99 �� ���� �߰�
Public Property Get OldNNewClss() As Variant
    OldNNewClss = wOldNNewClss
End Property

'S_201312_���_99 �� ���� �߰�
'�ǹ� �����ĺ� ��ȣ
Public Property Get GunMoolMngNo() As Variant
    GunMoolMngNo = wGunMoolMngNo
End Property

'������ȣ ��ȯ
Public Property Get ZipCode() As Variant
    ZipCode = wZipCode
End Property

'�⺻�ּ� ��ȯ
Public Property Get Address() As Variant
    Address = wAddress
End Property

'S_201312_���_99 �� ���� �߰�
'���ּ�(�ǹ� ��/�ι�) ��ȯ
Public Property Get AddressDetail() As Variant
    AddressDetail = wAddressDetail
End Property

'S_201312_���_99 �� ���� �߰�
'�����ּ� ��ȯ
Public Property Get AddressAssist() As Variant
    AddressAssist = wAddressAssist
End Property


'wAddressAssist
''------------------------------------------------------
''2002-05-15 (WEN) Shaikan
''------------------------------------------------------
Public Property Let Connection(adoConnect As Object)

    ''S_201312_���_99 �� ���� ����
''    Set adoCon = adoConnect
    Set adoWizCon = adoConnect

End Property



Public Function Show() As Boolean
    Dim FindZip As frmZip
    
    'S_201312_���_99 �� ���� �߰�
    On Error GoTo Err_Rtn
    
    Set FindZip = New frmZip
    
    'S_201312_���_99 �� ���� �߰�-���θ�,���� �ּұ���
    FindZip.g_OldNNewClss = wOldNNewClss
    
    'S_201312_���_99 �� ���� ����(Index �κ� �߰�)
    FindZip.txtName(CInt(FindZip.g_OldNNewClss)) = sSearch(CInt(FindZip.g_OldNNewClss))
    FindZip.Show vbModal
    
    If FindZip.ReturnStatus Then
''        wZipCode = Mid(FindZip.lstZip.Text, 1, 7)
''        wAddress = Trim(Mid(FindZip.lstZip.Text, 11))

        If FindZip.g_OldNNewClss = "0" Then                             '���θ� �ּ�
            wOldNNewClss = "0"                                          '���θ�,���� �ּұ���
            wZipCode = Mid(FindZip.txtZipCode(0).Text, 1, 7)            '������ȣ
            wAddress = FindZip.txtAddr(0).Text                          '�⺻�ּ�
            wAddressDetail = FindZip.txtAddr2(0).Text                   '�ǹ� ��/�ι�
            wAddressAssist = FindZip.txtAssist(0).Text            '�����ּ�
            wGunMoolMngNo = FindZip.txtGunMoolMngNo(0).Text                    '�ǹ� ���� �ĺ� ��ȣ
        Else

            wOldNNewClss = "1"                                          '���θ�,���� �ּұ���
''            wZipCode = Mid(FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("ZipCode")), 1, 7)
            wZipCode = Mid(FindZip.txtZipCode(1).Text, 1, 7)            '������ȣ
''            wAddress = FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("City")) & " "
''            wAddress = wAddress & FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("Section")) & " "
''            wAddress = wAddress & FindZip.grdZipList(1).TextMatrix(FindZip.grdZipList(1).Row, FindZip.grdZipList(1).ColIndex("Village")) & " "
            wAddress = FindZip.txtAddr(1).Text                          '�����ּ�
''            wAddressDetail = FindZip.txtAddr2(1).Text                   '���ּ�
            wAddressDetail = ""                                         '���ּ�
            wAddressAssist = ""                                          '�����ּ�
            wGunMoolMngNo = FindZip.txtGunMoolMngNo(1).Text             '�ǹ� ���� �ĺ� ��ȣ
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

