VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************
'** System �� : Mrpplus2
'** ����    :
'** Author    : Wizard
'** �ۼ���    :
'** ����      :
'** ��������  :
'------------------------------------------------------------------------------------------------------------------
' * �����̷�
'------------------------------------------------------------------------------------------------------------------
' ����        �۾���  ��û��          ��û��ȣ           ��û���� �� ���泻��
'------------------------------------------------------------------------------------------------------------------
' 2013.12.12  ���¿�                 S_201312_���_99    �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'*******************************************************************************

Option Explicit

Type TInfo
    sInfoDate As String
    sInfo     As String
    'InfoUser As String
End Type

Type TInfoUser
    sInfoDate   As String
    nInfoSeq    As Integer
    sInfoUser   As String
End Type

Private Const CLASS_NAME As String = "CInfo"


'S_201312_���_99 �� ���� �߰�-�ڻ�����-��ϼ�����
Type TCompanyInfo2
    sCompanyID      As String            '�ڻ�-�ڵ�
    sKCompany       As String           '�ڻ�-�ѱ۸�
    sECompany       As String           '�ڻ�-������
    sShortCompany   As String           '�ڻ�-���
    sChief          As String           '�ڻ�-��ǥ��
    sCompanyNo      As String            '�ڻ�-����ڹ�ȣ
    sRegistID       As String           '�ڻ�-�ֹε�Ϲ�ȣ
    sCondition      As String           '�ڻ�-����
    sCategory       As String           '�ڻ�-����
    sZipCode        As String           '�ڻ�-������ȣ
    sOldNNewClss        As String       '�ڻ�-���θ�,�����ּ� ���� 0:���θ�, 1:����
    sGunMoolMngNo        As String      '�ڻ�-�ǹ����� �����ĺ���ȣ
    sAddress1       As String        '�ڻ�-���θ��ּ�1
    sAddress2       As String        '�ڻ�-���θ��ּ�2
    sAddressAssist       As String      '�ڻ�-���θ������ּ�
    sAddressJiBun1       As String           '�ڻ�-�ּ�1
    sAddressJiBun2       As String           '�ڻ�-�ּ�2
    sPhone1          As String           '�ڻ�-��ȭ��ȣ1
    sPhone2         As String           '�ڻ�-��ȭ��ȣ2
    sFaxNO          As String           '�ڻ�-�ѽ���ȣ
    sEMail          As String           '�ڻ�-�̸���
    sHomePage       As String           '�ڻ�-Ȩ������
    sBank1          As String           '�ڻ�-���¹�ȣ1
    sBank2          As String           '�ڻ�-���¹�ȣ2
    sBank3          As String           '�ڻ�-���¹�ȣ3
    sRPYn           As String           '��ǥ����
    sRPYn_OLD       As String           '����������
    sUseYn          As String           '��뿩��
    
    '�߰�����*********************************************************
    '
    ' --WebPage�α�������
    sWebPortFrom                         As String   'WebPage��ƮFrom
    sWebPortTo                           As String   'WebPage��ƮTo
    sWebID1                              As String   'WebPage�α���ID1
    sWebPass1                            As String   'WebPage�α��ξ�ȣ1
    sWebAuthCode1                        As String   'WebPage�α��������ڵ�1
    sWebID2                              As String   'WebPage�α���ID2
    sWebPass2                            As String   'WebPage�α��ξ�ȣ2
    sWebAuthCode2                        As String   'WebPage�α��������ڵ�2
    
    ' --FTP�α�������
    sFTPPage                             As String   'FTP�ּ�
    sFTPPortFrom                         As String   'FTP��ƮFrom
    sFTPPortTo                           As String   'FTP��ƮTo
    sFTPID1                              As String   'FTP�α���ID1
    sFTPPass1                            As String   'FTP�α��ξ�ȣ1
    sFTPAuthCode1                        As String   'FTP�α��������ڵ�1
    sFTPID2                              As String   'FTP�α���ID2
    sFTPPass2                            As String   'FTP�α��ξ�ȣ2
    sFTPAuthCode2                        As String   'FTP�α��������ڵ�2
    
    ' --SMS����1��������
    sSMSURL1                             As String   '�������ۼ���1�ּ�
    sSMSPortFrom1                        As String   '�������ۼ���1��ƮFrom
    sSMSPortTo1                          As String   '�������ۼ���1��ƮTo
    sSMSID1                              As String   '�������ۼ���1���̵�
    sSMSPASS1                            As String   '�������ۼ���1��ȣ
    sSMSAuthCode1                        As String   '�������ۼ���1�����ڵ�

    ' --SMS����2��������
    sSMSURL2                             As String   '�������ۼ���2�ּ�
    sSMSPortFrom2                        As String   '�������ۼ���2��ƮFrom
    sSMSPortTo2                          As String   '�������ۼ���2��ƮTo
    sSMSID2                              As String   '�������ۼ���2���̵�
    sSMSPASS2                            As String   '�������ۼ���2��ȣ
    sSMSAuthCode2                        As String   '�������ۼ���2�����ڵ�
    '*****************************************************************
End Type


'S_201312_���_99 �� ����  �߰�----------------------------
Private m_nFileLength As Long
Private m_nNumBlocks As Integer
Private m_nLeftOver As Long
Private m_sByteData() As Byte
Private Const m_nBlockSize = 4096
'--------------------------------------------------------------

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetNewInfoSeq(sTable As String, sField As String, Optional sWhere As String = "") As Long
    GetNewInfoSeq = GetMaxValue(sTable, sField, sWhere)
End Function

Public Function GetInfoByDate(sSDate As String, sEDate As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_sInfoByDate"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)

        Set GetInfoByDate = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInforByDate"), Err.Description
End Function

' ���ΰ������ ���̺��� �ش糯¥�� �����Ͱ� �ִ��� Ȯ��..
Public Function CheckDate(sDate As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_sCheckDate"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        
        Set CheckDate = .Execute
    End With
    
    Set adoCmd = Nothing
    
End Function



' �ش� ��¥�� ���κ� �������� ���...
Public Function GetPersonInfoList(sDate As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
        
    On Error GoTo ErrorHandler
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_sInfoUserList"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        
        Set GetPersonInfoList = .Execute
    End With
    
    Exit Function
    
ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "CInfo.GetPersonInfoList"
    
End Function


' �ش� ���κ� �������׿� �´� ����� ID �� �̸�
Public Function GetPersonInfoID(sDate As String, sNum As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    
    On Error GoTo ErrorHandler
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_sInfoUserID"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, sNum)
        
        Set GetPersonInfoID = .Execute
    End With
    
    Exit Function
    
ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "CInfo.GetPersonInfoID"

End Function



Public Function GetInfoUserByUserID(sDate As String, sUserID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrorHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_sInfoUserByUserID"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sUserID)

        Set GetInfoUserByUserID = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function
    
ErrorHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInfoUserByUserID"), Err.Description
End Function



Public Function AddInfo(NewInfo As TInfo, NewInfoUser As TInfoUser, PersonID() As String, nSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim i%
    
    On Error GoTo ErrHandler
    
    AddInfo = True
    
    Set adoCmd = New ADODB.Command
    With adoCmd     ' ��ü ��������
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_iInfo"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, NewInfo.sInfoDate)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 3000, Left(NewInfo.sInfo, 3000))
        
        .Execute
    End With
    
    Call ClearParameter(adoCmd)
    
    With adoCmd     ' ����ں� ��������
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_iInfoUser"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, NewInfoUser.sInfoDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, NewInfoUser.nInfoSeq)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 1000, NewInfoUser.sInfoUser)
        
        .Execute
    End With
    
    Call ClearParameter(adoCmd)
    
    With adoCmd     ' ����ں� �������� ���� ID ��� ����
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_dInfoUserList"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, NewInfoUser.sInfoDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, NewInfoUser.nInfoSeq)
        
        .Execute
    End With
    
    Call ClearParameter(adoCmd)
    
    With adoCmd     ' ����ں� �������� ����(PersonID) �Է�
        .CommandText = "xp_Info_iInfoUserList"
        .Prepared = True
    
        For i = 0 To nSeq
            Call ClearParameter(adoCmd)
    
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, NewInfoUser.sInfoDate)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, NewInfoUser.nInfoSeq)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, PersonID(i))
        
            .Execute
        Next i
    End With
            
    Set adoCmd = Nothing
    
Exit Function

ErrHandler:
    Set adoCmd = Nothing
    AddInfo = False
    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddInfo"), Err.Description

End Function


'S_201312_���_99 �� ����  �߰�
Public Function SaveCompanyInfo(TItem As TCompanyInfo2) As Boolean
    Dim SQL() As String

    ReDim SQL(0)

  
    With TItem
        SQL(0) = " UPDATE [mt_SetCompany]                            " & vbCrLf
        SQL(0) = SQL(0) & " SET KCompany    ='" & .sKCompany & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " ECompany        ='" & .sECompany & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " ShortCompany    ='" & .sShortCompany & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " Chief           ='" & .sChief & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " CompanyNo       ='" & .sCompanyNo & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " RegistID       ='" & .sRegistID & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " Condition       ='" & .sCondition & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Category        ='" & .sCategory & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " ZipCode         ='" & .sZipCode & "' ,   " & vbCrLf
        
        'S_201312_���_99 �� ���� �߰�-------------------------------------------------------------
        SQL(0) = SQL(0) & " OldNNewClss     ='" & .sOldNNewClss & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " GunMoolMngNo    ='" & .sGunMoolMngNo & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Address1     ='" & .sAddress1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Address2     ='" & .sAddress2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " AddressAssist   ='" & .sAddressAssist & "' ,   " & vbCrLf
        '------------------------------------------------------------------------------------
        SQL(0) = SQL(0) & " AddressJiBun1        ='" & .sAddressJiBun1 & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " AddressJiBun2        ='" & .sAddressJiBun2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Phone1          ='" & .sPhone1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Phone2          ='" & .sPhone2 & "' ,  " & vbCrLf
        SQL(0) = SQL(0) & " FaxNo           ='" & .sFaxNO & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " EMail           ='" & .sEMail & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " HomePage        ='" & .sHomePage & "' ,   " & vbCrLf
   
        'Web �߰�����*********************************************************
        ' --WebPage�α�������
        SQL(0) = SQL(0) & " WebPortFrom     ='" & .sWebPortFrom & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebPortTo       ='" & .sWebPortTo & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebID1          ='" & .sWebID1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebPass1        ='" & .sWebPass1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebAuthCode1    ='" & .sWebAuthCode1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebID2          ='" & .sWebID2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebPass2        ='" & .sWebPass2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " WebAuthCode2    ='" & .sWebAuthCode2 & "' ,   " & vbCrLf
        ' --FTP�α�������
        SQL(0) = SQL(0) & " FTPPage         ='" & .sFTPPage & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPPortFrom     ='" & .sFTPPortFrom & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPPortTo       ='" & .sFTPPortTo & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPID1          ='" & .sFTPID1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPPass1        ='" & .sFTPPass1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPAuthCode1    ='" & .sFTPAuthCode1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPID2          ='" & .sFTPID2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPPass2        ='" & .sFTPPass2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " FTPAuthCode2    ='" & .sFTPAuthCode2 & "' ,   " & vbCrLf
        ' --SMS����1��������
        SQL(0) = SQL(0) & " SMSURL1         ='" & .sSMSURL1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSPortFrom1    ='" & .sSMSPortFrom1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSPortTo1      ='" & .sSMSPortTo1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSID1          ='" & .sSMSID1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSPASS1        ='" & .sSMSPASS1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSAuthCode1    ='" & .sSMSAuthCode1 & "' ,   " & vbCrLf
        ' --SMS����2��������
        SQL(0) = SQL(0) & " SMSURL2         ='" & .sSMSURL2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSPortFrom2    ='" & .sSMSPortFrom2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSPortTo2      ='" & .sSMSPortTo2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSID2          ='" & .sSMSID2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSPASS2        ='" & .sSMSPASS2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " SMSAuthCode2    ='" & .sSMSAuthCode2 & "' ,   " & vbCrLf
        '****************************************************************
        
        SQL(0) = SQL(0) & " Bank1           ='" & .sBank1 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Bank2           ='" & .sBank2 & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " Bank3           ='" & .sBank3 & "' ,   " & vbCrLf
''        SQL(0) = SQL(0) & " RPYn           ='" & .sRPYn & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " UseYn           ='" & .sUseYn & "' ,   " & vbCrLf
        SQL(0) = SQL(0) & " UpdateDate      =dbo.fn_TO_CHAR_DATETIME( GETDATE(), 'YYYY-MM-DD HH:MI:SS' ) ,   " & vbCrLf
        SQL(0) = SQL(0) & " UpdateUserID    ='" & g_sUserName & "'  " & vbCrLf
        SQL(0) = SQL(0) & " WHERE CompanyID= '" & TItem.sCompanyID & "' "

        If .sRPYn = "Y" Then
            ReDim Preserve SQL(2)
            SQL(1) = " UPDATE [mt_SetCompany]                            " & vbCrLf
            SQL(1) = SQL(1) & " SET RPYn    ='Y',            " & vbCrLf
            SQL(1) = SQL(1) & " UpdateDate      =dbo.fn_TO_CHAR_DATETIME( GETDATE(), 'YYYY-MM-DD HH:MI:SS' ) ,   " & vbCrLf
            SQL(1) = SQL(1) & " UpdateUserID    ='" & g_sUserName & "'  " & vbCrLf
            SQL(1) = SQL(1) & " WHERE CompanyID= '" & TItem.sCompanyID & "' "
            
            SQL(2) = " UPDATE [mt_SetCompany]                            " & vbCrLf
            SQL(2) = SQL(2) & " SET RPYn    ='N' ,           " & vbCrLf
            SQL(2) = SQL(2) & " UpdateDate      =dbo.fn_TO_CHAR_DATETIME( GETDATE(), 'YYYY-MM-DD HH:MI:SS' ) ,   " & vbCrLf
            SQL(2) = SQL(2) & " UpdateUserID    ='" & g_sUserName & "'  " & vbCrLf
            SQL(2) = SQL(2) & " WHERE CompanyID Not IN ( '" & TItem.sCompanyID & "') "

            
        End If

    End With

    SaveCompanyInfo = HandleDBMulti(SQL)
End Function

'S_201312_���_99 �� ����  �߰�
'�ڻ� ���� ��������
Public Function GetInfo(Optional nChkCompany As Integer, Optional sCompanyID As String, Optional sCompany As String) As Recordset
    Dim adoCmd As ADODB.Command
    On Error GoTo ErrHandler
    
    
    Set adoCmd = New ADODB.Command
    
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Info_GetCompanyInfo"
 
        'S_201110_�����ؽ�_03 �� ���� �߰�
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCompany)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCompanyID)      '�ڻ��ڵ�
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sCompany)    '�ڻ��ȣ
        
      
        Set GetInfo = .Execute
    End With
    
    Set adoCmd = Nothing
    Exit Function
    
ErrHandler:
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CInfo.GetInfo"
    Set adoCmd = Nothing
    
End Function

'S_201312_���_99 �� ����  �߰�
Public Function GetCompanyLogo(sPath As String, sLogoName As String) As Boolean
    Dim i%
    Dim rs As ADODB.Recordset
    Dim DestFile As Integer
    
    On Error GoTo ErrGetCompanyLogo:
    
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT LogoSize, Logo FROM [mt_Info]", g_adoCon, adOpenKeyset, adLockOptimistic

    m_nFileLength = rs!LogoSize
    DestFile = FreeFile
    
    Open sPath & sLogoName For Binary As DestFile
    
    
    m_nNumBlocks = m_nFileLength / m_nBlockSize
    m_nLeftOver = m_nFileLength Mod m_nBlockSize
    
    m_sByteData() = rs!Logo.GetChunk(m_nLeftOver)
    Put DestFile, , m_sByteData()
    
    For i = 1 To m_nNumBlocks
        m_sByteData() = rs!Logo.GetChunk(m_nBlockSize)
        Put DestFile, , m_sByteData()
    Next i
    rs.Close
    
    Close DestFile
    
    Set rs = Nothing
    
    GetCompanyLogo = True
    Exit Function
ErrGetCompanyLogo:
    Set rs = Nothing
    GetCompanyLogo = False
End Function

