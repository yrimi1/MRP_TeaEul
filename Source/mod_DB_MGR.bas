Attribute VB_Name = "mod_DB_MGR"

'**************************************************************************************************
'** System �� : MRRPLUS2
'** Author    : Wizard
'** �ۼ���    : �̰��
'** ����      : DB�� Data ������ Business Logic ó���� ���� Module �̴�
'** ��������  : 2012.03.06
'**------------------------------------------------------------------------------------------------
'** ��û����    ��û��          ��ûID                  ��������        ������      ��û �� ���泻��
' 2013.12.10  ���¿�                 S_201312_��������_99    �����ּҿ��� ���θ�
'**************************************************************************************************
Option Explicit
 
'S_201312_��������_99 �� ���� �߰�-�ٸ� ��ü�� ���߱� ����
Public Function Gf_DB_CM_S_CustomOne(prs As ADODB.Recordset, psCustomid As String) As Boolean
    Dim lssql                           As String
    On Error GoTo Err_Rtn
    '�����δ��� �ҽ�

    lssql = ""
    lssql = lssql & "  SELECT Custom_ID=MC.CustomID                                                  " & vbCrLf
    lssql = lssql & "       , Custom_Nat=MC.KCustom                                                  " & vbCrLf
    lssql = lssql & "       , Custom_Short=MC.KCustom                                                " & vbCrLf
    lssql = lssql & "       , Custom_ENG=MC.ECustom                                                  " & vbCrLf
    lssql = lssql & "       , Custom_JPN=''                                                          " & vbCrLf
    lssql = lssql & "       , Custom_Gbn=MC.TradeID                                                  " & vbCrLf
    lssql = lssql & "       , Custom_Gbn_Name = dbo.fn_cm_sCodeInfo('CSG', MC.TradeID, 'N' )         " & vbCrLf
    lssql = lssql & "       , Custom_Main_JOB=''                                                     " & vbCrLf
    lssql = lssql & "       , Custom_Main_JOB_Name =''                                               " & vbCrLf
    lssql = lssql & "       , Custom_NO=MC.CustomNO                                                  " & vbCrLf
    lssql = lssql & "       , Chief                                                                  " & vbCrLf
    lssql = lssql & "       , Job_Category=MC.Category                                               " & vbCrLf
    lssql = lssql & "       , Job_Type=MC.Condition                                                  " & vbCrLf
    lssql = lssql & "       , ZipCode                                                                " & vbCrLf
    lssql = lssql & "       , OldNNewClss                                                            " & vbCrLf
    lssql = lssql & "       , GunMoolMngNo                                                           " & vbCrLf
    lssql = lssql & "       , Address1                                                               " & vbCrLf
    lssql = lssql & "       , Address2                                                               " & vbCrLf
    lssql = lssql & "       , AddressAssist                                                          " & vbCrLf
    lssql = lssql & "       , AddressJiBun1                                                          " & vbCrLf
    lssql = lssql & "       , AddressJiBun2                                                          " & vbCrLf
    lssql = lssql & "       , Phone1                                                                 " & vbCrLf
    lssql = lssql & "       , Phone2                                                                 " & vbCrLf
    lssql = lssql & "       , FaxNo                                                                  " & vbCrLf
    lssql = lssql & "       , Email                                                                  " & vbCrLf
    lssql = lssql & "       , HomePage                                                               " & vbCrLf
    lssql = lssql & "       , Custom_Charge= ''                                                      " & vbCrLf
    lssql = lssql & "       , Custom_Charge_Phone=''                                                 " & vbCrLf
    lssql = lssql & "       , USE_YN=(CASE USECLSS WHEN '*' THEN 'Y' ELSE 'N' END)                   " & vbCrLf
    lssql = lssql & "       , Business_charge                                                        " & vbCrLf
    lssql = lssql & "       , Business_charge_Name = dbo.fn_cm_sUserName ('P',Business_charge)       " & vbCrLf
    lssql = lssql & "       , Comments                                                               " & vbCrLf
    lssql = lssql & "       , PaymentCondition                                                       " & vbCrLf
    lssql = lssql & "       , Create_date=MC.Setdate                                                 " & vbCrLf
    lssql = lssql & "       , Create_user_ID=''                                                      " & vbCrLf
    lssql = lssql & "       , Update_date=MC.Setdate                                                 " & vbCrLf
    lssql = lssql & "       , Update_user_ID=''                                                      " & vbCrLf
    lssql = lssql & "    FROM MT_Custom     MC                                                       " & vbCrLf
    lssql = lssql & "   WHERE MC.CustomID                              =  '" & psCustomid & "'       " & vbCrLf

    If Gf_DB_OpenRS(g_adoCon, prs, lssql) = False Then GoTo Err_Rtn
    Gf_DB_CM_S_CustomOne = True
'    case B_S_ITEM when 'B0000' then isnull(article,'')
    Exit Function
Err_Rtn:

    MsgBox "������ Select �� ���� �߻��߽��ϴ�!!" & vbCrLf & _
            Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_S_CustomOne]"
    Call Gs_DB_CloseRs(prs)
End Function


'================================================================
'*  ��ü Setting ������ �����´�.
'*  ��������: 2012.03.26
'*  ������  : �̰��
'*  Parameter  :
'*   pRepresentYN : ��ǥ ����� ����
'*   psCompany_No : ����ڹ�ȣ
'*---------------------------------------------------------------
'*  �����̷�:
'*---------------------------------------------------------------
'*  ��������    ������  ���泻��
'================================================================
Public Function Gf_DB_CM_GetCompanyInfo(prs As ADODB.Recordset, pRepresentYN As String, Optional psCompany_No As String) As Boolean
    Dim lssql                           As String
    Dim rs                              As ADODB.Recordset
    On Error GoTo Err_Rtn
    
    lssql = ""

    lssql = lssql & "  SELECT TOP 1 Company_ID=CompanyID      " & vbCrLf 'ȸ�� ID
    lssql = lssql & "       , company_No=CompanyNo      " & vbCrLf '����ڵ�Ϲ�ȣ
    lssql = lssql & "       , company_Name=KCompany     " & vbCrLf '��ȣ
    lssql = lssql & "       , Chief                     " & vbCrLf '�����̸�
    lssql = lssql & "       , company_type=Condition    " & vbCrLf '����
    lssql = lssql & "       , Category                  " & vbCrLf '����
    lssql = lssql & "       , Zip_Code=ZipCode          " & vbCrLf '�����ȣ
    'S_201312_��������_99 �� ���� �߰�-------------------------------------------------------
    lssql = lssql & "       , OldNNewClss               " & vbCrLf '�ּұ���(0:���θ�,1:�����ּ�)
    lssql = lssql & "       , GunMoolMngNo              " & vbCrLf '�ǹ������ĺ���ȣ
    lssql = lssql & "       , Address1                  " & vbCrLf '���θ��ּ�1
    lssql = lssql & "       , Address2                  " & vbCrLf '���θ��ּ�2
    lssql = lssql & "       , AddressAssist             " & vbCrLf '���θ� ���� �ּ�
    '--------------------------------------------------------------------------------------
    'S_201312_��������_99 �� ���� ����(OLD:Address1)
    lssql = lssql & "       , AddressJiBun1                  " & vbCrLf '�����ּ�1
    'S_201312_��������_99 �� ���� ����(OLD:Address2)
    lssql = lssql & "       , AddressJiBun2                  " & vbCrLf '�����ּ�2
    lssql = lssql & "       , Phone  = Phone1           " & vbCrLf '��ȭ��ȣ
    lssql = lssql & "       , Phone2                    " & vbCrLf '��ȭ��ȣ2
    lssql = lssql & "       , FaxNO                     " & vbCrLf '�ѽ���ȣ
    lssql = lssql & "       , Bank1                     " & vbCrLf '����1
    lssql = lssql & "       , Bank2                     " & vbCrLf '����2
    lssql = lssql & "       , Bank3                     " & vbCrLf '����3
    
        '2013.12.12 �߰�
    '�ڻ� Web���� �߰� ����
    lssql = lssql & "       , WebPortFrom, WebPortTo            " & vbCrLf '����ƮFrom,To
    lssql = lssql & "       , WebID1, WebPass1, WebAuthCode1    " & vbCrLf 'Web���̵�1,��ȣ1,�����ڵ�1
    lssql = lssql & "       , WebID2, WebPass2, WebAuthCode2    " & vbCrLf 'Web���̵�2,��ȣ2,�����ڵ�2
    lssql = lssql & "       , FTPPage, FTPPortFrom ,FTPPortTo   " & vbCrLf 'FTP�ּ�,��ƮFrom,To
    lssql = lssql & "       , FTPID1, FTPPass1, FTPAuthCode1    " & vbCrLf 'FTP���̵�1,��ȣ1,�����ڵ�1
    lssql = lssql & "       , FTPID2, FTPPass2, FTPAuthCode2    " & vbCrLf 'FTP���̵�2,��ȣ2,�����ڵ�2
    lssql = lssql & "       , SMSURL1, SMSPortFrom1, SMSPortTo1 " & vbCrLf 'SMS����1�ּ�,��ƮFrom,To
    lssql = lssql & "       , SMSID1, SMSPASS1, SMSAuthCode1    " & vbCrLf 'SMS����1���̵�,��ȣ,�����ڵ�
    lssql = lssql & "       , SMSURL2, SMSPortFrom2, SMSPortTo2 " & vbCrLf 'SMS����2�ּ�,��ƮFrom,To
    lssql = lssql & "       , SMSID2, SMSPASS2, SMSAuthCode2    " & vbCrLf 'SMS����2���̵�,��ȣ,�����ڵ�

    
    
    lssql = lssql & "    FROM mt_setCompany             " & vbCrLf
    lssql = lssql & "   WHERE 1= 1                      " & vbCrLf
    
    '��ǥ����
    If pRepresentYN <> "" Then
        lssql = lssql & "   AND  RPYN='" & pRepresentYN & "'                      " & vbCrLf
    End If

    lssql = lssql & " order by  RPYN Desc, company_Name            " & vbCrLf
    
    If Gf_DB_OpenRS(g_adoCon, prs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DB_CM_GetCompanyInfo = True
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then
        MsgBox "�ý��ۻ���ü �⺻���� Select �� ���� �߻��߽��ϴ�!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetCompanyInfo]"
    End If
    Call Gs_DB_CloseRs(prs)
End Function



'S_201312_��������_99 �� ���� �߰�
'================================================================
'*  ���θ� �ּ� �˻��� ���� ������ DB���� ������ �������� ������
'*  ��������: 2013.12.12
'*  ������  : ���¿�
'*  Parameter  :
'*   pUseYN : ��� ����
'*---------------------------------------------------------------
'*  �����̷�:
'*---------------------------------------------------------------
'*  ��������    ������  ���泻��
'================================================================
Public Function Gf_DBConnInfo(prs As ADODB.Recordset, pUseYN As String) As Boolean
    Dim lssql                           As String
    Dim rs                              As ADODB.Recordset
    On Error GoTo Err_Rtn
    
    lssql = ""
    lssql = lssql & "  SELECT [ConnectioinType]             " & vbCrLf '��������
    lssql = lssql & "       , [SeverCode]                   " & vbCrLf '�����ڵ�
    lssql = lssql & "       , [SeverName]                   " & vbCrLf '������
    lssql = lssql & "       , [SeverAlias]                  " & vbCrLf '������Ī
    lssql = lssql & "       , [SeverAddress]                " & vbCrLf '�����ּ�
    lssql = lssql & "       , [MangCompany]                 " & vbCrLf '������ü
    lssql = lssql & "       , [DBNameMain]                  " & vbCrLf '����DB��
    lssql = lssql & "       , [DBNameSub]                   " & vbCrLf '����DB��
    lssql = lssql & "       , [PortFrom]                    " & vbCrLf '������Ʈ
    lssql = lssql & "       , [PortTo]                      " & vbCrLf '������Ʈ
    lssql = lssql & "       , [AuthCode1]                   " & vbCrLf '�����ڵ�1
    lssql = lssql & "       , [AuthCode2]                   " & vbCrLf '�����ڵ�2
    lssql = lssql & "       , [SQLAuthType]                 " & vbCrLf 'SQL����Ÿ��
    lssql = lssql & "       , [SQLID]                       " & vbCrLf 'SQL�α���ID
    lssql = lssql & "       , [SQLPass]                     " & vbCrLf 'SQL�α��ξ�ȣ
    lssql = lssql & "       , [PassAuthCode]                " & vbCrLf '��ȣ�����ڵ�
    lssql = lssql & "       , [Comments]                    " & vbCrLf 'Comment
    lssql = lssql & "       , [UseYN]                       " & vbCrLf '��뿩��
    lssql = lssql & "    FROM DBConnInfo                    " & vbCrLf
    lssql = lssql & "   WHERE 1= 1                          " & vbCrLf
    
    '��뿩��
    If pUseYN <> "" Then
        lssql = lssql & "   AND  UseYN='" & pUseYN & "'     " & vbCrLf
    End If
    
    lssql = lssql & " order by  [SeverName]                 " & vbCrLf
    
    If Gf_DB_OpenRS(g_adoCon, prs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DBConnInfo = True
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " / " & Err.Description, vbCritical, "[Gf_DBConnInfo]"
    End If
    Call Gs_DB_CloseRs(prs)
End Function

'S_201312_��������_99 �� ���� �߰�-�ٸ� ��ü�� ���߱� ����
Public Function Gf_DB_CM_GetUserList(prs As ADODB.Recordset, psPersonID As String, psPersonName As String, _
                                     pbIncRetired As Boolean, psdeptCode As String, psDutyCode As String, _
                                     Optional psUserID As String) As Boolean
'================================================================
'*  ����������� Select �Ѵ�.
'*  ��������: 2013.12.12
'*  ������  : �̰��
'*---------------------------------------------------------------
'*  �����̷�:
'*---------------------------------------------------------------
'*  ��������    ������  ���泻��
'================================================================
'* Parameter
'   psUseYN :1 = �������
'            0 = ��ü
'================================================================
    Dim lssql                       As String
    lssql = ""
    lssql = lssql & "  SELECT CM.PERSON_ID , CM.NAME AS PERSON_NAME  , CM.USER_ID     , DBO.ConvP2P(CM.PASSWORD)  AS PASSWORD  " & vbCrLf
    lssql = lssql & "       , CM.DEPT_ID , CD.DEPT_NAME            , CM.DUTY_ID     , CDT.DUTY_NAME                            " & vbCrLf
    lssql = lssql & "       , CM.ENTER_DATE, CM.RETIRE_DATE          , DBO.ConvP2P(CM.REGIST_ID) AS REGIST_ID , CM.HAND_PHONE  " & vbCrLf
    lssql = lssql & "       , CM.PHONE     , CM.BIRTH_DAY            , CM.SOLAR_YN    , CM.ZIP_CODE                            " & vbCrLf
    lssql = lssql & "       , CM.OldNNewClss , ISNULL(CM.GunMoolMngNo,'') AS GunMoolMngNo , ISNULL(CM.ADDRESS1,'') AS ADDRESS1 " & vbCrLf
    lssql = lssql & "       , ISNULL(CM.ADDRESS2,'') AS ADDRESS2  , ISNULL(CM.AddressAssist,'') AS AddressAssist               " & vbCrLf
    lssql = lssql & "       , ISNULL(CM.AddressJiBun1,'') AS AddressJiBun1 , ISNULL(CM.AddressJiBun2,'') AS AddressJiBun2        " & vbCrLf
    lssql = lssql & "       , CM.EMAIL                , CM.REMARK      , CM.USE_YN                              " & vbCrLf
    lssql = lssql & "       , CM.SEX_CODE  , CM.FOREIGN_YN           , CM.WORK_GROUP  , WORK_GROUP_Name =  dbo.[fn_cm_sCodeInfo]('N', CM.WORK_GROUP, 'WKG')  " & vbCrLf
    lssql = lssql & "       , CM.RESABLY_ID ,CRA.RESABLY_NAME        , CM.COMPANY_ID                                           " & vbCrLf
    lssql = lssql & "       , CREATE_DATE  = CONVERT( VARCHAR(30)    , CM.CREATE_DATE, 120) , CM.CREATE_USER_ID                " & vbCrLf
    lssql = lssql & "       , UPDATE_DATE  = CONVERT( VARCHAR(30)    , CM.UPDATE_DATE, 120) , CM.UPDATE_USER_ID                " & vbCrLf
    lssql = lssql & "    FROM [CM_PERSON] CM  left outer join  [CM_DEPT] CD  on CM.DEPT_ID =  CD.DEPT_ID                       " & vbCrLf
    lssql = lssql & "                         left outer join  [CM_DUTY] CDT on CM.DUTY_ID =  CDT.DUTY_ID                      " & vbCrLf
    lssql = lssql & "                         left outer join  [CM_RESABLY] CRA on CM.RESABLY_ID  = CRA.RESABLY_ID             " & vbCrLf
    lssql = lssql & "   WHERE 1=1                                                                                              " & vbCrLf
    lssql = lssql & "     AND CM.USE_YN                                 =  'Y'                                                 " & vbCrLf
    
    '�����ID
    If Trim(psPersonID) <> "" Then
    lssql = lssql & "     AND  CM.PERSON_ID                            LIKE   '%" & Trim(psPersonID) & "%'                     " & vbCrLf
    End If
    
    
    'UserID
    If Trim(psUserID) <> "" Then
    lssql = lssql & "     AND  CM.USER_ID                              LIKE   '%" & Trim(psUserID) & "%'                     " & vbCrLf
    End If
    
    
    
    '����ڸ�"
    If Trim(psPersonName) <> "" Then
    lssql = lssql & "     AND ( CM.NAME                                 LIKE   '%" & Trim(psPersonName) & "%'  )                " & vbCrLf
    End If
    
    '����� ���� ���� �ÿ���
    If pbIncRetired = False Then
    lssql = lssql & "     AND ( CM.RETIRE_DATE                          = '' OR CM.RETIRE_DATE  IS NULL )                      " & vbCrLf
    End If
    
    '�μ��� �˻�
    If psdeptCode <> "" And psdeptCode <> "ALL" Then
    lssql = lssql & "     AND CM.DEPT_ID                                = '" & psdeptCode & "'                                  " & vbCrLf
    End If
    
    '������ �˻�
    If psDutyCode <> "" And psDutyCode <> "ALL" Then
        If InStr(1, psDutyCode, ",") > 0 Then
            lssql = lssql & "     AND CM.DUTY_ID                                in  ( " & psDutyCode & ")                      " & vbCrLf
        Else
            lssql = lssql & "     AND CM.DUTY_ID                                = '" & psDutyCode & "'                         " & vbCrLf
        End If
    End If
    
    lssql = lssql & " ORDER BY PERSON_NAME                                                                                     " & vbCrLf
    '------------------------------------------------------------------------------------------------------------------
    
    If Gf_DB_OpenRS(g_adoCon, prs, lssql) = False Then GoTo Err_Rtn


    Gf_DB_CM_GetUserList = True
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then
        MsgBox "��������� ��ȸ �� ������ �߻��߽��ϴ�!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetUserList]"
    End If
    Call Gs_DB_CloseRs(prs)

End Function


'S_201203_��������_02 �� ���� �߰�
Public Function Gf_DB_OpenRS(ByVal Conn As ADODB.Connection, rs As ADODB.Recordset, cSrc As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Boolean
 
    On Error GoTo Err_Rtn

Retry_rtn:
    Gs_DB_CloseRs rs

    Set rs = New ADODB.Recordset
    rs.Open cSrc, Conn, curType, lockType, opt

    '������ ���� ����� ó��
    
    If Err.Number <> 0 Then
       Gf_DB_OpenRS = False
    End If

    Gf_DB_OpenRS = True
    
    Err.Clear
    Exit Function
Err_Rtn:
    'DB ���� ���н� �ڵ� Retry (3ȸ ����)
   If Err.Number = -2147467259 And giDBConTryCount <= 3 Then
        Set g_adoCon = Nothing
        If Gf_DB_ConnectDB() = False Then Exit Function
        Set Conn = g_adoCon
        GoTo Retry_rtn
        
    End If
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[Gf_DB_OpenRS]"
    End If
    
    
    
    
End Function





'S_201203_��������_02 �� ���� �߰�
 Public Sub Gs_DB_CloseRs(prs As ADODB.Recordset)
    On Error Resume Next
    prs.Close
    Set prs = Nothing
 End Sub


