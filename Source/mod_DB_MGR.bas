Attribute VB_Name = "mod_DB_MGR"

'**************************************************************************************************
'** System 명 : MRRPLUS2
'** Author    : Wizard
'** 작성자    : 이경미
'** 내용      : DB의 Data 관리를 Business Logic 처리를 위한 Module 이다
'** 생성일자  : 2012.03.06
'**------------------------------------------------------------------------------------------------
'** 요청일자    요청자          요청ID                  변경일자        변경자      요청 및 변경내용
' 2013.12.10  오승욱                 S_201312_태을염직_99    지번주소에서 도로명
'**************************************************************************************************
Option Explicit
 
'S_201312_태을염직_99 에 의한 추가-다른 업체와 맞추기 위함
Public Function Gf_DB_CM_S_CustomOne(prs As ADODB.Recordset, psCustomid As String) As Boolean
    Dim lssql                           As String
    On Error GoTo Err_Rtn
    '조은인더스 소스

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

    MsgBox "고객정보 Select 중 오류 발생했습니다!!" & vbCrLf & _
            Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_S_CustomOne]"
    Call Gs_DB_CloseRs(prs)
End Function


'================================================================
'*  업체 Setting 정보를 가져온다.
'*  생성일자: 2012.03.26
'*  생성자  : 이경미
'*  Parameter  :
'*   pRepresentYN : 대표 사업장 여부
'*   psCompany_No : 사업자번호
'*---------------------------------------------------------------
'*  변경이력:
'*---------------------------------------------------------------
'*  변경일자    변경자  변경내용
'================================================================
Public Function Gf_DB_CM_GetCompanyInfo(prs As ADODB.Recordset, pRepresentYN As String, Optional psCompany_No As String) As Boolean
    Dim lssql                           As String
    Dim rs                              As ADODB.Recordset
    On Error GoTo Err_Rtn
    
    lssql = ""

    lssql = lssql & "  SELECT TOP 1 Company_ID=CompanyID      " & vbCrLf '회사 ID
    lssql = lssql & "       , company_No=CompanyNo      " & vbCrLf '사업자등록번호
    lssql = lssql & "       , company_Name=KCompany     " & vbCrLf '상호
    lssql = lssql & "       , Chief                     " & vbCrLf '사장이름
    lssql = lssql & "       , company_type=Condition    " & vbCrLf '업태
    lssql = lssql & "       , Category                  " & vbCrLf '종목
    lssql = lssql & "       , Zip_Code=ZipCode          " & vbCrLf '우편번호
    'S_201312_태을염직_99 에 의한 추가-------------------------------------------------------
    lssql = lssql & "       , OldNNewClss               " & vbCrLf '주소구분(0:도로명,1:지번주소)
    lssql = lssql & "       , GunMoolMngNo              " & vbCrLf '건물고유식별번호
    lssql = lssql & "       , Address1                  " & vbCrLf '도로명주소1
    lssql = lssql & "       , Address2                  " & vbCrLf '도로명주소2
    lssql = lssql & "       , AddressAssist             " & vbCrLf '도로명 보조 주소
    '--------------------------------------------------------------------------------------
    'S_201312_태을염직_99 에 의한 수정(OLD:Address1)
    lssql = lssql & "       , AddressJiBun1                  " & vbCrLf '지번주소1
    'S_201312_태을염직_99 에 의한 수정(OLD:Address2)
    lssql = lssql & "       , AddressJiBun2                  " & vbCrLf '지번주소2
    lssql = lssql & "       , Phone  = Phone1           " & vbCrLf '전화번호
    lssql = lssql & "       , Phone2                    " & vbCrLf '전화번호2
    lssql = lssql & "       , FaxNO                     " & vbCrLf '팩스번호
    lssql = lssql & "       , Bank1                     " & vbCrLf '계좌1
    lssql = lssql & "       , Bank2                     " & vbCrLf '계좌2
    lssql = lssql & "       , Bank3                     " & vbCrLf '계좌3
    
        '2013.12.12 추가
    '자사 Web관련 추가 정보
    lssql = lssql & "       , WebPortFrom, WebPortTo            " & vbCrLf '웹포트From,To
    lssql = lssql & "       , WebID1, WebPass1, WebAuthCode1    " & vbCrLf 'Web아이디1,암호1,인증코드1
    lssql = lssql & "       , WebID2, WebPass2, WebAuthCode2    " & vbCrLf 'Web아이디2,암호2,인증코드2
    lssql = lssql & "       , FTPPage, FTPPortFrom ,FTPPortTo   " & vbCrLf 'FTP주소,포트From,To
    lssql = lssql & "       , FTPID1, FTPPass1, FTPAuthCode1    " & vbCrLf 'FTP아이디1,암호1,인증코드1
    lssql = lssql & "       , FTPID2, FTPPass2, FTPAuthCode2    " & vbCrLf 'FTP아이디2,암호2,인증코드2
    lssql = lssql & "       , SMSURL1, SMSPortFrom1, SMSPortTo1 " & vbCrLf 'SMS서버1주소,포트From,To
    lssql = lssql & "       , SMSID1, SMSPASS1, SMSAuthCode1    " & vbCrLf 'SMS서버1아이디,암호,인증코드
    lssql = lssql & "       , SMSURL2, SMSPortFrom2, SMSPortTo2 " & vbCrLf 'SMS서버2주소,포트From,To
    lssql = lssql & "       , SMSID2, SMSPASS2, SMSAuthCode2    " & vbCrLf 'SMS서버2아이디,암호,인증코드

    
    
    lssql = lssql & "    FROM mt_setCompany             " & vbCrLf
    lssql = lssql & "   WHERE 1= 1                      " & vbCrLf
    
    '대표여분
    If pRepresentYN <> "" Then
        lssql = lssql & "   AND  RPYN='" & pRepresentYN & "'                      " & vbCrLf
    End If

    lssql = lssql & " order by  RPYN Desc, company_Name            " & vbCrLf
    
    If Gf_DB_OpenRS(g_adoCon, prs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DB_CM_GetCompanyInfo = True
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then
        MsgBox "시스템사용업체 기본정보 Select 중 오류 발생했습니다!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetCompanyInfo]"
    End If
    Call Gs_DB_CloseRs(prs)
End Function



'S_201312_태을염직_99 에 의한 추가
'================================================================
'*  도로명 주소 검색을 위한 위저드 DB연결 정보를 서버에서 가져옴
'*  생성일자: 2013.12.12
'*  생성자  : 오승욱
'*  Parameter  :
'*   pUseYN : 사용 여부
'*---------------------------------------------------------------
'*  변경이력:
'*---------------------------------------------------------------
'*  변경일자    변경자  변경내용
'================================================================
Public Function Gf_DBConnInfo(prs As ADODB.Recordset, pUseYN As String) As Boolean
    Dim lssql                           As String
    Dim rs                              As ADODB.Recordset
    On Error GoTo Err_Rtn
    
    lssql = ""
    lssql = lssql & "  SELECT [ConnectioinType]             " & vbCrLf '접속종류
    lssql = lssql & "       , [SeverCode]                   " & vbCrLf '서버코드
    lssql = lssql & "       , [SeverName]                   " & vbCrLf '서버명
    lssql = lssql & "       , [SeverAlias]                  " & vbCrLf '서버별칭
    lssql = lssql & "       , [SeverAddress]                " & vbCrLf '서버주소
    lssql = lssql & "       , [MangCompany]                 " & vbCrLf '관리업체
    lssql = lssql & "       , [DBNameMain]                  " & vbCrLf '메인DB명
    lssql = lssql & "       , [DBNameSub]                   " & vbCrLf '보조DB명
    lssql = lssql & "       , [PortFrom]                    " & vbCrLf '시작포트
    lssql = lssql & "       , [PortTo]                      " & vbCrLf '종료포트
    lssql = lssql & "       , [AuthCode1]                   " & vbCrLf '인증코드1
    lssql = lssql & "       , [AuthCode2]                   " & vbCrLf '인증코드2
    lssql = lssql & "       , [SQLAuthType]                 " & vbCrLf 'SQL인증타입
    lssql = lssql & "       , [SQLID]                       " & vbCrLf 'SQL로그인ID
    lssql = lssql & "       , [SQLPass]                     " & vbCrLf 'SQL로그인암호
    lssql = lssql & "       , [PassAuthCode]                " & vbCrLf '암호인증코드
    lssql = lssql & "       , [Comments]                    " & vbCrLf 'Comment
    lssql = lssql & "       , [UseYN]                       " & vbCrLf '사용여부
    lssql = lssql & "    FROM DBConnInfo                    " & vbCrLf
    lssql = lssql & "   WHERE 1= 1                          " & vbCrLf
    
    '사용여부
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

'S_201312_태을염직_99 에 의한 추가-다른 업체와 맞추기 위함
Public Function Gf_DB_CM_GetUserList(prs As ADODB.Recordset, psPersonID As String, psPersonName As String, _
                                     pbIncRetired As Boolean, psdeptCode As String, psDutyCode As String, _
                                     Optional psUserID As String) As Boolean
'================================================================
'*  사용자정보를 Select 한다.
'*  생성일자: 2013.12.12
'*  생성자  : 이경미
'*---------------------------------------------------------------
'*  변경이력:
'*---------------------------------------------------------------
'*  변경일자    변경자  변경내용
'================================================================
'* Parameter
'   psUseYN :1 = 사용중인
'            0 = 전체
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
    
    '사용자ID
    If Trim(psPersonID) <> "" Then
    lssql = lssql & "     AND  CM.PERSON_ID                            LIKE   '%" & Trim(psPersonID) & "%'                     " & vbCrLf
    End If
    
    
    'UserID
    If Trim(psUserID) <> "" Then
    lssql = lssql & "     AND  CM.USER_ID                              LIKE   '%" & Trim(psUserID) & "%'                     " & vbCrLf
    End If
    
    
    
    '사용자명"
    If Trim(psPersonName) <> "" Then
    lssql = lssql & "     AND ( CM.NAME                                 LIKE   '%" & Trim(psPersonName) & "%'  )                " & vbCrLf
    End If
    
    '퇴사자 포함 안할 시에는
    If pbIncRetired = False Then
    lssql = lssql & "     AND ( CM.RETIRE_DATE                          = '' OR CM.RETIRE_DATE  IS NULL )                      " & vbCrLf
    End If
    
    '부서별 검색
    If psdeptCode <> "" And psdeptCode <> "ALL" Then
    lssql = lssql & "     AND CM.DEPT_ID                                = '" & psdeptCode & "'                                  " & vbCrLf
    End If
    
    '직무별 검색
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
        MsgBox "사용자정보 조회 중 오류가 발생했습니다!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetUserList]"
    End If
    Call Gs_DB_CloseRs(prs)

End Function


'S_201203_태을염직_02 에 의한 추가
Public Function Gf_DB_OpenRS(ByVal Conn As ADODB.Connection, rs As ADODB.Recordset, cSrc As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Boolean
 
    On Error GoTo Err_Rtn

Retry_rtn:
    Gs_DB_CloseRs rs

    Set rs = New ADODB.Recordset
    rs.Open cSrc, Conn, curType, lockType, opt

    '오류에 따른 사용자 처리
    
    If Err.Number <> 0 Then
       Gf_DB_OpenRS = False
    End If

    Gf_DB_OpenRS = True
    
    Err.Clear
    Exit Function
Err_Rtn:
    'DB 연결 실패시 자동 Retry (3회 까지)
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





'S_201203_태을염직_02 에 의한 추가
 Public Sub Gs_DB_CloseRs(prs As ADODB.Recordset)
    On Error Resume Next
    prs.Close
    Set prs = Nothing
 End Sub


