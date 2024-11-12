Attribute VB_Name = "Start"
'**************************************************************************************************
'** System 명 : MRRPLUS2
'** Author    : Wizard
'** 작성자    :
'** 내용      : 거래처 등록
'** 생성일자  :
'** 변경일자  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  요청사항 ID: S_201312_태을염직_99
'  요청자:
'  변경날짜 : 2013.12.12
'  작업자   : 오승욱
'  요청내용 : 지번주소에서 도로명 주소로 입력가능하게
'  변경내용 : 도로명,구 지번주소 옵션 버튼 추가
'**************************************************************************************************
Option Explicit

'***************************************************************************************************
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Public Const WH_GETMESSAGE = &H3
Public Const WM_COMMAND = &H111

'***************************************************************************************************
Public Const FORMAT_DEFECTID  As String = "000"
Public Const FORMAT_ARTICLEID As String = "0000"
Public Const FORMAT_BANKID    As String = "00"
Public Const FORMAT_BANDID    As String = "00"
Public Const FORMAT_BASISID   As String = "0"
Public Const FORMAT_BILLID    As String = "0000"
Public Const FORMAT_CARDID    As String = "000"
Public Const FORMAT_CLASSID   As String = "00"
Public Const FORMAT_COLORID   As String = "00"
Public Const FORMAT_CUSTOMID  As String = "000"
Public Const FORMAT_DEPARTID  As String = "00"
Public Const FORMAT_DUTYID    As String = "00"
Public Const FORMAT_DYEAUXID  As String = "0000"
Public Const FORMAT_FORMID    As String = "00"
Public Const FORMAT_GRADEID   As String = "0"
Public Const FORMAT_KINDID    As String = "00"
Public Const FORMAT_LABELID   As String = "00"
Public Const FORMAT_MACHINEID As String = "00"
Public Const FORMAT_MENUID    As String = "0000"
Public Const FORMAT_ORDERID   As String = "0000000000"
Public Const FORMAT_OUTWAREID As String = "00"
Public Const FORMAT_PERSONID  As String = "0000"
Public Const FORMAT_PROCESSID As String = "0"
Public Const FORMAT_TAGID     As String = "000"
Public Const FORMAT_TEAMID    As String = "00"
Public Const FORMAT_TOTALID   As String = "0"
Public Const FORMAT_TRADEID   As String = "00"
Public Const FORMAT_WIDHID    As String = "00"
Public Const FORMAT_WORKID    As String = "00"

Public Const COLOR_GRIDROW As Long = &HE0E0E0

Public Type TLPARAM_GETMSGPROC
    hWnd   As Long
    msg    As Long
    wParam As Long
    lParam As Long
    time   As Long
    ptX    As Long
    ptY    As Long
End Type

Public g_hWndHook As Long

' CodeFind 대분류
Public Enum ECODEFIND
    LG_CUSTOM = 0
    LG_ARTICLE = 1
    LG_PERSON = 2
    LG_DEFECT = 3
    LG_ORDER = 4
    LG_DYE = 5
    LG_AUX = 6
    LG_WORK = 7
    LG_THREAD = 8
    LG_STUFFWIDTH = 9
    LG_PROCESS = 10
    
End Enum

Public Enum EORDERMAKE
    OM_EXPAND = 0
    OM_REDUCE = 1
    OM_COMPACT = 2
End Enum

''' 회사 정보
''Public Type TCOMPANYINFO
''    Logo      As String
''    TradeName As String
''    Chief     As String
''    RegistNO  As String
''    Condition As String
''    Category  As String
''    Address1  As String
''    Address2  As String
''    ZipCode   As String
''    Phone     As String
''    FaxNO     As String
''    StartTip  As String
''    Advertise As String
''End Type

Public g_adoCon As ADODB.Connection
'***************************************************************************************************

Public g_sAppName As String                      '프로그램명

Public g_sCompName As String

'S_201312_태을염직_99 에 의한 수정-주석처리-Defind.Bas 에 정의
''Public g_companyInfo As TCOMPANYINFO
'***************************************************************************************************
Public g_sUserName As String
Public g_sPassword As String
Public g_sPersonName As String
'***************************************************************************************************
Public g_nPointPos%
Public g_bSamwooYN As String



Sub Main()

    Dim sAppname As String
    Dim rs As ADODB.Recordset
    '프로그램 두번 실행되지 않도록 함
    If App.PrevInstance Then Exit Sub

    Call LoadRegistry
    Call LoadINI
    Call SplashShow(3) '스프레쉬 윈도우


    sAppname = App.EXEName
    
    
    If InStr(1, UCase(sAppname), UCase("Samwoo")) > 0 Then        '2022.11.08,lkm,, 태을과 삼우 분리로 추가함
        g_bSamwooYN = "Y"
        g_sDatabase = "SamwoDFC"
        g_companyInfo.Company_Name = ""         '삼우
    Else
        g_bSamwooYN = "N"
        g_companyInfo.Company_Name = ""         '태을염직 , 아래소스를 타지못해서 비워둠...20221122 yhr
    End If
    
       
       
       
    PlusMDI.Show
    frmLogin.Show vbModal 'Login Form을 Load후 UserID와 Passord를 Check 함.
    
    'S_201312_태을염직_99 에 의한 추가
    '암호화 위한 XOR 연산용 데이터 초기화
''    arrEncCode = Array(1, 84, 62, 23, 59, 48, 66, 11, 43, 93, 37, 50, 43, 19, 77, 29, 5, 69, 49, 21)
    Call SetXorData
    
    'S_201312_태을염직_99 에 의한 추가
    '-------------------------------------
    ' 도로명 주소 검색을 위한 위저드 DB연결 정보를 서버에서 가져옴
    '-------------------------------------
    If Gf_DBConnInfo(rs, "Y") = True Then
    
         If rs.EOF = False Then

             g_DBConnInfo.ConnectioinType = Trim(CheckNull(rs!ConnectioinType))  '접속종류
             g_DBConnInfo.SeverCode = Trim(CheckNull(rs!SeverCode))              '서버코드
             g_DBConnInfo.SeverName = Trim(CheckNull(rs!SeverName))              '서버명
             g_DBConnInfo.SeverAlias = Trim(CheckNull(rs!SeverAlias))            '서버별칭
             g_DBConnInfo.SeverAddress = Trim(CheckNull(rs!SeverAddress))        '서버주소
             g_DBConnInfo.MangCompany = Trim(CheckNull(rs!MangCompany))          '관리업체
             g_DBConnInfo.DBNameMain = Trim(CheckNull(rs!DBNameMain))            '메인DB명
             g_DBConnInfo.DBNameSub = Trim(CheckNull(rs!DBNameSub))              '보조DB명
             g_DBConnInfo.PortFrom = Trim(CheckNull(rs!PortFrom))                '시작포트
             g_DBConnInfo.PortTo = Trim(CheckNull(rs!PortTo))                    '종료포트
             g_DBConnInfo.AuthCode1 = Trim(CheckNull(rs!AuthCode1))              '인증코드1
             g_DBConnInfo.AuthCode2 = Trim(CheckNull(rs!AuthCode2))              '인증코드2
             g_DBConnInfo.SQLAuthType = Trim(CheckNull(rs!SQLAuthType))          'SQL인증타입(1:SQL,2:윈도우)
             g_DBConnInfo.SQLID = Trim(CheckNull(rs!SQLID))                      'SQL로그인ID
             g_DBConnInfo.SQLPass = Trim(CheckNull(rs!SQLPass))                  'SQL로그인암호
             g_DBConnInfo.PassAuthCode = Trim(CheckNull(rs!PassAuthCode))        '암호인증코드
                
            'XOR 연산 데이터 배열 재선언
            Call SetXorDataReDim(Len(g_DBConnInfo.PassAuthCode))
             
         End If
            
        'DB에서 읽어온 DB연결 정보를 프로그램내에서 사용 하는 Global변수에 넣어줌
        g_sWizServer = g_DBConnInfo.SeverAddress & IIf(g_DBConnInfo.PortFrom = "", "", ", " & g_DBConnInfo.PortFrom)
        g_sWizDatabase = g_DBConnInfo.DBNameMain
        g_sWizSQLID = g_DBConnInfo.SQLID
''        g_sWizPassword = g_DBConnInfo.SQLPass
        g_sWizPassword = deCode(g_DBConnInfo.PassAuthCode)          '암호화 된 값을 복호화 함
        g_sWizSQLAuthType = Trim(CheckNull(rs!SQLAuthType))         'SQL인증타입(1:SQL,2:윈도우)

    End If
    
    Set rs = Nothing
    '-------------------------------------
    
    'S_201312_태을염직_99 에 의한 추가
    '-------------------------------------
    '자사정보 Get
    '-------------------------------------
    If g_companyInfo.Company_Name = "" Then
        If Gf_DB_CM_GetCompanyInfo(rs, "") = True Then
    
            If rs.EOF = False Then
            
                g_companyInfo.Company_ID = Trim(CheckNull(rs!Company_ID))           '사업장ID
                g_companyInfo.Company_Name = Trim(CheckNull(rs!Company_Name))       '상호
                g_companyInfo.Chief = Trim(CheckNull(rs!Chief))                     '대표자명
                
                'S_201312_태을염직_99 에 의한 추가---------------------------------------------------------------
                g_companyInfo.OldNNewClss = Trim(CheckNull(rs!OldNNewClss))         '주소구분(0:도로명,1:지번)
                g_companyInfo.GunMoolMngNo = Trim(CheckNull(rs!GunMoolMngNo))       '건물고유식별번호
                g_companyInfo.Address1 = Trim(CheckNull(rs!Address1))               '도로명주소1
                g_companyInfo.Address2 = Trim(CheckNull(rs!Address2))               '도로명주소2
                g_companyInfo.AddressAssist = Trim(CheckNull(rs!AddressAssist))     '도로명보조주소
                '------------------------------------------------------------------------------------------------
                
                'S_201312_태을염직_99 에 의한 수정(OLD:Address1)
                g_companyInfo.AddressJiBun1 = Trim(CheckNull(rs!AddressJiBun1))     '지번주소1
                'S_201312_태을염직_99 에 의한 수정(OLD:Address2)
                g_companyInfo.AddressJiBun2 = Trim(CheckNull(rs!AddressJiBun2))     '지번주소2
                
                g_companyInfo.Company_type = Trim(CheckNull(rs!Company_type))       '업태
                g_companyInfo.Category = Trim(CheckNull(rs!Category))               '업종
                g_companyInfo.Company_No = Trim(CheckNull(rs!Company_No))           '사업자번호
            
                'S_201303_조일_06 에 의한 추가-전화번호 및 계좌번호 추가
                g_companyInfo.Phone = Trim(CheckNull(rs!Phone))                     '대표전화번호
                g_companyInfo.Phone2 = Trim(CheckNull(rs!Phone2))                   '전화번호2
                g_companyInfo.FaxNO = Trim(CheckNull(rs!FaxNO))                     '팩스번호
                g_companyInfo.BANK1 = Trim(CheckNull(rs!BANK1))                     '계좌번호1
                g_companyInfo.BANK2 = Trim(CheckNull(rs!BANK2))                     '계좌번호2
                g_companyInfo.BANK3 = Trim(CheckNull(rs!BANK3))                     '계좌번호3
            End If
        End If
    End If
    
    Set rs = Nothing
    '-------------------------------------
    
    
End Sub

'***************************************************************************************************
'*Author: Shaikan
'*
'*Description:
'*  Registry에 등록되어 있는 자사정보 읽어오기.
'***************************************************************************************************
Private Sub LoadRegistry()
''''    'S_201312_태을염직_99d 에 의한 수정-OLD소스
''    g_companyInfo.Logo = GetSetting("MRPPlus", "Company", "Logo")
''    g_companyInfo.Company_Name = GetSetting("MRPPlus", "Company", "TradeName")  '[1] 상호
''    g_companyInfo.Chief = GetSetting("MRPPlus", "Company", "Chief")          '[2] 대표자
''    g_companyInfo.RegistNO = GetSetting("MRPPlus", "Company", "RegistNO")    '[3] 사업자등록번호
''    g_companyInfo.Condition = GetSetting("MRPPlus", "Company", "Condition")  '[4] 업태
''    g_companyInfo.Category = GetSetting("MRPPlus", "Company", "Category")    '[5] 업종
''    g_companyInfo.Address1 = GetSetting("MRPPlus", "Company", "Address1")    '[6] 주소
''    g_companyInfo.Address2 = GetSetting("MRPPlus", "Company", "Address2")
''    g_companyInfo.ZipCode = GetSetting("MRPPlus", "Company", "ZipCode")      '[7] 우편번호
''    g_companyInfo.Phone = GetSetting("MRPPlus", "Company", "Phone")          '[8] 전화번호
''    g_companyInfo.FaxNO = GetSetting("MRPPlus", "Company", "FaxNO")          '[9] 팩스번호
''    g_companyInfo.StartTip = GetSetting("MRPPlus", "Company", "StartTip", vbChecked) '[10] 시작시 표시

    'S_201312_태을염직_99d 에 의한 수정-NEW 소스--------------------------------------------------
    g_companyInfo.Logo = GetSetting("MRPPlus", "Company", "Logo")
    g_companyInfo.Company_Name = GetSetting("MRPPlus", "Company", "Company_Name")       '[1] 상호
    g_companyInfo.Chief = GetSetting("MRPPlus", "Company", "Chief")                     '[2] 대표자
    g_companyInfo.Company_No = GetSetting("MRPPlus", "Company", "Company_No")           '[3] 사업자등록번호
    g_companyInfo.Company_type = GetSetting("MRPPlus", "Company", "Company_type")       '[4] 업태
    g_companyInfo.Category = GetSetting("MRPPlus", "Company", "Category")               '[5] 업종
    g_companyInfo.ZipCode = GetSetting("MRPPlus", "Company", "ZipCode")                 '[6] 우편번호
    g_companyInfo.OldNNewClss = GetSetting("MRPPlus", "Company", "OldNNewClss")         '[7] 주소구분(0:도로명주소,1:지번주소)
    g_companyInfo.GunMoolMngNo = GetSetting("MRPPlus", "Company", "GunMoolMngNo")       '[8] 건물고유식별코드
    g_companyInfo.Address1 = GetSetting("MRPPlus", "Company", "Address1")               '[9] 도로명 기본주소
    g_companyInfo.Address2 = GetSetting("MRPPlus", "Company", "Address2")               '[10] 도로명 상세주소
    g_companyInfo.AddressAssist = GetSetting("MRPPlus", "Company", "AddressAssist")     '[11] 도로명 보조주소
    g_companyInfo.AddressJiBun1 = GetSetting("MRPPlus", "Company", "AddressJiBun1")     '[12] 지번 기본주소
    g_companyInfo.AddressJiBun2 = GetSetting("MRPPlus", "Company", "AddressJiBun2")     '[13] 지번 상세주소
    g_companyInfo.Phone = GetSetting("MRPPlus", "Company", "Phone")                     '[14] 대표전화번호
    g_companyInfo.Phone2 = GetSetting("MRPPlus", "Company", "Phone2")                   '[15] 전화번호
    g_companyInfo.FaxNO = GetSetting("MRPPlus", "Company", "FaxNO")                     '[16] 팩스번호
    g_companyInfo.BANK1 = GetSetting("MRPPlus", "Company", "BANK1")                     '[17] 계좌번호1
    g_companyInfo.BANK2 = GetSetting("MRPPlus", "Company", "BANK2")                     '[18] 계좌번호2
    g_companyInfo.BANK3 = GetSetting("MRPPlus", "Company", "BANK3")                     '[19] 계좌번호3
    g_companyInfo.StartTip = GetSetting("MRPPlus", "Company", "StartTip", vbChecked)    '[20] 시작시 표시
    '----------------------------------------------------------------------------------------------------

End Sub

Public Sub LoadINI()
    Dim nLength&, sValue$, sWindowsPath$
    Dim sServer$, sDatabase$

    sValue = String(255, &H0)
    nLength = GetWindowsDirectory(sValue, Len(sValue))
    sWindowsPath = Left(sValue, nLength)

    m_sAppFile = sWindowsPath & "\Wizard.ini"
    
    'g_sAppName = App.EXEName
    
'    '-------------------------------------------------------------
'    '삼우일 경우
'    '-------------------------------------------------------------
'    If (InStr(1, App.EXEName, "_Samwoo") > 0) Or (g_bSamwooYN = True) Then
''        g_sServer = GetIniValue("SQLServer", "Server", "WZServer")
''        g_sDatabase = GetIniValue("SQLServer", "Database", "MRPPlus")
'        g_nPrintPort = CLng(GetIniValue("COMPort", "TagPrinter", "2"))
'        g_sCompName = "삼우 D.F.C"
'
'    '-------------------------------------------------------------
'    '태을일 경우
'    '-------------------------------------------------------------
'    Else
'        g_sServer = GetIniValue("SQLServer", "Server", "WZServer")
'        g_sDatabase = GetIniValue("SQLServer", "Database", "MRPPlus")
'        g_nPrintPort = CLng(GetIniValue("COMPort", "TagPrinter", "2"))
'        g_sCompName = "태을염직"
'    End

        g_sServer = GetIniValue("SQLServer", "Server", "WZServer")
        g_sDatabase = GetIniValue("SQLServer", "Database", "MRPPlus")
        g_nPrintPort = CLng(GetIniValue("COMPort", "TagPrinter", "2"))
    

End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* 동적으로 생성된 메뉴에서 어떤 메뉴를 선택했는지 Hooking 하는 CALLBACK 함수.
'*   - nCode
'*   - wParam
'*   - lParam
'*   = Return Value : N/A
'********************************************************************************
Function GetMsgProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tMsg As TLPARAM_GETMSGPROC
    Dim hWndHook As Long

    Call CopyMemory(tMsg, ByVal lParam, Len(tMsg))

    If tMsg.msg = WM_COMMAND Then
        If tMsg.wParam > 1000 And tMsg.wParam < 9999 Then
            Call PlusMDI.RunForm(CLng(tMsg.wParam))
        ElseIf tMsg.wParam = 1 Then
            frmSplash.cmdInformation.Visible = True
            frmSplash.cmdOK.Visible = True

            frmSplash.Show
        End If
    End If
    
    '    '//Hokk오류 발생 vbmode로 테스트
    If Command() = "" Then
        GetMsgProc = CallNextHookEx(hWndHook, nCode, wParam, lParam)
    End If
    
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* 각 버튼의 초기값(Visible, Enable, Image, Cursor)을 설정한다.
'*   - oForm : 버튼이 있는 폼
'*   = Return Value : N/A
'********************************************************************************
Public Sub SetOperate(oForm As Form)
    Dim oControl As Object

    On Error Resume Next

    With oForm
        For Each oControl In .Controls
            If (TypeOf oControl Is SSCommand) Or (TypeOf oControl Is CommandButton) _
                Or (TypeOf oControl Is SSOption) Or (TypeOf oControl Is OptionButton) Then
                oControl.MousePointer = ssCustom
                oControl.MouseIcon = LoadResPicture("POINTER", vbResCursor)
            End If
        Next oControl

        .pnlEdit.Enabled = False

        .cmdOperate(ID_SAVE).Visible = False
        .cmdOperate(ID_CANCEL).Visible = False
        .cmdExit.Cancel = True

        .cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
        .cmdOperate(ID_ADDNEW).Picture = LoadResPicture("ADDNEW", vbResIcon)
        .cmdOperate(ID_UPDATE).Picture = LoadResPicture("UPDATE", vbResIcon)
        .cmdOperate(ID_DELETE).Picture = LoadResPicture("DELETE", vbResIcon)
        .cmdOperate(ID_SAVE).Picture = LoadResPicture("SAVE", vbResIcon)
        .cmdOperate(ID_CANCEL).Picture = LoadResPicture("CANCEL", vbResIcon)

        .cmdExcel.Picture = LoadResPicture("EXCEL", vbResIcon)
        .cmdHTML.Picture = LoadResPicture("HTML", vbResIcon)
        .cmdPrint.Picture = LoadResPicture("PRINT", vbResIcon)
        .cmdReport.Picture = LoadResPicture("REPORT", vbResIcon)
        .cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
        .cmdSelect.Picture = LoadResPicture("SELECT", vbResIcon)

        .cmdSearch.MousePointer = ssCustom

        .cmdOperate(ID_ADDNEW).Tag = PERM_ADDNEW
        .cmdOperate(ID_UPDATE).Tag = PERM_UPDATE
        .cmdOperate(ID_DELETE).Tag = PERM_DELETE
        .cmdExcel.Tag = PERM_OUTPUT
        .cmdHTML.Tag = PERM_OUTPUT
        .cmdPrint.Tag = PERM_OUTPUT

'        Call SetPermision(oForm)
    End With
End Sub

Public Sub SetPermision(oForm As Form)
    Dim i%, oControl As Object

    On Error Resume Next

    For i = 0 To UBound(g_perm)
        If oForm.Tag = g_perm(i).MenuID Then Exit For
    Next i

    For Each oControl In oForm.Controls
        If (TypeOf oControl Is SSCommand) Or (TypeOf oControl Is CommandButton) Then
            Select Case oControl.Tag
            Case PERM_ADDNEW
                oControl.Enabled = g_perm(i).AddNew
            Case PERM_UPDATE
                oControl.Enabled = g_perm(i).Update
            Case PERM_DELETE
                oControl.Enabled = g_perm(i).Delete
            Case PERM_OUTPUT
                oControl.Enabled = g_perm(i).Output
            End Select
        End If
    Next oControl
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* 추가, 수정 / 저장, 취소 버튼 클릭시 호출하여 각각의 버튼들의 상태를 변경.
'*   - oForm : 버튼이 있는 폼
'*   - bFlag : 상태 값
'*   = Return Value : N/A
'********************************************************************************
Public Sub ChangeMode(oForm As Form, ByVal bReadOnly As Boolean)
    On Error Resume Next

    With oForm
        .pnlEdit.Enabled = Not bReadOnly
        .pnlSearch.Enabled = bReadOnly
        .pnlMsg.Visible = Not bReadOnly

        .grdData.Enabled = bReadOnly
        .cmdSearch.Enabled = bReadOnly

        .cmdOperate(ID_ADDNEW).Enabled = bReadOnly
        .cmdOperate(ID_UPDATE).Enabled = bReadOnly
        .cmdOperate(ID_DELETE).Enabled = bReadOnly

        .cmdOperate(ID_SAVE).Visible = Not bReadOnly
        .cmdOperate(ID_CANCEL).Visible = Not bReadOnly

        .cmdPrint.Visible = bReadOnly
        .cmdExcel.Visible = bReadOnly
        .cmdHTML.Visible = bReadOnly

        If bReadOnly Then
            .cmdExit.Cancel = True
        Else
            .cmdOperate(ID_CANCEL).Cancel = True
        End If
    End With

    If bReadOnly Then Call SetMsg(LoadResString(301))
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* MDI Form의 상태바에 메시지를 출력한다.
'*   - sMsg : 출력할 메시지 (Default : "")
'*   = Return Value : N/A
'********************************************************************************
Public Sub SetMsg(Optional sMsg As String = "")
    PlusMDI.MainStatus.Panels(1) = sMsg
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* Find 객체에서 해당 데이터를 검색한다.
'*   - nLargeID    : 대분류
'*   - nMiddleID   : 중분류
'*   - bIgnoreData : 전체검색 = True, 조건검색 = False
'*   - txtCode     : 검색조건 / 코드, 명이 들어갈 TextBox
'*   - sFirst      : 검색된 데이터가 들어갈 변수 1
'*   - sSecond     : 검색된 데이터가 들어갈 변수 2
'*   - nThird      : 검색된 데이터가 들어갈 변수 3
'*   = Return Value : 성공 TRUE, 실패 FALSE
'********************************************************************************
Public Function ReturnCode(nLargeID As ECODEFIND, Optional nMiddleID, Optional bIgnoreData As Boolean, _
    Optional txtCode As Object, Optional sData1, Optional sData2, Optional sData3, Optional sData4, Optional sData5, Optional sData6) As Boolean

    Dim oFind As PlusFind2.CCodeFind

    Set oFind = New PlusFind2.CCodeFind
    With oFind
        .Connection = g_adoCon

        If bIgnoreData Then
            ReturnCode = .Find(nLargeID, nMiddleID)
        Else
            ReturnCode = .Find(nLargeID, nMiddleID, txtCode.Text)
        End If

        If ReturnCode Then
            txtCode.Tag = .Data(0)
            txtCode.Text = .Data(1)

            If Not IsMissing(sData1) Then sData1 = .Data(CInt(sData1))
            If Not IsMissing(sData2) Then sData2 = .Data(CInt(sData2))
            If Not IsMissing(sData3) Then sData3 = .Data(CInt(sData3))
            If Not IsMissing(sData4) Then sData4 = .Data(CInt(sData4))
            If Not IsMissing(sData5) Then sData5 = .Data(CInt(sData5))
            If Not IsMissing(sData6) Then sData6 = .Data(CInt(sData6))
        Else
            txtCode.Tag = ""
'            txtCode.Text = ""

            If Not IsMissing(sData1) Then sData1 = ""
            If Not IsMissing(sData2) Then sData2 = ""
            If Not IsMissing(sData3) Then sData3 = ""
            If Not IsMissing(sData4) Then sData4 = ""
            If Not IsMissing(sData5) Then sData5 = ""
            If Not IsMissing(sData6) Then sData6 = ""
            
            MsgBox LoadResString(203), vbInformation
        End If
    End With

    Set oFind = Nothing
End Function

Public Function ReturnRef(nLargeID As ECODEFIND, Optional nMiddleID, Optional bIgnoreData As Boolean, _
    Optional txtCode As Object, Optional sData1, Optional sData2, Optional sData3, Optional sData4, Optional sData5, Optional sData6, Optional sData7) As Boolean

    Dim oWizFind As PlusFind2.CCodeRef

    Set oWizFind = New PlusFind2.CCodeRef
    With oWizFind
        .Connection = g_adoCon

        If bIgnoreData Then
            ReturnRef = .Find(nLargeID, nMiddleID)
        Else
            ReturnRef = .Find(nLargeID, nMiddleID, txtCode.Text)
        End If

        If ReturnRef Then
            txtCode.Tag = .Data(0)
            txtCode.Text = .Data(1)

            If Not IsMissing(sData1) Then sData1 = .Data(CInt(sData1))
            If Not IsMissing(sData2) Then sData2 = .Data(CInt(sData2))
            If Not IsMissing(sData3) Then sData3 = .Data(CInt(sData3))
            If Not IsMissing(sData4) Then sData4 = .Data(CInt(sData4))
            If Not IsMissing(sData5) Then sData5 = .Data(CInt(sData5))
            If Not IsMissing(sData6) Then sData6 = .Data(CInt(sData6))
            If Not IsMissing(sData7) Then sData6 = .Data(CInt(sData7))
        Else
            txtCode.Tag = ""
'            txtCode.Text = ""

            If Not IsMissing(sData1) Then sData1 = ""
            If Not IsMissing(sData2) Then sData2 = ""
            If Not IsMissing(sData3) Then sData3 = ""
            If Not IsMissing(sData4) Then sData4 = ""
            If Not IsMissing(sData5) Then sData5 = ""
            If Not IsMissing(sData6) Then sData6 = ""
            If Not IsMissing(sData7) Then sData7 = ""
        End If
    End With

    Set oWizFind = Nothing
End Function

Public Sub SetDtpDate(nSet As Integer, oStartDate As Object, oEndDate As Object)
    If nSet = 0 Then        ' 전월
        oStartDate = DateSerial(Year(Date), Month(Date) - 1, 1)
        oEndDate = DateSerial(Year(Date), Month(Date), 1 - 1)
    ElseIf nSet = 1 Then    ' 금월
        oStartDate = DateSerial(Year(Date), Month(Date), 1)
        oEndDate = Date
    ElseIf nSet = 2 Then    ' 금일
        oStartDate = Date
        oEndDate = Date
    ElseIf nSet = 3 Then    ' 금년
        oStartDate = DateSerial(Year(Date), 1, 1)
        oEndDate = Date
    End If
End Sub

'********************************************************************************
'* AUTHOR : Littblue
'* CREATE : 2002-03-20 (WED)
'* UPDATE :
'*
'* CombBox에 데이터 채우기.
'*   - Table : 콤보박스에 데이터를 채울 Table 명
'*   - Field : 콤보박스에 데이터를 채울 Field 명
'*   - NewCombo : 콤보박스 오브젝트
'********************************************************************************

Public Sub MakeCodeCombo(CboBox As ComboBox, nClass As ECODE, Optional bSearch As Boolean = False, Optional bSeq As Boolean = True)
    Dim oCode As PlusLib2.CCode
    Dim rs    As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler

    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon

    oCode.CodeType = nClass
    Set rs = oCode.GetCode()
    Set oCode = Nothing

    i = 1
    With CboBox
        .Clear

        If bSearch Then .AddItem "(전체)"

        Do Until rs.EOF
            If bSeq Then
                .AddItem i & ". " & rs(1)
            Else
                .AddItem rs(1)
            End If
            .ItemData(.NewIndex) = rs(0)

            rs.MoveNext
            i = i + 1
        Loop
        rs.Close
        Set rs = Nothing

        .ListIndex = 0
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCode = Nothing

    Err.Raise Err.Number, "Start.MakeCodeCombo", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function MakeCardID(sCardID As String, nType As EORDERMAKE, Optional sSplitID As String = "") As String
    If nType = OM_EXPAND Then
        MakeCardID = Left(sCardID, 2) & "-" & Mid(sCardID, 3, 2) & "-" & Mid(sCardID, 5, 4)
        If Len(sSplitID) > 0 Then
            MakeCardID = MakeCardID & "(" & sSplitID & ")"
        End If
    ElseIf nType = OM_REDUCE Then
        MakeCardID = Replace(sCardID, "-", "")
        MakeCardID = Replace(MakeCardID, "(", "")
        MakeCardID = Replace(MakeCardID, ")", "")
    Else
    
        MakeCardID = CStr(CInt(Mid(sCardID, 3, 2))) & "-" & CStr(CInt(Mid(sCardID, 5, 4)))
    End If
End Function

Public Function MakeOrderID(sOrderID As String, nType As EORDERMAKE) As String
    If nType = OM_EXPAND Then
        MakeOrderID = Left(sOrderID, 4) & "-" & Mid(sOrderID, 5, 2) & "-" & Mid(sOrderID, 7, 4)
    ElseIf nType = OM_REDUCE Then
        MakeOrderID = Replace(sOrderID, "-", "")
    Else
        MakeOrderID = Mid(sOrderID, 5, 2) & "-" & Mid(sOrderID, 7, 4)
    End If
End Function

Public Function MakeTaxSeq(sTaxSeq As String, nType As EORDERMAKE) As String
    If nType = OM_EXPAND Then
        MakeTaxSeq = Left(sTaxSeq, 2) & "-" & Mid(sTaxSeq, 3, 2) & "-" & Mid(sTaxSeq, 5, 4)
    ElseIf nType = OM_REDUCE Then
        MakeTaxSeq = Replace(sTaxSeq, "-", "")
    Else
        MakeTaxSeq = Mid(sTaxSeq, 3, 2) & "-" & Mid(sTaxSeq, 5, 4)
    End If
End Function

Public Function MakeWorkUnitID(sWorkUnitID As String, nType As EORDERMAKE) As String
    If nType = OM_EXPAND Then
        MakeWorkUnitID = Left(sWorkUnitID, 2) & "-" & Mid(sWorkUnitID, 3, 2) & "-" & Mid(sWorkUnitID, 5, 6)
    Else
        MakeWorkUnitID = Replace(sWorkUnitID, "-", "")
    End If
End Function

Public Function MakeArticle(sArticle As String, sWidth As String) As String
    MakeArticle = Trim(sArticle) & " " & Trim(sWidth) & "“"
End Function

Public Function MakeRating(vFlex As Variant, vLoss As Variant) As String
    MakeRating = MakeRating & IIf(IsNumeric(vFlex), vFlex, "0") & "+"
    MakeRating = MakeRating & IIf(IsNumeric(vLoss), vLoss, "0")
End Function

Public Function MakeOrderUnit(vOrderUnit As Variant, Optional bLongName As Boolean = True) As String
    If IsNumeric(vOrderUnit) Then
        Select Case CInt(vOrderUnit)
        Case 0
            MakeOrderUnit = IIf(bLongName, "YDS", "Y")
        Case 1
            MakeOrderUnit = IIf(bLongName, "MTS", "M")
        End Select
    Else
        MakeOrderUnit = ""
    End If
End Function
