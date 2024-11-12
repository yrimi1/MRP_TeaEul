Attribute VB_Name = "Declare"
'**************************************************************************************************
'** System 명 : MRRPLUS2-PlusFind
'** Author    : Wizard
'** 작성자    :
'** 내용      :
'** 생성일자  :
'** 변경일자  :
'**------------------------------------------------------------------------------------------------
'
'  요청사항 ID: S_201312_삼우_99
'  요청자:
'  변경날짜 : 2013.12.12
'  작업자   : 오승욱
'  요청내용 : 지번주소에서 도로명 주소로 입력가능하게
'  변경내용 :
'**************************************************************************************************
Option Explicit

Public adoCon As ADODB.Connection

'S_201312_삼우_99 에 의한 추가--------------------------------
Public adoWizCon As ADODB.Connection
Public g_sWizServer$
Public g_sWizDatabase$
Public g_sWizSQLAuthType$           'DB인증방식(1:SQL,2:윈도우)
Public g_sWizSQLID$
Public g_sWizPassword$
Public g_bChkWizDBConn As Boolean
'--------------------------------------------------------

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-22 (THU)
'* UPDATE : 2001-11-30 (FRI)
'*
'* Operate Button의 Index 상수
'********************************************************************************
Public Const ID_ADDNEW As Integer = 0
Public Const ID_UPDATE As Integer = 1
Public Const ID_DELETE As Integer = 2
Public Const ID_SAVE   As Integer = 3
Public Const ID_CANCEL As Integer = 4

'********************************************************
'*
'* Description: CodeFind 대분류
'*
'********************************************************
Public Enum ECODEFIND
    LG_CUSTOM = 0
    LG_ARTICLE = 1
    LG_PERSON = 2
    LG_DEFECT = 3
    LG_ORDER = 4
    LG_DYE = 5
    LG_AUX = 6
    LG_WORK = 7
End Enum

'********************************************************
'*
'* Description: CodeFind 검색방법
'*
'********************************************************
Public Enum EFindClss
    FL_BY_CODE = 0
    FL_BY_NAME = 1
    FL_BY_BTN = 2
End Enum

'S_201312_삼우_99 에 의한 추가
Public giDBConTryCount As Integer

Public Function CheckNull(NewValue As Variant) As String
    If IsNull(NewValue) Then
        CheckNull = ""
    Else
        CheckNull = Trim(NewValue)
    End If
End Function

Public Function CheckNum(NewValue As Variant) As Long
    If IsNumeric(NewValue) Then
        CheckNum = NewValue
    Else
        CheckNum = 0
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* Error 메시지 박스를 출력한다.
'*   - nNum  : 에러번호
'*   - sSrc  : 에러번호
'*   - sDesc : 에러설명
'*   - bExit : "프로그램을 종료합니다." 를 출력할지 선택 (Default = False)
'*   = Return Value : N/A
'********************************************************************************
Public Sub ErrorBox(nNum As Long, sSrc As String, sDesc As String, Optional sTitle As String = "", Optional bExit As Boolean = False)
    Dim sMsg$

    sMsg = "오류가 발생하였습니다. !!!" & vbCrLf & vbCrLf & _
        "오류 번호 : " & CStr(nNum) & vbCrLf & _
        "오류 위치 : " & sSrc & vbCrLf & _
        "오류 설명 " & sDesc & _
        IIf(bExit, vbCrLf & vbCrLf, "")

    sTitle = IIf(Len(sTitle) > 0, sTitle, App.Title)
    Call MsgBox(sMsg, vbCritical, sTitle)
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* VideoSoft FlexGrid에서 화면에 보이는 Row의 갯수를 구한다.
'*   - oGrid : VSFlexGrid
'*   = Return Value : 화면에 보이는 Row의 개수
'********************************************************************************
Public Function GetVisibleVSGridRowCount(oGrid As VSFlexGrid) As Long
    Dim iLoop As Long

    GetVisibleVSGridRowCount = 0

    With oGrid
        For iLoop = .FixedRows To .Rows - .FixedRows
            If Not .RowHidden(iLoop) And .RowHeight(iLoop) > 0 Then
                GetVisibleVSGridRowCount = GetVisibleVSGridRowCount + 1
            End If
        Next iLoop
    End With
End Function


'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-22 (THU)
'* UPDATE :
'*
'* 텍스트 박스의 컬렉션 객체을 받아 내용을 sValue 값으로 초기화 한다.
'*   - oTextBox : 텍스트 박스의 컬렉션 개체
'*   - sValue   : 텍스트 박스를 초기화할 값 (Default = "")
'*   = Return Value : N/A
'********************************************************************************
Public Sub ClearText(oTextBoxs As Object, Optional sValue As String = "")
    Dim oTextBox

    On Error Resume Next

    For Each oTextBox In oTextBoxs
        oTextBox.Text = sValue
    Next
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-17 (SAT)
'* UPDATE :
'*
'* 포커스를 다음 TabIndex의 객체로 이동시킨다.
'*   = Return Value : N/A
'********************************************************************************
Public Sub NextFocus()
    SendKeys "{TAB}"
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* KeyCode로 넘어온 값을 비교하여 ↑키면 포커스를 다음으로 ↓키면 포커스를
'*     이전으로 이동시킨다.
'*   - nKeyCode : KeyCode
'*   = Return Value : N/A
'********************************************************************************
Public Sub MoveFocus(nKeyCode As Integer)
    If nKeyCode = vbKeyDown Then
        nKeyCode = 0
        SendKeys "{TAB}"
    ElseIf nKeyCode = vbKeyUp Then
        nKeyCode = 0
        SendKeys "+{TAB}"
    End If
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* 콤보 박스의 ItemDate()에 nValue와 같은 값이 있는지 검사한다.
'*   - ComboBox : 콤보 박스
'*   - nValue   : ItemData
'*   = Return Value : nValue와 같은 값이 들어있는 ItemData()의 ListIndex.
'********************************************************************************
Public Function FindComboBox(oComboBox As ComboBox, nValue As Long) As Integer
    Dim iLoop As Integer

    FindComboBox = -1
    With oComboBox
        For iLoop = 0 To .ListCount - 1
            If .ItemData(iLoop) = nValue Then
                FindComboBox = iLoop
                Exit For
            End If
        Next iLoop
    End With
End Function

'********************************************************************************
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

        .cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
        .cmdSelect.Picture = LoadResPicture("SELECT", vbResIcon)

        .cmdSearch.MousePointer = ssCustom
    End With
End Sub

'********************************************************************************
'*
'* VideoSoft FlexGrid의 초기값을 설정한다.
'*   - oGrid : VSFlexGrid
'*   = Return Value : N/A
'********************************************************************************
Public Sub SetVSFlexGrid(oGrid As VSFlexGrid, Optional sScroll As String)
    Dim iCount As Integer

    With oGrid
        .Redraw = flexRDNone

        .Rows = 1
        .RowHeight(0) = 450
        .ColWidth(0) = 360

        'S_201312_삼우_99 에 의한 수정
'        .ScrollBars = flexScrollBarVertical
        If sScroll <> "" Then
            .ScrollBars = sScroll
            
        Else
            .ScrollBars = flexScrollBarVertical
        End If
        .ScrollTrack = True
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        .AllowSelection = False
        .AllowBigSelection = False
        .ExtendLastCol = True
        
        .Editable = flexEDNone
        .MousePointer = flexCustom

        .RowHeightMin = 270
        .WordWrap = True

        .ColAlignment(0) = flexAlignCenterCenter
        For iCount = .FixedCols To .Cols - 1
            .FixedAlignment(iCount) = flexAlignCenterCenter
        Next iCount

        .Redraw = flexRDDirect
    End With
End Sub

'********************************************************************************
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

        If bReadOnly Then
            .cmdExit.Cancel = True
            .optSize(0).Value = True
        Else
            .cmdOperate(ID_CANCEL).Cancel = True
            .optSize(1).Value = True
        End If
    End With
End Sub

'S_201312_삼우_99 에 의한 추가
'시도 Select
Public Function Gf_DB_CM_GetSiDoList(pRs As ADODB.Recordset, psUseYN As String) As Boolean

    Dim lssql                           As String
    On Error GoTo Err_Rtn
    lssql = ""
    lssql = lssql & "  SELECT SiDo_Code As Code_ID,SiDo_Name AS CODE_NAME,SiDo_Eng_Name,Seq    " & vbCrLf
    lssql = lssql & "  FROM ZipSiDo                                                                    " & vbCrLf
    lssql = lssql & "  WHERE 1 = 1                                                                      " & vbCrLf
    
        
    lssql = lssql & "     AND SiDo_Code <>'00'                                                  " & vbCrLf
     
     
    '-----------------------------------
    '* 사용여부
    '-----------------------------------
    If psUseYN <> "" Then
        lssql = lssql & "     AND USE_YN                                    =   '" & psUseYN & "'           " & vbCrLf
    End If
    
    
    lssql = lssql & "  ORDER BY SEQ                                                                      " & vbCrLf
     
     'S_201312_삼우_99 에 의한 수정
    If Gf_DB_OpenRS(adoWizCon, pRs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DB_CM_GetSiDoList = True
    
    Exit Function
    
Err_Rtn:

    If Err.Number <> 0 Then
        MsgBox "시도 List Select 중 오류 발생했습니다!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetSiDoList]"
    End If
    
    Call Gs_DB_CloseRs(pRs)

End Function

'시군구 Select
'S_201312_삼우_99 에 의한 추가
Public Function Gf_DB_CM_GetSiGunGuList(pRs As ADODB.Recordset, _
                                        psCodeGroup As String, _
                                        psUseYN As String) As Boolean

    Dim lssql                           As String
    On Error GoTo Err_Rtn
    lssql = ""
    lssql = lssql & "  SELECT SiGunGu_Code As Code_ID,SiGunGu_Name AS CODE_NAME,SiGunGu_Eng_Name,Seq    " & vbCrLf
    lssql = lssql & "  FROM ZipSiGunGu                                                                    " & vbCrLf
    lssql = lssql & "  WHERE 1 = 1                                                                      " & vbCrLf
    
    '-----------------------------------
    '* Code Group
    '-----------------------------------
    If psCodeGroup <> "" And psCodeGroup <> "00" Then
        lssql = lssql & "     AND SubString(SiGunGu_Code,1,2)                                  =   '" & psCodeGroup & "'        " & vbCrLf
    End If
    
    '-----------------------------------
    '* 사용여부
    '-----------------------------------
    If psUseYN <> "" Then
        lssql = lssql & "     AND USE_YN                                    =   '" & psUseYN & "'           " & vbCrLf
    End If
    
    lssql = lssql & "  ORDER BY SEq                                                                      " & vbCrLf
   
     'S_201312_삼우_99 에 의한 수정
    If Gf_DB_OpenRS(adoWizCon, pRs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DB_CM_GetSiGunGuList = True
    
    Exit Function
    
Err_Rtn:

    If Err.Number <> 0 Then
        MsgBox "오류가 발생했습니다!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetSiGunGuList]"
    End If
    
    Call Gs_DB_CloseRs(pRs)

End Function

'''S_201312_삼우_99 에 의한 추가
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
        Set adoCon = Nothing
''        If Gf_DB_ConnectDB() = False Then Exit Function
        Set Conn = adoCon
        GoTo Retry_rtn
        
    End If
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[Gf_DB_OpenRS]"
    End If
       
    
    
End Function

'''S_201312_삼우_99 에 의한 추가
 Public Sub Gs_DB_CloseRs(pRs As ADODB.Recordset)
    On Error Resume Next
    pRs.Close
    Set pRs = Nothing
 End Sub


''
'''S_201312_삼우_99 에 의한 추가
'''****************************************************************
'''*Description:
'''*  ADO를 이용하여 위저드 우변번호 Database에 접속하기
'''****************************************************************
''Public Function ConnectWizDB() As Boolean
''
''    Dim sWizConnect$
''
''    On Error GoTo ErrHandler
''
''    If adoWizCon Is Nothing Then
''
''        If Command() <> "" Then
''            '//테스용
''           ' MsgBox "DB Test 임시"
''          '  g_sServer = "wizis.iptime.org,1433"
''          '  g_sDatabase = "ZipDB"
''
''''            '윈도우인증
''''            sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
''
''            'SQL인증
''            sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=wizardis" & _
''                   ";Initial Catalog=" & sWizDatabase & _
''                   ";Data Source=" & sWizServer & _
''                   ";Use Procedure for Prepare=1;Auto Translate=True;"
''        Else
''
''''            '윈도우인증
''''            sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
''
''            'SQL인증
''            sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=wizardis" & _
''                   ";Initial Catalog=" & sWizDatabase & _
''                   ";Data Source=" & sWizServer & _
''                   ";Use Procedure for Prepare=1;Auto Translate=True;"
''
''        End If
''
''        'S_201312_삼우_99 에 의한 추가-우편번호 조회 관련 connection
''        Set adoWizCon = New ADODB.Connection
''        With adoWizCon
''            .CommandTimeout = 60
''            .ConnectionString = sWizConnect
''            .CursorLocation = adUseClient
''            .Open sWizConnect
''        End With
''
''        ConnectWizDB = True
''    ElseIf adoWizCon.State = adStateOpen Then
''        ConnectWizDB = True
''    Else
''        ConnectWizDB = False
''    End If
''
''    Exit Function
''ErrHandler:
''''    Unload frmSplash
''
''''    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection 실패", True)
''
''    ConnectWizDB = False
''End Function

