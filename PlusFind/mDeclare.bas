Attribute VB_Name = "Declare"
'**************************************************************************************************
'** System �� : MRRPLUS2-PlusFind
'** Author    : Wizard
'** �ۼ���    :
'** ����      :
'** ��������  :
'** ��������  :
'**------------------------------------------------------------------------------------------------
'
'  ��û���� ID: S_201312_���_99
'  ��û��:
'  ���泯¥ : 2013.12.12
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� :
'**************************************************************************************************
Option Explicit

Public adoCon As ADODB.Connection

'S_201312_���_99 �� ���� �߰�--------------------------------
Public adoWizCon As ADODB.Connection
Public g_sWizServer$
Public g_sWizDatabase$
Public g_sWizSQLAuthType$           'DB�������(1:SQL,2:������)
Public g_sWizSQLID$
Public g_sWizPassword$
Public g_bChkWizDBConn As Boolean
'--------------------------------------------------------

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-22 (THU)
'* UPDATE : 2001-11-30 (FRI)
'*
'* Operate Button�� Index ���
'********************************************************************************
Public Const ID_ADDNEW As Integer = 0
Public Const ID_UPDATE As Integer = 1
Public Const ID_DELETE As Integer = 2
Public Const ID_SAVE   As Integer = 3
Public Const ID_CANCEL As Integer = 4

'********************************************************
'*
'* Description: CodeFind ��з�
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
'* Description: CodeFind �˻����
'*
'********************************************************
Public Enum EFindClss
    FL_BY_CODE = 0
    FL_BY_NAME = 1
    FL_BY_BTN = 2
End Enum

'S_201312_���_99 �� ���� �߰�
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
'* Error �޽��� �ڽ��� ����Ѵ�.
'*   - nNum  : ������ȣ
'*   - sSrc  : ������ȣ
'*   - sDesc : ��������
'*   - bExit : "���α׷��� �����մϴ�." �� ������� ���� (Default = False)
'*   = Return Value : N/A
'********************************************************************************
Public Sub ErrorBox(nNum As Long, sSrc As String, sDesc As String, Optional sTitle As String = "", Optional bExit As Boolean = False)
    Dim sMsg$

    sMsg = "������ �߻��Ͽ����ϴ�. !!!" & vbCrLf & vbCrLf & _
        "���� ��ȣ : " & CStr(nNum) & vbCrLf & _
        "���� ��ġ : " & sSrc & vbCrLf & _
        "���� ���� " & sDesc & _
        IIf(bExit, vbCrLf & vbCrLf, "")

    sTitle = IIf(Len(sTitle) > 0, sTitle, App.Title)
    Call MsgBox(sMsg, vbCritical, sTitle)
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* VideoSoft FlexGrid���� ȭ�鿡 ���̴� Row�� ������ ���Ѵ�.
'*   - oGrid : VSFlexGrid
'*   = Return Value : ȭ�鿡 ���̴� Row�� ����
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
'* �ؽ�Ʈ �ڽ��� �÷��� ��ü�� �޾� ������ sValue ������ �ʱ�ȭ �Ѵ�.
'*   - oTextBox : �ؽ�Ʈ �ڽ��� �÷��� ��ü
'*   - sValue   : �ؽ�Ʈ �ڽ��� �ʱ�ȭ�� �� (Default = "")
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
'* ��Ŀ���� ���� TabIndex�� ��ü�� �̵���Ų��.
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
'* KeyCode�� �Ѿ�� ���� ���Ͽ� ��Ű�� ��Ŀ���� �������� ��Ű�� ��Ŀ����
'*     �������� �̵���Ų��.
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
'* �޺� �ڽ��� ItemDate()�� nValue�� ���� ���� �ִ��� �˻��Ѵ�.
'*   - ComboBox : �޺� �ڽ�
'*   - nValue   : ItemData
'*   = Return Value : nValue�� ���� ���� ����ִ� ItemData()�� ListIndex.
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
'* �� ��ư�� �ʱⰪ(Visible, Enable, Image, Cursor)�� �����Ѵ�.
'*   - oForm : ��ư�� �ִ� ��
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
'* VideoSoft FlexGrid�� �ʱⰪ�� �����Ѵ�.
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

        'S_201312_���_99 �� ���� ����
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
'* �߰�, ���� / ����, ��� ��ư Ŭ���� ȣ���Ͽ� ������ ��ư���� ���¸� ����.
'*   - oForm : ��ư�� �ִ� ��
'*   - bFlag : ���� ��
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

'S_201312_���_99 �� ���� �߰�
'�õ� Select
Public Function Gf_DB_CM_GetSiDoList(pRs As ADODB.Recordset, psUseYN As String) As Boolean

    Dim lssql                           As String
    On Error GoTo Err_Rtn
    lssql = ""
    lssql = lssql & "  SELECT SiDo_Code As Code_ID,SiDo_Name AS CODE_NAME,SiDo_Eng_Name,Seq    " & vbCrLf
    lssql = lssql & "  FROM ZipSiDo                                                                    " & vbCrLf
    lssql = lssql & "  WHERE 1 = 1                                                                      " & vbCrLf
    
        
    lssql = lssql & "     AND SiDo_Code <>'00'                                                  " & vbCrLf
     
     
    '-----------------------------------
    '* ��뿩��
    '-----------------------------------
    If psUseYN <> "" Then
        lssql = lssql & "     AND USE_YN                                    =   '" & psUseYN & "'           " & vbCrLf
    End If
    
    
    lssql = lssql & "  ORDER BY SEQ                                                                      " & vbCrLf
     
     'S_201312_���_99 �� ���� ����
    If Gf_DB_OpenRS(adoWizCon, pRs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DB_CM_GetSiDoList = True
    
    Exit Function
    
Err_Rtn:

    If Err.Number <> 0 Then
        MsgBox "�õ� List Select �� ���� �߻��߽��ϴ�!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetSiDoList]"
    End If
    
    Call Gs_DB_CloseRs(pRs)

End Function

'�ñ��� Select
'S_201312_���_99 �� ���� �߰�
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
    '* ��뿩��
    '-----------------------------------
    If psUseYN <> "" Then
        lssql = lssql & "     AND USE_YN                                    =   '" & psUseYN & "'           " & vbCrLf
    End If
    
    lssql = lssql & "  ORDER BY SEq                                                                      " & vbCrLf
   
     'S_201312_���_99 �� ���� ����
    If Gf_DB_OpenRS(adoWizCon, pRs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DB_CM_GetSiGunGuList = True
    
    Exit Function
    
Err_Rtn:

    If Err.Number <> 0 Then
        MsgBox "������ �߻��߽��ϴ�!!" & vbCrLf & _
                Err.Number & Err.Description, vbCritical, "[Gf_DB_CM_GetSiGunGuList]"
    End If
    
    Call Gs_DB_CloseRs(pRs)

End Function

'''S_201312_���_99 �� ���� �߰�
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
        Set adoCon = Nothing
''        If Gf_DB_ConnectDB() = False Then Exit Function
        Set Conn = adoCon
        GoTo Retry_rtn
        
    End If
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[Gf_DB_OpenRS]"
    End If
       
    
    
End Function

'''S_201312_���_99 �� ���� �߰�
 Public Sub Gs_DB_CloseRs(pRs As ADODB.Recordset)
    On Error Resume Next
    pRs.Close
    Set pRs = Nothing
 End Sub


''
'''S_201312_���_99 �� ���� �߰�
'''****************************************************************
'''*Description:
'''*  ADO�� �̿��Ͽ� ������ �캯��ȣ Database�� �����ϱ�
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
''            '//�׽���
''           ' MsgBox "DB Test �ӽ�"
''          '  g_sServer = "wizis.iptime.org,1433"
''          '  g_sDatabase = "ZipDB"
''
''''            '����������
''''            sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
''
''            'SQL����
''            sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=wizardis" & _
''                   ";Initial Catalog=" & sWizDatabase & _
''                   ";Data Source=" & sWizServer & _
''                   ";Use Procedure for Prepare=1;Auto Translate=True;"
''        Else
''
''''            '����������
''''            sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
''
''            'SQL����
''            sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=wizardis" & _
''                   ";Initial Catalog=" & sWizDatabase & _
''                   ";Data Source=" & sWizServer & _
''                   ";Use Procedure for Prepare=1;Auto Translate=True;"
''
''        End If
''
''        'S_201312_���_99 �� ���� �߰�-�����ȣ ��ȸ ���� connection
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
''''    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection ����", True)
''
''    ConnectWizDB = False
''End Function

