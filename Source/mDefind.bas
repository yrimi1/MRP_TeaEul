Attribute VB_Name = "Defind"
'******************************************************************************
' �����̷�
'------------------------------------------------------------------------------
'
'��ûID : S_201203_��������_02
'��û���� : 2012.03.05
'��û���� : ������ �� ��µǰ�
'���泻�� : TCOMPANYINFO �߰�
'
' ��ûID : 201208_��������_03
' ��û�� : ����� �븮
' ��û���� : 2012.08.20
' ��û���� : �۾���������ī�忡 ������� ��� ���ַ�(����) ǥ���� �ְ� ������/�е� ��ſ� ������ ���� ��û
' �������� : 2012.08.21
' ���泻�� : Formulas(12), Formulas(15) �� �״�� ��� -> Select �׸� �ٲ�
'
' 2013.12.12  ���¿�                 S_201312_��������_99    �����ּҿ��� ���θ�
'******************************************************************************

Option Explicit

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'***************************************************************************************************
'* Date : 2002-09-05
'*
'* Description: �Ʒ��� API�� ����� �ϴ��� MSDN ���� (�ּ��ް� ������ �׷캰�� �ް�)
'***************************************************************************************************
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' Computer �̸� ��������
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
ByVal lParam As Any) As Long
                                        

'***************************************************************************************************
'* Date : 2002-09-05
'*
'* Description: API���� ����ϴ� ���
'***************************************************************************************************
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNORMAL = 1

Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_BAD_FORMAT = 11&
Public Const ERROR_GEN_FAILURE = 31&

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const MF_STRING = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_SEPERATOR = &H800&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_USERS = &H80000003

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NO_MORE_ITEMS = 259&

Public Const STX = "{"
Public Const ETX = "}"
Public Const ACK = "<"
Public Const NAK = ">"

Private Const WM_CLOSE = &H10
'***************************************************************************************************
'* Date : 2002-09-05
'*
'* Description: ...
'***************************************************************************************************
Public m_sAppFile$
Public g_sServer$
Public g_sDatabase$
Public g_nPrintPort%


Public g_sSamwooServer$
Public g_sSamwooDatabase$


'***************************************************************************************************
'* Date : 2000-06-16 (FRI)
'*
'* Description: Operate Button�� Index ���
'***************************************************************************************************
Public Const ID_ADDNEW As Integer = 0
Public Const ID_UPDATE As Integer = 1
Public Const ID_DELETE As Integer = 2
Public Const ID_SAVE   As Integer = 3
Public Const ID_CANCEL As Integer = 4

Public Enum EDATEFORMAT
    DF_LONG = 0
    DF_SHORT = 1
    DF_FULL = 2
    DF_MID = 3
    DF_MD = 4       ' 12/31 ���·� ��ȯ(����� �߰�)
End Enum

Private g_oCrystalReport As CRPEAuto.Report

'***************************************************************************************************
'* Date : 2000-06-16 (FRI)
'*
'* Description: �������� ���� ����� ����ü
'***************************************************************************************************
Public Const PERM_ADDNEW As String = "PERM_ADDNEW"
Public Const PERM_UPDATE As String = "PERM_UPDATE"
Public Const PERM_DELETE As String = "PERM_DELETE"
Public Const PERM_OUTPUT As String = "PERM_OUTPUT"

Public Const CompanyName As String = ""

'S_201312_��������_99 �� ���� �߰�-------------------------------------
Public g_sWizServer$
Public g_sWizDatabase$
Public g_sWizSQLAuthType$           'DB�������(1:SQL,2:������)
Public g_sWizSQLID$
Public g_sWizPassword$
Public g_bChkWizDBConn As Boolean
Public g_adoWizCon As ADODB.Connection
Public g_DBConnInfo As TDBConnInfo
'--------------------------------------------------------

Public Type TPERMISION
    MenuID As String
    AddNew As Boolean
    Update As Boolean
    Delete As Boolean
    Output As Boolean
End Type

Public g_perm() As TPERMISION


'S_201312_��������_99 �� ���� �߰�
' ȸ�� ����
Public Type TCOMPANYINFO
    Company_ID                          As String 'ȸ�� ID
    Logo                                As String
    Company_Name                        As String '��ȣ
    Chief                               As String '��ǥ���̸�
    Company_No                          As String '����ڵ�Ϲ�ȣ
    Company_type                        As String '���� (Condition As String )
    Category                            As String '����
    OldNNewClss                         As String       '�ּұ���(0:���θ�,1:�����ּ�)
    GunMoolMngNo                        As String       '�ǹ������ĺ���ȣ
    Address1                            As String       '���θ��ּ�1
    Address2                            As String       '���θ��ּ�2
    AddressAssist                       As String       '���θ� ���� �ּ�
    AddressJiBun1                       As String      '�����ּ�1
    AddressJiBun2                       As String       '�����ּ�2
    ZipCode                             As String
    Phone                               As String
    Phone2                              As String
    FaxNO                               As String
    StartTip                            As String
    Advertise                           As String
    Represent_YN                        As String
    
    BANK1                               As String       '���¹�ȣ1
    BANK2                               As String       '���¹�ȣ2
    BANK3                               As String       '���¹�ȣ3
    
    '�߰�����*********************************************************
    ' --WebPage�α�������
    WebPortFrom                         As String   'WebPage��ƮFrom
    WebPortTo                           As String   'WebPage��ƮTo
    WebID1                              As String   'WebPage�α���ID1
    WebPass1                            As String   'WebPage�α��ξ�ȣ1
    WebAuthCode1                        As String   'WebPage�α��������ڵ�1
    WebID2                              As String   'WebPage�α���ID2
    WebPass2                            As String   'WebPage�α��ξ�ȣ2
    WebAuthCode2                        As String   'WebPage�α��������ڵ�2
    
    ' --FTP�α�������
    FTPPage                             As String   'FTP�ּ�
    FTPPortFrom                         As String   'FTP��ƮFrom
    FTPPortTo                           As String   'FTP��ƮTo
    FTPID1                              As String   'FTP�α���ID1
    FTPPass1                            As String   'FTP�α��ξ�ȣ1
    FTPAuthCode1                        As String   'FTP�α��������ڵ�1
    FTPID2                              As String   'FTP�α���ID2
    FTPPass2                            As String   'FTP�α��ξ�ȣ2
    FTPAuthCode2                        As String   'FTP�α��������ڵ�2
    
    ' --SMS����1��������
    SMSURL1                             As String   '�������ۼ���1�ּ�
    SMSPortFrom1                        As String   '�������ۼ���1��ƮFrom
    SMSPortTo1                          As String   '�������ۼ���1��ƮTo
    SMSID1                              As String   '�������ۼ���1���̵�
    SMSPASS1                            As String   '�������ۼ���1��ȣ
    SMSAuthCode1                        As String   '�������ۼ���1�����ڵ�

    ' --SMS����2��������
    SMSURL2                             As String   '�������ۼ���2�ּ�
    SMSPortFrom2                        As String   '�������ۼ���2��ƮFrom
    SMSPortTo2                          As String   '�������ۼ���2��ƮTo
    SMSID2                              As String   '�������ۼ���2���̵�
    SMSPASS2                            As String   '�������ۼ���2��ȣ
    SMSAuthCode2                        As String   '�������ۼ���2�����ڵ�
    '*****************************************************************
End Type

'S_201312_��������_99 �� ���� �߰�
'DB��������
Public Type TDBConnInfo
    ConnectioinType                     As String '��������
    SeverCode                           As String '�����ڵ�
    SeverName                           As String '������
    SeverAlias                          As String '������Ī
    SeverAddress                        As String '�����ּ�
    MangCompany                         As String '������ü
    DBNameMain                          As String '����DB��
    DBNameSub                           As String '����DB��
    PortFrom                            As String '������Ʈ
    PortTo                              As String '������Ʈ
    AuthCode1                           As String '�����ڵ�1
    AuthCode2                           As String '�����ڵ�2
    SQLAuthType                         As String 'SQL����Ÿ��
    SQLID                               As String 'SQL�α���ID
    SQLPass                             As String 'SQL�α��ξ�ȣ
    PassAuthCode                        As String '��ȣ�����ڵ�

End Type

Public g_companyInfo As TCOMPANYINFO

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* Microsoft FlexGrid�� �ʱⰪ�� �����Ѵ�.
'*   - oGrid : MSFlexGrid
'*   = Return Value : N/A
'********************************************************************************
Public Sub SetFlexGrid(oGrid As MSFlexGrid)
    Dim iCount As Integer

    With oGrid
        .Redraw = False

        .Rows = 1
        .RowHeight(0) = 450
        .ColWidth(0) = 360

        .ScrollBars = flexScrollBarVertical
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False

        .RowHeightMin = 275
        .WordWrap = True

        .ColAlignment(0) = flexAlignCenterCenter
        For iCount = 0 To .Cols - 1
            .FixedAlignment(iCount) = flexAlignCenterCenter
        Next iCount

        .Redraw = True
    End With
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* VideoSoft FlexGrid�� �ʱⰪ�� �����Ѵ�.
'*   - oGrid : VSFlexGrid
'*   = Return Value : N/A
'********************************************************************************
Public Sub SetVSFlexGrid(oGrid As VSFlexGrid)
    Dim iCount As Integer

    With oGrid
        .Redraw = flexRDNone

        .Rows = 1
        .RowHeight(0) = 450
        .ColWidth(0) = 360

        .ScrollBars = flexScrollBarVertical
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
'        .MousePointer = flexCustom

        .RowHeightMin = 350
        .WordWrap = True

        .ColAlignment(0) = flexAlignCenterCenter
        For iCount = .FixedCols To .Cols - 1
            .FixedAlignment(iCount) = flexAlignCenterCenter
        Next iCount

        ' Fixed������ �Ӽ�
'        If .Rows > .FixedRows Then
'            .GridLinesFixed = flexGridFlat
'            .GridColorFixed = vbWhite
'            .Cell(flexcpBackColor, 0, 0, .FixedRows - 1, .FixedCols) = vbWhite                      'Fixed ������ FixedCol�� ������ ������� �Ѵ�.
'        End If
        .Redraw = flexRDDirect
    End With
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* FlexGrid�� �ش��÷��� �����Ѵ�.
'*   - oGrid            : FlexGrid
'*   - iCol             : ������ �÷�
'*   - bPrevForwardSort : ���Ĺ�� (True = Descending, False = Ascending)
'*   = Return Value : N/A
'*
'* ��뿹 :
'*   �� Form Module �� ����κп� ������ ���� ������ �����Ѵ�.
'*      Dim m_bSortForward As Boolean
'*   �� Grid�� MouseDown�̺�Ʈ�� ������ �ڵ带 �ۼ��Ѵ�.
'*      With oFlexGrid
'*          If .Rows = .FixedRows Or .MouseRow < 0 Or .MouseRow >= .FixedRows Then Exit Sub
'*
'*          Call SortFlexGrid(grdData, .MouseCol, m_bSortForward)
'*          m_bSortForward = Not m_bSortForward
'*
'*          Call ShowData
'*      End With
'********************************************************************************
Public Sub SortGrid(oGrid As Object, ByVal iCol As Long, ByVal bPrevForwardSort As Boolean)
    Dim nPrevRow%

    With oGrid
        .Col = iCol

        If bPrevForwardSort Then
            Select Case .ColAlignment(.Col)
            Case flexAlignCenterCenter
                .Sort = flexSortGenericDescending
            Case flexAlignLeftCenter
                .Sort = flexSortStringDescending
            Case Else
                .Sort = flexSortNumericDescending
            End Select
        Else
            Select Case .ColAlignment(.Col)
            Case flexAlignCenterCenter
                .Sort = flexSortGenericAscending
            Case flexAlignLeftCenter
                .Sort = flexSortStringAscending
            Case Else
                .Sort = flexSortNumericAscending
            End Select
        End If

        .Col = .FixedCols
        .ColSel = .Cols - 1
    End With
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* Microsoft FlexGrid ???
'*   - oGrid : VSFlexGrid
'*   = Return Value : N/A
'********************************************************************************
'Public Sub SpreadGrid(oGrid As MSFlexGrid)
'    Dim iCount As Integer, RowCount As Integer
'
'    On Error Resume Next
'
'    With oGrid
'        ' �� ���� ���� �����մϴ�.
'        RowCount = .MouseRow
'        If RowCount < 1 Then Exit Sub
'
'        ' ����ϰų� Ȯ���� �ʵ带 ã���ϴ�.
'        While RowCount > 0 And IsNumeric(.TextArray(RowCount * .Cols))
'            RowCount = RowCount - 1
'        Wend
'
'        '   ù° ������ ��ҵ�/Ȯ��� ��ȣ�� �����ݴϴ�.
'        If .TextArray(RowCount * .Cols) = "��" Then
'            .TextArray(RowCount * .Cols) = "��"
'        Else
'            .TextArray(RowCount * .Cols) = "��"
'        End If
'
'        ' ���� �Ӹ��� �Ʒ����� �׸��� Ȯ���մϴ�.
'        RowCount = RowCount + 1
'        If RowCount <= .Rows - 1 Then
'            If .RowHeight(RowCount) = 0 Then
'                Do While IsNumeric(.TextArray(RowCount * .Cols))
'                    .RowHeight(RowCount) = 285   ' Default row height.
'                    RowCount = RowCount + 1
'                    If RowCount >= .Rows Then Exit Do
'                Loop
'            '   ���� �Ӹ��� �Ʒ����� �׸��� ����մϴ�.
'            Else
'                Do While IsNumeric(.TextArray(RowCount * .Cols))
'                    .RowHeight(RowCount) = 0    '    Hide row.
'                    RowCount = RowCount + 1
'                    If RowCount >= .Rows Then Exit Do
'                Loop
'            End If
'        End If
'    End With
'End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* FlexGrid�� �����͸� �ؽ�Ʈ ���Ϸ� �����Ѵ�.
'*   - oGrid : FlexGrid
'*   - sFile : ������ ������ ����� ���ϸ�
'*   = Return Value : ������ ���Ͽ� ��ϵ� ���ڵ� ����
'********************************************************************************
Public Function MakeTextGrid(oGrid As Object, ByVal sFile As String) As Integer
    Dim FileNo As Integer
    Dim iRowCount%, iColCount%, iDataCount%

    On Error GoTo ErrHandler

    If oGrid.Rows = oGrid.FixedRows Then
        MakeTextGrid = 0
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    FileNo = FreeFile
    Open App.Path & sFile For Output Shared As #FileNo

    With oGrid
        For iColCount = 0 To .Cols - 2
                If .ColWidth(iColCount) > 0 Then
                    If .MergeRow(0) Then
                        Write #FileNo, .TextArray(.Cols + iColCount) & iColCount,
                    Else
                        Write #FileNo, .TextArray(iColCount),
                    End If
                End If
        Next iColCount

        If .ColWidth(.Cols - 1) > 0 Then Write #FileNo, .TextArray(.Cols - 1),
        Write #FileNo, "LastTemp"

        For iRowCount = .FixedRows To .Rows - 1
            For iColCount = 0 To .Cols - 1
                If .ColWidth(iColCount) > 0 Then
                    If .ColAlignment(iColCount) = flexAlignRightCenter Then
                        If IsNumeric(.TextArray(iRowCount * .Cols + iColCount)) Then
                            Write #FileNo, CDbl(.TextArray(iRowCount * .Cols + iColCount)),
                        Else
                            Write #FileNo, 0,
                        End If
                    Else
                        Write #FileNo, .TextArray(iRowCount * .Cols + iColCount),
                    End If
                End If
            Next iColCount
            Write #FileNo, " "
            iDataCount = iDataCount + 1
        Next iRowCount
    End With
    MakeTextGrid = iDataCount

    Close #FileNo

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Close #FileNo
    MakeTextGrid = 0

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* FlexGrid�� �����͸� Excel�� �����Ѵ�.
'*   - oGrid : FlexGrid
'*   = Return Value : ���� TRUE, ���� FALSE
'********************************************************************************
Public Function MakeExcelGrid(oGrid As Object) As Boolean
    Dim xlApp   As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    Dim iCol&, irow&, iCols&

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    With oGrid
        iCols = .Cols
        For iCol = 0 To iCols - 1
            For irow = 0 To .Rows - 1
                xlSheet.Cells(irow + 3, iCol + 1) = .TextArray(irow * iCols + iCol)
            Next
        Next
    End With

    xlApp.Visible = True

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    MakeExcelGrid = True

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    MakeExcelGrid = False

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-12 (MON)
'* UPDATE :
'*
'* FlexGrid�� �����͸� HTML ���Ϸ� �����Ѵ�.
'*   - oGrid : FlexGrid
'*   - sFile : HTML ������ ����� ���ϸ�
'*   = Return Value : ���� TRUE, ���� FALSE
'********************************************************************************
Public Function MakeHtmlGrid(oGrid As Object, ByVal sFile As String) As Boolean
    Dim FileNo%, i&, j&

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    FileNo = FreeFile

    Open sFile For Output As #FileNo

    Print #FileNo, "<html>"
    Print #FileNo, "<head>"
    Print #FileNo, "<style type=text/css>"
    Print #FileNo, "table {font-size:9pt}"

    Print #FileNo, "</style>"
    Print #FileNo, "</head>"

    Print #FileNo, "<body bgcolor= #FFFFFF Text =#000000 >"
    Print #FileNo, "<font size=2>"
    Print #FileNo, "<table width=100% border=1 >"

    With oGrid
        ' �׸��� Ÿ��Ʋ �����
        Print #FileNo, "<tr bgcolor = #CCCCCC > "
        For i = 1 To .Cols - 1
            Print #FileNo, "<td align = center height = 33>"; .TextMatrix(0, i); "</td>"
        Next i
        Print #FileNo, "</tr>"

        ' ������ �ֱ�
        For i = 1 To .Rows - 1
            Print #FileNo, "<tr>"

            For j = 1 To .Cols - 1
                If Len(.TextMatrix(i, j)) = 0 Then
                    Print #FileNo, "<td>"; "&nbsp"; "</td>"
                Else
                    If IsNumeric(.TextMatrix(i, j)) Then
                        Print #FileNo, "<td align = Right valign = middle height = 28>"; .TextMatrix(i, j); "</td>"
                    Else
                        Print #FileNo, "<td align = left valign = middle height = 28>"; .TextMatrix(i, j); "</td>"
                    End If
                End If
            Next j

            Print #FileNo, "</tr>"
        Next i
    End With

    Print #FileNo, "</font>"
    Print #FileNo, "</body>"
    Print #FileNo, "</html>"

    Close #FileNo

    MakeHtmlGrid = True

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Close #FileNo
    MakeHtmlGrid = False

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* Micro Soft FlexGrid���� ȭ�鿡 ���̴� Row�� ������ ���Ѵ�.
'*   - oGrid : MSFlexGrid
'*   = Return Value : ȭ�鿡 ���̴� Row�� ����
'********************************************************************************
'Public Function GetVisibleGridRowCount(oGrid As MSFlexGrid) As Long
'    Dim iLoop As Long
'
'    GetVisibleGridRowCount = 0
'
'    With oGrid
'        For iLoop = .FixedRows To .Rows - .FixedRows
'            If .RowHeight(iLoop) > 0 Then
'                GetVisibleGridRowCount = GetVisibleGridRowCount + 1
'            End If
'        Next iLoop
'    End With
'End Function

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
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* VideoSoft FlexGrid���� ȭ�鿡 ���̴� TopRow�� �ε����� ���Ѵ�.
'*   - oGrid : VSFlexGrid
'*   = Return Value : ȭ�鿡 ���̴� Top Row�� �ε���
'********************************************************************************
Public Function GetVisibleVSGridTopRow(oGrid As VSFlexGrid) As Long
    Dim iLoop As Long

    GetVisibleVSGridTopRow = 0

    With oGrid
        For iLoop = .FixedRows To .Rows - .FixedRows
            If Not .RowHidden(iLoop) And .RowHeight(iLoop) > 0 Then
                GetVisibleVSGridTopRow = iLoop
                Exit Function
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
'* CREATE : 2000-06-22 (THU)
'* UPDATE :
'*
'* �޺� �ڽ��� �÷��� ��ü�� �޾� ���� �ڽ��� ListIndex�� iIndex�� �ʱ�ȭ�Ѵ�.
'*   - sComboBoxs : �޺� �ڽ��� �÷��� ��ü
'*   - iIndex     : �޺� �ڽ��� ListIndex�� �ʱ�ȭ�� �� (Default = -1)
'*   = Return Value : N/A
'********************************************************************************
Public Sub ClearCombo(oComboBoxs As Object, Optional iIndex As Long = "-1")
    Dim oComboBox

    On Error Resume Next

    For Each oComboBox In oComboBoxs
        oComboBox.ListIndex = iIndex
    Next
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* Information �޽��� �ڽ��� ����Ѵ�.
'*   - sMsg : ����� �޽��� ����
'*   = Return Value : N/A
'********************************************************************************
Public Sub MessageBox(sMsg As String, Optional nKind As VbMsgBoxStyle = vbInformation)
    Call MsgBox(sMsg, nKind, App.Title)
End Sub

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

    sMsg = LoadResString(991) & vbCrLf & vbCrLf & _
        LoadResString(992) & CStr(nNum) & vbCrLf & _
        LoadResString(993) & sSrc & vbCrLf & _
        LoadResString(994) & sDesc & _
        IIf(bExit, vbCrLf & vbCrLf & LoadResString(995), "")

    sTitle = IIf(Len(sTitle) > 0, sTitle, App.Title)
    Call MsgBox(sMsg, vbInformation, sTitle)
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2001-11-30 (FRI)
'* UPDATE :
'*
'* Question(Yes/No) ���� �޽��� �ڽ��� ����Ѵ�.
'*   - nMsg  : ����� �޽��� ����
'*   = Return Value : Yes���� TRUE, No���� FALSE
'********************************************************************************
Public Function QuestionBox(sMsg As String) As Boolean
    QuestionBox = IIf(MsgBox(sMsg, vbQuestion + vbYesNo, App.Title) = vbYes, True, False)
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-17 (SAT)
'* UPDATE :
'*
'* ���� String �������� ������ nCount�� �Ѿ�� ���� ���� �����Ѵ�.
'*   - sText  : �ؽ�Ʈ
'*   - nCount : �Ҽ������� �ڸ� ��
'*   = Return Value : ����� �ؽ�Ʈ
'********************************************************************************
Public Function SetCurrency(ByVal sText As String, Optional nCount As Integer = 0, Optional nSpace As Integer = 0) As String
    Dim iCount As Integer
    Dim sBaseFmt As String

    sBaseFmt = "#,##0"
    If nCount > 0 Then
        sBaseFmt = "#,##0."
        For iCount = 0 To nCount - 1
            sBaseFmt = sBaseFmt & "0"
        Next iCount
    End If

    If IsNumeric(sText) Then
        SetCurrency = Format(sText, sBaseFmt) & Space(nSpace)
    Else
        SetCurrency = "0" & Space(nSpace)
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2002-03-25
'* UPDATE :
'*
'* ���� ������ ��ȯ�Ѵ�.
'*   - nCount : �Ҽ������� �ڸ� ��
'*   = Return Value : ����� ����
'********************************************************************************
Public Function GetFormat(Optional nCount As Integer = 0) As String
    Dim iCount As Integer

    If nCount > 0 Then
        GetFormat = "#,##0."
        For iCount = 0 To nCount - 1
            GetFormat = GetFormat & "0"
        Next iCount
    Else
        GetFormat = "#,##0"
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-06-21 (WEN)
'* UPDATE :
'*
'* ��¥ String �������� ������ iFormat�� �Ѿ�� ���� ���� �����Ѵ�.
'*   - iFormat : ���� �ε���
'*   - sDate   : ��¥ String ������
'*   = Return Value : String (����� �ؽ�Ʈ)
'********************************************************************************
Public Function MakeDate(ByVal iFormat As EDATEFORMAT, ByVal sDate As String) As String
    Dim sFmt As String

    If iFormat = DF_FULL Then
        sFmt = "YYYY�� MM�� DD��"
    ElseIf iFormat = DF_LONG Then
        sFmt = "YYYY-MM-DD"
    ElseIf iFormat = DF_SHORT Then
        sFmt = "YYYYMMDD"
    ElseIf iFormat = DF_MID Then
        sFmt = "YY-MM-DD"
    ElseIf iFormat = DF_MD Then
        sFmt = "MM/DD"
    End If

    If IsDate(sDate) Then
        MakeDate = Format(sDate, sFmt)
    ElseIf Len(sDate) = 8 Then
        If iFormat = DF_MD Then
            MakeDate = Mid(sDate, 5, 2) & "/" & Right(sDate, 2)
        Else
            MakeDate = Format(Left(sDate, 4) & "-" & Mid(sDate, 5, 2) & "-" & Mid(sDate, 7), sFmt)
        End If
    Else
        MakeDate = ""
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* �ؽ�Ʈ �ڽ��� ��ü �ؽ�Ʈ�� �����Ѵ�.
'*   - oTextBox : �ؽ�Ʈ �ڽ�
'*   = Return Value : N/A
'********************************************************************************
Public Sub GotFocusText(oTextBox As Object)
    On Error Resume Next

    With oTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* �ؽ�Ʈ �ڽ��� KeyCode�� ���ڷ� �޾� ����� KeyCode�� ���� ��ȯ�Ѵ�.
'*   - sText    : TextBox�� Text
'*   - KeyAscii : KeyCode
'*   - bNumber  : ���ڸ� �Է¹����� True, �ƴϸ� False (Default = False)
'*   - nLen     : �Է¹��� �ִ����
'*   = Return Value : ����� KeyCode��
'********************************************************************************
Public Function KeyPress(sText As String, ByVal KeyAscii As Integer, Optional bNumber As Boolean = False, Optional nLen) As Integer
    If KeyAscii <> vbKeyReturn Then
        If bNumber Then
            If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
                If IsMissing(nLen) Then
                    KeyPress = KeyAscii
                Else
                    If Len(sText) >= nLen And KeyAscii <> vbKeyBack Then
                        KeyPress = 0
                    ElseIf Len(sText) < nLen And KeyAscii = vbKeyDelete Then
                        If InStr(1, sText, ".", vbTextCompare) > 0 Then
                            KeyPress = 0
                        Else
                            KeyPress = KeyAscii
                        End If
                    Else
                        KeyPress = KeyAscii
                    End If
                End If
            Else
                KeyPress = 0
            End If
        Else
            If IsMissing(nLen) Then
                KeyPress = Asc(UCase(Chr(KeyAscii)))
            Else
                If Len(sText) >= nLen And KeyAscii <> vbKeyBack Then
                    KeyPress = 0
                Else
                    KeyPress = Asc(UCase(Chr(KeyAscii)))
                End If
            End If
        End If
    End If
End Function

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
    ElseIf nKeyCode = vbKeyReturn Then
        nKeyCode = 0
        SendKeys "{TAB}"
    End If
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
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* ���ڷ� �Ѿ�� ���� �����ش�
'*   - oForm    : ������ ��
'*   - sCaption : ������ ���� Ÿ��Ʋ (Default = "")
'*   = Return Value : N/A
'********************************************************************************
Public Sub ShowForm(oForm As Form, Optional sCaption As String = "")
    Screen.MousePointer = vbHourglass

    With oForm
        .Show
        .Caption = sCaption
        .ZOrder vbBringToFront
    End With

    Screen.MousePointer = vbDefault
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* ���ڷ� �Ѿ�� vValue�� ���� �˻��Ͽ� String���� ��ȯ�Ѵ�.
'*   - vValue : �˻��� ��
'*   = Return Value : �˻� �� ����� ��
'********************************************************************************
Public Function CheckNull(vValue As Variant) As String
    If IsNull(vValue) Then
        CheckNull = ""
    Else
        CheckNull = Trim(CStr(vValue))
    End If
End Function

Public Sub WholeSelect(NewText As TextBox)
    With NewText
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* ���ڷ� �Ѿ�� vValue�� ���� �������� �˻��Ͽ� ���ڷ� ��ȯ�Ѵ�.
'*   - vValue : �˻��� ��
'*   = Return Value : �˻� �� ����� ��
'********************************************************************************
Public Function CheckNum(vValue As Variant) As Currency
    If IsNumeric(vValue) Then
        CheckNum = CCur(vValue)
    Else
        CheckNum = 0
    End If
End Function

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* SplashForm�� nPauseTime ��ŭ Display��Ų�� Unload��ĭ��.
'*   ���� nPauseTime ��ŭ DoEvents�� �ܺ��۾��� �����Ѵ�.
'*   - nPauseTime :
'*   = Return Value : N/A
'********************************************************************************
Public Sub SplashShow(nPauseTime As Single)
    Dim nStart As Single

    frmSplash.Show
    frmSplash.Refresh

    nStart = Timer   ' ���� �ð��� �����մϴ�.
    Do While Timer < nStart + nPauseTime
        DoEvents    ' �ٸ� ���ν����� �ѱ�ϴ�.
    Loop

    Unload frmSplash
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* The higher the "Movement", the slower the window "explosion".
'*   - oForm : A form
'*   - nMove :
'*   = Return Value : N/A
'********************************************************************************
Public Sub ExplodeForm(oForm As Form, ByVal nMove As Integer)
    Dim rcForm As RECT
    Dim nWidth%, nHeight%, i%, X%, Y%, cx%, cy%
    Dim nScreen&, nBrush&

    Call GetWindowRect(oForm.hWnd, rcForm)
    nWidth = (rcForm.Right - rcForm.Left)
    nHeight = rcForm.Bottom - rcForm.Top

    nScreen = GetDC(0)
    nBrush = CreateSolidBrush(oForm.BackColor)

    For i = 1 To nMove
        cx = nWidth * (i / nMove)
        cy = nHeight * (i / nMove)
        X = rcForm.Left + (nWidth - cx) / 2
        Y = rcForm.Top + (nHeight - cy) / 2
        Call Rectangle(nScreen, X, Y, X + cx, Y + cy)
    Next i

    X = ReleaseDC(0, nScreen)
    Call DeleteObject(nBrush)
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2000-08-07 (TUE)
'* UPDATE :
'*
'* The larger the "Movement" value, the slower the "Implosion"
'*   - oForm : A form
'*   - nMove :
'*   = Return Value : N/A
'********************************************************************************
Public Sub ImplodeForm(oForm As Form, nMove As Integer)
    Dim rcForm As RECT
    Dim nWidth%, nHeight%, i%, X%, Y%, cx%, cy%
    Dim nScreen&, nBrush&

    Call GetWindowRect(oForm.hWnd, rcForm)
    nWidth = (rcForm.Right - rcForm.Left)
    nHeight = rcForm.Bottom - rcForm.Top
    nScreen = GetDC(0)
    nBrush = CreateSolidBrush(oForm.BackColor)

    For i = nMove To 1 Step -1
        cx = nWidth * (i / nMove)
        cy = nHeight * (i / nMove)
        X = rcForm.Left + (nWidth - cx) / 2
        Y = rcForm.Top + (nHeight - cy) / 2
        Call Rectangle(nScreen, X, Y, X + cx, Y + cy)
    Next i

    X = ReleaseDC(0, nScreen)
    Call DeleteObject(nBrush)
End Sub

'********************************************************************************
'* AUTHOR : Shaikan
'* CREATE : 2002-01-23 (WED)
'* UPDATE :
'*
'* ������ ����� ���� ���α׷����� ����.
'*   - lHwnd : ������ ȣ���� ���� HWND
'*   - sFile : �� ���ϸ�
'********************************************************************************
Public Sub RelateOpen(lHwnd As Long, sFile As String)
    Dim lReturn As Long

    lReturn = ShellExecute(lHwnd, "open", sFile, vbNullString, vbNullString, SW_SHOWNORMAL)

    Select Case lReturn
        Case ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND
            MsgBox LoadResString(901), vbCritical
        Case ERROR_BAD_FORMAT
            MsgBox LoadResString(902), vbCritical
        Case ERROR_GEN_FAILURE
            MsgBox LoadResString(903), vbCritical
    End Select
End Sub

'****************************************************************
'*Author: Shaikan
'*
'*Description:
'*  INI ���Ͽ��� �ش� Section�� Key�� �ش��ϴ� ���� �о�´�.
'*
'****************************************************************
Public Function GetIniValue(NewSection As String, NewKey As String, Optional NewDefault) As String
    Dim ReturnLength As Long, ReturnValue As String

    ReturnValue = String$(255, &H0)
    If GetPrivateProfileString(NewSection, NewKey, "", ReturnValue, Len(ReturnValue), m_sAppFile) = 0 Then
        If IsMissing(NewDefault) Then '�־��� Default���� ���� ���
            GetIniValue = ""
        Else '�־��� Default���� ���� ���
            GetIniValue = NewDefault
        End If
    Else
        GetIniValue = Left(ReturnValue, InStr(ReturnValue, Chr(0)) - 1)
    End If
End Function

'****************************************************************
'*Author: Shaikan
'*
'*Description:
'*  INI ���Ͽ��� �ش� Section�� Key�� �ش��ϴ� ���� �о�´�.
'*
'****************************************************************
Public Sub SetIniValue(sSection As String, sKey As String, sValue As String, sFileName As String)
    Call WritePrivateProfileString(sSection, sKey, sValue, sFileName)
End Sub

'****************************************************************
'*Author: Shaikan
'*
'*Description:
'*  INI ������ ������ �˱� ���� Windows Folder ��������
'*
'****************************************************************
Public Function GetWindowsPath() As String
    Dim nLength&, sValue$, sWindowsPath

    sValue = String$(255, &H0)
    nLength = GetWindowsDirectory(sValue, Len(sValue))
    GetWindowsPath = Left(sValue, nLength)
End Function

Public Function GetFlexColWidth(nChar As Integer) As Integer
    GetFlexColWidth = (nChar * 90) + 90
End Function

Public Function ArithUpper(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    X = X + (0.49 * Factor)
    ArithUpper = Fix(X * Factor + 0.5 * Sgn(X)) / Factor
End Function

Public Function GetNumeric(sValue As Variant) As Currency
    If Not IsNumeric(sValue) Then
        GetNumeric = 0
    Else
        GetNumeric = CCur(sValue)
    End If
End Function

Public Function KeyPressIsNumeric(ByVal KeyAscii As Integer, Optional IsNumber As Boolean = False) As Integer
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    Else
        If IsNumber Then '������ ��츸 �Է�
            If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
                KeyPressIsNumeric = KeyAscii
            Else
                KeyPressIsNumeric = 0
            End If
        End If
    End If
End Function


Public Sub PrintReport(sReport As String, rs As ADODB.Recordset, Optional vParams, Optional bPreview As Boolean = True)
    Dim oCrystalApp    As CRPEAuto.Application
    Dim oCrystalParams As CRPEAuto.ParameterFieldDefinitions
    Dim i%, k%

    On Error GoTo ErrHandler

    Set oCrystalApp = New CRPEAuto.Application
    Set g_oCrystalReport = oCrystalApp.OpenReport(App.Path & sReport)
    
    g_oCrystalReport.Database.Tables.Item(1).SetPrivateData 3, rs

    If Not IsMissing(vParams) Then
        Set oCrystalParams = g_oCrystalReport.ParameterFields

        For i = 0 To UBound(vParams)
            oCrystalParams.Item(i + 1).SetCurrentValue vParams(i)
            k = oCrystalParams.Count
        Next i

        Set oCrystalParams = Nothing
    End If

    If bPreview Then
        g_oCrystalReport.Preview App.Title, , , , , &H10B0000     ' &H1000000(MAXIMIZE) + &H80000(SYSMENU) + &H10000(MAXIMIZEBOX) + &H20000(MINIMIZEBOX)
    Else
        g_oCrystalReport.PrintOut False
    End If
    rs.Close

    Set oCrystalApp = Nothing
    Set rs = Nothing

    Exit Sub

ErrHandler:
    Set oCrystalApp = Nothing
    Set g_oCrystalReport = Nothing
    Set rs = Nothing

    If Err.Number = 20545 Then Exit Sub ' Requested cancel by user

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function GetComputer() As String
    Dim sBuffer As String
    Dim lLength As Long

    sBuffer = Space(255 + 1)
    lLength = Len(sBuffer)
    
    If CBool(GetComputerName(sBuffer, lLength)) Then
''        GetComputer = Left(sBuffer, lLength)
        '2012.03.06 ����-�ѱ���ǻ�͸� �߶󳻼� DB�Է½� ����
        GetComputer = MidH(sBuffer, 1, lLength)
    Else
        GetComputer = ""
    End If
End Function

Public Function HLen(str As String) As Integer
     Dim i As Integer, chlen As Integer
     
     For i = 1 To Len(str)
          If Asc(Mid$(str, i, 1)) < 0 Then
               chlen = chlen + 2
          Else
               chlen = chlen + 1
          End If
     Next i
     HLen = chlen
End Function

Public Sub PrintWorkCard(cryReport As CrystalReport, sCardID As String, sSplitID As String, sPatternID As String, Optional bPreview As Boolean = True)
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    Dim i%, j%, iFormulas%, iSeq%
    Dim sTmpSplitID$, sProcessPlan$
    Dim sReport$, nRollCnt%, nRollGroup%, nGroupFlag%, sDate$, sSTime$, sETime$
    Dim sRollDetail(3) As String
    Dim nTubeQty(2) As Integer
    Dim nTubeRoll(2) As Integer
    
    On Error GoTo ErrHandler
           
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    Set rs = oCard.GetWorkCard(sCardID, sSplitID)
    Set oCard = Nothing
    
    With cryReport
        .Reset
        .PrintFileType = crptText
        .ReportFileName = App.Path & "\Report\WorkCard.Rpt"
        
        .Formulas(0) = "Title='�۾���������ī��'"
        .Formulas(1) = "BarCode1='*" & rs!CardID & Format(rs!SplitID, "0000") & "*'"
        .Formulas(2) = "BarCode2='" & MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID) & "'"
        .Formulas(3) = "CardID='" & MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID) & "'"
        .Formulas(4) = "StuffINCustom='" & Trim(rs!StuffInCustom) & "'"
        .Formulas(5) = "ItemNo='" & Trim(rs!Item) & "'"
        .Formulas(6) = "OrderID='" & MakeOrderID(rs!OrderID, OM_EXPAND) & "'"
        .Formulas(7) = "PrintDate='" & MakeDate(DF_LONG, Date) & "'"
        .Formulas(8) = "KCustom='" & Trim(rs!kCustom) & "'"
        .Formulas(9) = "OrderNo='" & rs!OrderNo & "'"
        .Formulas(10) = "Color='" & IIf(rs!OrderSeq = 0, "", Trim(rs!Color)) & "'"
        .Formulas(11) = "Article='" & rs!Article & "'"
        
        '201208_��������_03 �� ���� ����(OLD�ҽ�)
''        .Formulas(12) = "WorkName='" & rs!WorkName & "'"
        '201208_��������_03 �� ���� ����(NEW�ҽ�)-���ַ�+����
        .Formulas(12) = "ColorQty='" & Format(rs!ColorQty, "#,###,##0") & " " & IIf(rs!UnitClss = "0", " yds", " MTS") & "'"
        
        .Formulas(13) = "CardRoll='" & Format(rs!Roll, "#,##0") & " ��" & "'"
        .Formulas(14) = "CardQty='" & Format(rs!Qty, "#,##0") & " " & " yds" & "'"
        
        '201208_��������_03 �� ���� ����(OLD�ҽ�)
''        .Formulas(15) = "StuffWidth='" & rs!StuffWidth & "'"
        '201208_��������_03 �� ���� ����(NEW�ҽ�)
        .Formulas(15) = "ChunkRate='" & rs!ChunkRate & "'"

        .Formulas(16) = "WorkWidth='" & rs!WorkWidth & IIf(rs!WorkDensity = 0, "", " / " & rs!WorkDensity) & "'"
        
        .ParameterFields(0) = "Remark" & ";" & CheckNull(rs!Remark) & ";True"
    End With
    rs.Close
    Set rs = Nothing
    
    '���� ���� ���� ī�常 ���� ���� ������ ������.
    If Len(Trim(sSplitID)) = 0 Then
        Set oCard = New PlusLib2.CCard
        oCard.Connection = g_adoCon
        Set rs = oCard.GetWorkCardSub(sCardID, sSplitID)
        Set oCard = Nothing
        
        If Not rs.EOF Then
            nRollCnt = 0
            nRollGroup = 0
            nGroupFlag = rs!RollGroup
            For i = 0 To rs.RecordCount - 1
                If nRollCnt = 15 Then
                    nRollGroup = nRollGroup + 1
                    nRollCnt = 0
                    If nGroupFlag <> rs!RollGroup Then
                        sRollDetail(nRollGroup) = "/" & Format(rs!RollQty, "@@@")
                    Else
                        sRollDetail(nRollGroup) = Format(rs!RollQty, "@@@@")
                    End If
                Else
                    If nGroupFlag <> rs!RollGroup Then
                        sRollDetail(nRollGroup) = sRollDetail(nRollGroup) & "/" & Format(rs!RollQty, "@@@")
                    Else
                        sRollDetail(nRollGroup) = sRollDetail(nRollGroup) & Format(rs!RollQty, "@@@@")
                    End If
                End If
                nRollCnt = nRollCnt + 1
                nGroupFlag = rs!RollGroup
                nTubeRoll(rs!RollGroup - 1) = nTubeRoll(rs!RollGroup - 1) + 1
                nTubeQty(rs!RollGroup - 1) = nTubeQty(rs!RollGroup - 1) + rs!RollQty
                
                rs.MoveNext
            Next i
            rs.Close
            Set rs = Nothing
        
            With cryReport
                .Formulas(17) = "RollDetail1='" & sRollDetail(0) & "'"
                .Formulas(18) = "RollDetail2='" & sRollDetail(1) & "'"
                .Formulas(19) = "RollDetail3='" & sRollDetail(2) & "'"
                .Formulas(20) = "RollDetail4='" & sRollDetail(3) & "'"
            End With
        End If
        
'        If Not rs.EOF Then
'            For i = 0 To rs.RecordCount - 1
'                sRollDetail(rs!RollGroup - 1) = sRollDetail(rs!RollGroup - 1) & rs!RollQty & "   "
'                nTubeRoll(rs!RollGroup - 1) = nTubeRoll(rs!RollGroup - 1) + 1
'                nTubeQty(rs!RollGroup - 1) = nTubeQty(rs!RollGroup - 1) + rs!RollQty
'                rs.MoveNext
'            Next i
'            rs.Close
'            Set rs = Nothing
'
'            With cryReport
'                .Formulas(17) = "RollDetail1='" & sRollDetail(0) & "'"
'                .Formulas(18) = "RollDetail2='" & sRollDetail(1) & "'"
'                .Formulas(19) = "RollDetail3='" & sRollDetail(2) & "'"
'            End With
'        End If
    End If
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    sProcessPlan = oCard.GetWorkProcessPlan(sCardID, sSplitID, sPatternID)
    Set oCard = Nothing
    
    With cryReport
        .Formulas(21) = "ProcessPlan='" & sProcessPlan & "'"
    End With
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    Set rs = oCard.GetWorkPattern(sCardID, sSplitID)
    Set oCard = Nothing
    
    iFormulas = 22
    If Not rs.EOF Then
        With cryReport
            For i = 0 To rs.RecordCount - 1
                .Formulas(iFormulas + rs!PlanSeq - 2) = "ProcessName" & rs!PlanSeq - 1 & "='" & rs!Process & "'"
                
                iFormulas = iFormulas + 1
                rs.MoveNext
            Next i
        End With
    End If
    rs.Close
    Set rs = Nothing

    iFormulas = 41
    iSeq = 1
    For i = 0 To Len(sSplitID)
        If i = 0 Then
            sTmpSplitID = ""
        Else
            sTmpSplitID = Left(sSplitID, i)
        End If
        Set oCard = New PlusLib2.CCard
        oCard.Connection = g_adoCon
        Set rs = oCard.GetWorkCardResult(sCardID, sTmpSplitID)
        Set oCard = Nothing
        With cryReport
            For j = 0 To rs.RecordCount - 1
                sDate = Mid(rs!StartDate, 5, 2) & "-" & Right(rs!StartDate, 2)
                sSTime = Left(rs!StartTime, 2) & ":" & Right(rs!StartTime, 2)
                sETime = Left(rs!EndTime, 2) & ":" & Right(rs!EndTime, 2)

                .Formulas(iFormulas + rs!WorkSeq - 2) = "Machine" & iSeq & "='" & rs!machineid & "'"
                .Formulas(iFormulas + 1 + rs!WorkSeq - 2) = "Person" & iSeq & "='" & rs!Person & "'"
                .Formulas(iFormulas + 2 + rs!WorkSeq - 2) = "StartDate" & iSeq & "='" & sDate & " " & sSTime & " ~ " & sETime & "'"
                .Formulas(iFormulas + 3 + rs!WorkSeq - 2) = "WorkWidth" & iSeq & "='" & IIf(rs!WorkWidth = 0, "", rs!WorkWidth) & "'"
                .Formulas(iFormulas + 4 + rs!WorkSeq - 2) = "WorkDensity" & iSeq & "='" & IIf(rs!WorkDensity = 0, "", rs!WorkDensity) & "'"
                .Formulas(iFormulas + 5 + rs!WorkSeq - 2) = "WorkTemper" & iSeq & "='" & IIf(rs!WorkTemper = 0, "", rs!WorkTemper) & "'"
                .Formulas(iFormulas + 6 + rs!WorkSeq - 2) = "WorkVelocity" & iSeq & "='" & IIf(rs!WorkVelocity = 0, "", rs!WorkVelocity) & "'"
                
                iSeq = iSeq + 1
                iFormulas = iFormulas + 7
                rs.MoveNext
            Next j
            rs.Close
            Set rs = Nothing
        End With
    Next i

    With cryReport
        .Formulas(iFormulas) = "1TubeRoll='" & nTubeRoll(0) & "��'"
        .Formulas(iFormulas + 1) = "1TubeQty='" & nTubeQty(0) & " yds'"
        .Formulas(iFormulas + 2) = "2TubeRoll='" & nTubeRoll(1) & "��'"
        .Formulas(iFormulas + 3) = "2TubeQty='" & nTubeQty(1) & "yds'"
    
        .SelectionFormula = ""
        .WindowState = crptMaximized
        If bPreview Then
            .Destination = crptToWindow
        Else
            .Destination = crptToPrinter
        End If
        .Action = 1
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Call ErrorBox(Err.Number, ".PrintWorkCard", Err.Description)
End Sub

'Public Sub PrintWorkCard(sCardID As String, sSplitID As String, sPatternID As String, Optional bPreview As Boolean = True)
'    Dim oCard As PlusLib2.CCard
'    Dim rs As ADODB.Recordset
'    Dim oExcel      As Excel.Application
'    Dim oExcelBook  As Excel.Workbook
'    Dim oExcelSheet As Excel.Worksheet
'    Dim oFs         As FileSystemObject
'    Dim i%, j%
'    Dim sTmpSplitID$, sProcessPlan$
'    Dim sReport$, nRollGroup%, sRollDetail$, sDate$, sSTime$, sETime$
'
'    On Error GoTo ErrHandler
'
'    Set oExcel = New Excel.Application
'    Set oExcelBook = oExcel.Workbooks.Open(App.Path & "\Report\WorkCard.xls")
'
'    Set oCard = New PlusLib2.CCard
'    oCard.Connection = g_adoCon
'
'    Set rs = oCard.GetWorkCard(sCardID, sSplitID)
'
'    With oExcel
'        .Cells(1, 9) = "*" & rs!CardID & Format(rs!SplitID, "0000") & "*"
'        .Cells(2, 9) = MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID)
'        .Cells(3, 2) = Trim(Replace(rs!StuffINCustom, "(��)", "��"))
'        .Cells(3, 5) = Trim(rs!ThreadName)
'        .Cells(3, 7) = MakeOrderID(rs!OrderID, OM_EXPAND)
'        .Cells(3, 10) = MakeDate(DF_LONG, Date)
'        .Cells(4, 2) = Replace(rs!kCustom, "(��)", "��")
'        .Cells(4, 6) = rs!OrderNo
'        .Cells(4, 10) = IIf(rs!OrderSeq = 0, "", Trim(rs!Color))
'        .Cells(5, 2) = rs!Article
'        .Cells(6, 2) = rs!WorkName
'        .Cells(5, 6) = Format(rs!Roll, "#,##0") & " ��"
'        .Cells(6, 6) = Format(rs!Qty, "#,##0") & " yds"
'        .Cells(5, 10) = rs!StuffWidth
'        .Cells(6, 10) = rs!WorkWidth & IIf(rs!WorkDensity = 0, "", " / " & rs!WorkDensity)
'    End With
'    rs.Close
'    Set rs = Nothing
'
'    '���� ���� ���� ī�常 ���� ���� ������ ������.
'    If Len(Trim(sSplitID)) = 0 Then
'        Set rs = oCard.GetWorkCardSub(sCardID, sSplitID)
'
'        sRollDetail = ""
'        If Not rs.EOF Then
'            With oExcel
'                For i = 0 To rs.RecordCount - 1
'                    If i > 0 And nRollGroup <> rs!RollGroup Then
'                        sRollDetail = sRollDetail & vbLf & rs!RollQty & "   "
'                    Else
'                        sRollDetail = sRollDetail & rs!RollQty & "   "
'                    End If
'                    nRollGroup = rs!RollGroup
'                    rs.MoveNext
'                Next i
'                .Cells(7, 2) = sRollDetail
'            End With
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
'
'    sProcessPlan = oCard.GetWorkProcessPlan(sCardID, sSplitID, sPatternID)
'
'    With oExcel
'        .Cells(8, 2) = sProcessPlan
'    End With
'
'    Set rs = oCard.GetWorkPattern(sCardID, sSplitID)
'
'    If Not rs.EOF Then
'        With oExcel
'            For i = 0 To rs.RecordCount - 1
'                .Cells(10 + rs!PlanSeq - 2, 1) = rs!Process
'
'                rs.MoveNext
'            Next i
'        End With
'    End If
'    rs.Close
'    Set rs = Nothing
'
'    For i = 0 To Len(sSplitID)
'        If i = 0 Then
'            sTmpSplitID = ""
'        Else
'            sTmpSplitID = Left(sSplitID, i)
'        End If
'        Set rs = oCard.GetWorkCardResult(sCardID, sTmpSplitID)
'
'        With oExcel
'            For j = 0 To rs.RecordCount - 1
'                sDate = Mid(rs!StartDate, 5, 2) & "-" & Right(rs!StartDate, 2)
'                sSTime = Left(rs!StartTime, 2) & ":" & Right(rs!StartTime, 2)
'                sETime = Left(rs!Endtime, 2) & ":" & Right(rs!Endtime, 2)
'
'                .Cells(10 + rs!WorkSeq - 2, 2) = rs!MachineID
'                .Cells(10 + rs!WorkSeq - 2, 3) = rs!Person
'                .Cells(10 + rs!WorkSeq - 2, 4) = sDate & " " & sSTime & " ~ " & sETime
'                .Cells(10 + rs!WorkSeq - 2, 5) = IIf(rs!WorkWidth = 0, "", rs!WorkWidth)
'                .Cells(10 + rs!WorkSeq - 2, 6) = IIf(rs!WorkDensity = 0, "", rs!WorkDensity)
'                .Cells(10 + rs!WorkSeq - 2, 7) = IIf(rs!WorkTemper = 9, "", rs!WorkTemper)
'                .Cells(10 + rs!WorkSeq - 2, 8) = IIf(rs!WorkVelocity = 0, "", rs!WorkVelocity)
'
'                rs.MoveNext
'            Next j
'            rs.Close
'            Set rs = Nothing
'        End With
'    Next i
'
'    sReport = App.Path & "\Report\TmpWorkCard.xls"
'    Set oFs = New FileSystemObject
'    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
'    Set oFs = Nothing
'
'    Call oExcelBook.SaveAs(sReport)
'
'    If bPreview Then
'        oExcel.WindowState = xlMaximized
'        oExcel.Application.Visible = True
''        oExcel.ActiveWindow.SelectedSheets.PrintPreview
'    Else
'        oExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=1
'        Call ProcessClose("XLMAIN")
'    End If
'
'    Set oExcelSheet = Nothing
'    Set oExcelBook = Nothing
'    Set oExcel = Nothing
'    Set oCard = Nothing
'
'    Exit Sub
'
'ErrHandler:
'    Set oCard = Nothing
'    Set rs = Nothing
'    Set oExcelSheet = Nothing
'    Set oExcelBook = Nothing
'    Set oExcel = Nothing
'
'    Call ProcessClose("XLMAIN")
'    Call ErrorBox(Err.Number, ".PrintWorkCard", Err.Description)
'End Sub
'
Public Sub ProcessClose(sProcessName As String)
    Dim lngHwnd As Long
    Dim lngRet As Long

    lngHwnd = FindWindow(sProcessName, vbNullString)
    If lngHwnd <> 0 Then
        lngRet = SendMessage(lngHwnd, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Function MakeStrBySpace(sStr As String, nLen As Integer, nDirect As Integer)
    If nDirect = 0 Then
        MakeStrBySpace = Space(nLen - Len(sStr)) & sStr
    Else
        MakeStrBySpace = sStr & Space(nLen - Len(sStr))
    End If
End Function

Public Function MakeNeedQty(nOrderQty As Long, nChunkRate As Single)
    MakeNeedQty = CLng(nOrderQty * (1 + (nChunkRate / 100)))
End Function

' ������ �⺻ �����ͷ� �ǵ��� �ֱ�
Public Sub ReturnPrinter(sPrinter As String)
    Dim dPrinter As Printer
    
    For Each dPrinter In Printers
        If dPrinter.DeviceName = sPrinter Then
            Set Printer = dPrinter
            Exit For
        End If
    Next
End Sub



'2012.03.06 �ű��߰�
'// TODO: �ѿ�ȥ�չ����� ���ʿ������� ������ŭ ���ڸ� �о�´�.
'// ***     MidH����
Function LeftH(str, strlen)
    Dim rValue, tmpStr, tmpASC, lenSum, f
 
  '  If isset(str) Then
        lenSum = 0
        rValue = ""
 
        For f = 1 To Len(str)
            tmpStr = Mid(str, f, 1)
            tmpASC = Asc(tmpStr)
            If tmpASC > 0 And tmpASC < 255 Then lenSum = lenSum + 1 Else lenSum = lenSum + 2
            rValue = rValue & tmpStr
            If (lenSum > strlen) Then Exit For
        Next
        LeftH = rValue
  '  End If
End Function

'2012.03.06 �ű��߰�
'//TODO: �ѿ�ȥ�չ����� �����ʿ������� ������ŭ ���ڸ� �о�´�.
'// ***     MidH����
Function RightH(str, strlen)
    Dim rValue, tmpStr, tmpASC, lenSum, f
 
   ' If isset(str) Then
        lenSum = 0
        rValue = ""
        str = StrReverse(str)
        For f = 1 To Len(str)
            tmpStr = Mid(str, f, 1)
            tmpASC = Asc(tmpStr)
            If tmpASC > 0 And tmpASC < 255 Then lenSum = lenSum + 1 Else lenSum = lenSum + 2
            rValue = rValue & tmpStr
            If (lenSum > strlen) Then Exit For
        Next
        RightH = StrReverse(rValue)
  '  End If
End Function

'2012.03.06 �ű��߰�
'//TODO: �ѿ�ȥ�չ����� start ��ġ�������� length ��ŭ ���ڸ� �о�´�.
Function MidH(s, start, length)
    Dim f, CharAt, VBLength, VBn1, VBn2, BLength, AddByte
    VBn2 = length
    VBLength = Len(s)
    BLength = 0
    For f = 1 To VBLength
        CharAt = Mid(s, f, 1)
        If Asc(CharAt) > 0 And Asc(CharAt) < 255 Then
            BLength = BLength + 1
        Else
            BLength = BLength + 2
        End If
        If BLength >= start Then
            Exit For
        End If
    Next
 
    VBn1 = f
    If VBn1 < 1 Then VBn1 = 1
 
    BLength = 0
    For f = VBn1 To VBLength
        CharAt = Mid(s, f, 1)
        If Asc(CharAt) > 0 And Asc(CharAt) < 255 Then
            BLength = BLength + 1
        Else
            BLength = BLength + 2
        End If
        If BLength = length Then
            VBn2 = f + 1
            Exit For
        ElseIf BLength > length Then
            VBn2 = f
            Exit For
        End If
    Next
    MidH = Mid(s, VBn1, VBn2 - VBn1)
End Function


