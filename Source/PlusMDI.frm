VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm PlusMDI 
   BackColor       =   &H8000000C&
   ClientHeight    =   7965
   ClientLeft      =   3105
   ClientTop       =   3795
   ClientWidth     =   12390
   Icon            =   "PlusMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   WindowState     =   2  '�ִ�ȭ
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   6495
      Top             =   2565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar MainStatus 
      Align           =   2  '�Ʒ� ����
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7635
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12426
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2999
            MinWidth        =   2999
            Object.ToolTipText     =   "User ID"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "�۾�����"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "���� 5:34"
            Object.ToolTipText     =   "���� �ð�"
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pnlMenu 
      Align           =   3  '���� ����
      Height          =   7275
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   12832
      _Version        =   196609
      ForeColor       =   -2147483640
      BackColor       =   12632256
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin VB.CommandButton cmdSize 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1860
         TabIndex        =   7
         Top             =   75
         Width           =   285
      End
      Begin VB.CommandButton cmdSize 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1500
         TabIndex        =   6
         Top             =   75
         Width           =   285
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   105
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "�޴� ���"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   1815
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlTool 
         Left            =   1740
         Top             =   2340
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H00808080&
         BorderStyle     =   0  '����
         FillColor       =   &H00808080&
         Height          =   7515
         Left            =   3090
         ScaleHeight     =   3272.354
         ScaleMode       =   0  '�����
         ScaleWidth      =   780
         TabIndex        =   3
         Top             =   15
         Visible         =   0   'False
         Width           =   72
      End
      Begin MSComctlLib.TreeView trvMenu 
         CausesValidation=   0   'False
         Height          =   5865
         Left            =   60
         TabIndex        =   4
         Top             =   495
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   10345
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTree"
         Appearance      =   1
         MousePointer    =   99
      End
      Begin VB.Image imgSplitter 
         Height          =   7560
         Left            =   3210
         MousePointer    =   9  'W E ũ�� ����
         Top             =   -15
         Width           =   105
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '�� ����
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlTool"
      _Version        =   393216
      BorderStyle     =   1
      MousePointer    =   99
      Begin Threed.SSPanel pnlNumber 
         Height          =   390
         Left            =   10155
         TabIndex        =   8
         Top             =   45
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   688
         _Version        =   196609
         Caption         =   "  ȭ�� ��ȣ (F12)"
         Alignment       =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtName 
            Height          =   300
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   9
            Top             =   45
            Width           =   525
         End
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "�ý���(&S)"
      Begin VB.Menu mnuScreen 
         Caption         =   "ȭ�� ��ȣ"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuMenu 
         Caption         =   "�޴� ���"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTools 
         Caption         =   "���� ���"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSP0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "�α� ��"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "�α� �ƿ�"
      End
      Begin VB.Menu mnuSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "��ȣ ����"
      End
      Begin VB.Menu mnuSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinterSet 
         Caption         =   "������ ����"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSP4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "�̸�����"
      End
      Begin VB.Menu mnuSP5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDirect 
         Caption         =   "�ٷ��μ�"
      End
   End
   Begin VB.Menu mnuErase 
      Caption         =   "���� ����"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "����"
      End
      Begin VB.Menu mnuSP6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "���"
      End
   End
End
Attribute VB_Name = "PlusMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************************************
' �����̷�
'------------------------------------------------------------------------------
'
'��ûID : S_201203_��������_02
'��û���� : 2012.03.05
'��û���� : ������ �� ��µǰ�
'���泻�� : Gf_DB_CM_GetCompanyInfo �߰�
'
'  ��û���� ID: S_201312_��������_99
'  ��û��:
'  ���泯¥ : 2013.12.12
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� : ���θ�,�� �����ּ� �ɼ� ��ư �߰�
'******************************************************************************

Option Explicit

Private Const MAX_SPLIT As Integer = 2500
Private Const MAX_MDIH  As Integer = 12000
Private Const MAX_MDIV  As Integer = 9000

Private m_bMoving  As Boolean ' Splitter ��뿡 ��� �� ����
Private m_bPreview As Boolean ' Print Preview�� ��� �� ����

Private m_nMenuCnt As Integer

Private Const vbAPINull As Long = 0&  ' NULL Pointer

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long

Private Type STARTUPINFO
    cb              As Long
    lpReserved      As String
    lpDesktop       As String
    lpTitle         As String
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadId  As Long
End Type

Private Const CREATE_SUSPENDED = 4&

Private m_stUpgrade As PROCESS_INFORMATION

Private m_oForm As Collection



Public Property Get PrintPreview() As Boolean
    PrintPreview = m_bPreview
End Property

Public Property Let PrintPreview(ByVal bPreview As Boolean)
    m_bPreview = bPreview
End Property

Private Sub MDIForm_Load()
    Dim rs As ADODB.Recordset
    
    Me.Show

    If Not ConnectDB() Then End

    If Len(Trim(g_companyInfo.Company_Name)) > 0 Then
        Me.Caption = LoadResString(101) & " - " & g_companyInfo.Company_Name
    Else
        Me.Caption = LoadResString(101)
    End If
    Me.Caption = Me.Caption & " (Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    
    'S_201312_��������_99 �� ���� ����-Start.Bas�� ������
''    'S_201203_��������_02 �� ���� �߰�
''    '-------------------------------------
''    '��ü���� Get
''    '-------------------------------------
''    If g_companyInfo.Company_Name = "" Then
''        If Gf_DB_CM_GetCompanyInfo(rs, "Y") = True Then
''
''            If rs.EOF = False Then
''                g_companyInfo.Company_Name = Trim(CheckNull(rs!KCompany))    '��ȣ
''                g_companyInfo.Chief = Trim(CheckNull(rs!Chief))                  '��ǥ�ڸ�
''                g_companyInfo.Address1 = Trim(CheckNull(rs!Address1))            '�ּ�1
''                g_companyInfo.Address2 = Trim(CheckNull(rs!Address2))            '�ּ�2
''                g_companyInfo.Company_type = Trim(CheckNull(rs!Condition))    '����
''                g_companyInfo.Category = Trim(CheckNull(rs!Category))            '����
''                g_companyInfo.Company_No = Trim(CheckNull(rs!CompanyNo))        '����ڹ�ȣ
''            End If
''        End If
''    End If
    
    
    ' �̹��� ����Ʈ�� ������ ���� (Ʈ���޴��� ���)
    With imgTree
        .ListImages.Add Key:="Unfolder", Picture:=LoadResPicture("UNFOLDER", vbResIcon)
        .ListImages.Add Key:="Folder", Picture:=LoadResPicture("FOLDER", vbResIcon)
        .ListImages.Add Key:="Close", Picture:=LoadResPicture("CLOSE", vbResIcon)
        .ListImages.Add Key:="Open", Picture:=LoadResPicture("OPEN", vbResIcon)
        .ListImages.Add Key:="Blank", Picture:=LoadResPicture("BLANK", vbResIcon)
        .ListImages.Add Key:="Check", Picture:=LoadResPicture("CHECK", vbResIcon)
    End With
    
    '�̹��� ����Ʈ�� ������ ���� (���ٿ� ���)
    With imlTool
        .ListImages.Add Key:="Back", Picture:=LoadResPicture("BACK", vbResIcon)
        .ListImages.Add Key:="Front", Picture:=LoadResPicture("FRONT", vbResIcon)
        .ListImages.Add Key:="Monitor", Picture:=LoadResPicture("MONITOR", vbResIcon)
        .ListImages.Add Key:="Menu", Picture:=LoadResPicture("MENU", vbResIcon)
        .ListImages.Add Key:="Quit", Picture:=LoadResPicture("QUIT", vbResIcon)
        .ListImages.Add Key:="Close", Picture:=LoadResPicture("FOLDER", vbResIcon)
    End With

    '����(Toolbar) ����
    With tbrMain
        .Buttons.Add Key:="Back", Caption:="�ڷ�", Style:=tbrDefault, Image:="Back"
        .Buttons.Add Key:="Front", Caption:="������", Style:=tbrDefault, Image:="Front"
        .Buttons.Add Style:=tbrSeparator
        .Buttons.Add Key:="Upgrade", Caption:="�ڵ����׷��̵�", Style:=tbrDefault, Image:="Monitor"
        .Buttons.Add Style:=tbrSeparator
        .Buttons.Add Key:="Menu", Caption:="�޴����", Style:=tbrCheck, Image:="Menu"
        .Buttons.Add Style:=tbrSeparator
        .Buttons.Add Key:="Close", Caption:="��δݱ�", Style:=tbrDefault, Image:="Close"
        .Buttons.Add Key:="Quit", Caption:="���� ", Style:=tbrDefault, Image:="Quit"
        .Buttons("Menu").Value = tbrPressed

        .MouseIcon = LoadResPicture("POINTER", vbResCursor)
    End With
    cmdSize(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdSize(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    trvMenu.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    Call SetFormCollection
    Call FirstFormLoad
    
    '//Hokk���� �߻� vbmode�� �׽�Ʈ
    If Command() = "" Then
        g_hWndHook = SetWindowsHookEx(WH_GETMESSAGE, AddressOf GetMsgProc, 0&, GetCurrentThreadId())
    End If


End Sub

Private Sub MDIForm_Activate()
'    On Error Resume Next
'
'    If Len(Trim(g_sUserName)) Then
'        With frmInfo
'            .Show
'
'            .ZOrder vbBringToFront
'        End With
'    End If
End Sub
Private Sub FirstFormLoad()
    On Error Resume Next

    If Len(Trim(g_sUserName)) Then
        With frmInfo
            .Show

            .ZOrder vbBringToFront
        End With
    End If
End Sub


Private Sub MDIForm_Resize()
    Dim iCount As Integer
    Dim lWidth As Long

    On Error Resume Next
    If Me.Width < MAX_MDIH Then Me.Width = MAX_MDIH
    If Me.Height < MAX_MDIV Then Me.Height = MAX_MDIV

    If Me.WindowState <> vbMinimized Then
        trvMenu.Height = Me.Height - 2060
        pnlNumber.Left = Me.Width - 2300 '2050
    End If

    Call SizeControls(imgSplitter.Left)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Dim oLogin As PlusLib2.CLogin
'
'    Set oLogin = New PlusLib2.CLogin
'    oLogin.Connection = g_adoCon
'
'    Call oLogin.LogOff(g_sUserName)
'
'    Set oLogin = Nothing

    If Not (g_adoCon Is Nothing) Then
        g_adoCon.Close
        Set g_adoCon = Nothing
    End If
   
    
    '//Hokk���� �߻� vbmode�� �׽�Ʈ
    If Command() = "" Then
        Call UnhookWindowsHookEx(g_hWndHook)
    End If
    
    Set m_oForm = Nothing

    If Not PlusMDI Is Nothing Then Set PlusMDI = Nothing
    If m_stUpgrade.hThread > 0 Then Call ResumeThread(m_stUpgrade.hThread)

    End
End Sub

'****************************************************************
'*Author: Shaikan
'*
'*Description: ���� �޴� ����
'*  TreeView Control�� ��� �޴��� �����Ѵ�.
'*
'****************************************************************
Public Sub MakeMenu(sUserID As String)
    Dim oMenu As PlusLib2.CMenu
    Dim rs    As ADODB.Recordset

    On Error Resume Next

    Set oMenu = New PlusLib2.CMenu
    oMenu.Connection = g_adoCon

    Set rs = oMenu.GetUserMenu(sUserID)
    Set oMenu = Nothing

    Dim i%, lMenuHwnd&, lPopHwnd&, lSubHwnd&
    Dim sMenuName$

    trvMenu.Nodes.Clear
    lMenuHwnd = GetMenu(Me.hWnd)

    With rs
        i = 0
        Do While Not .EOF
            DoEvents

            ReDim Preserve g_perm(i)
            g_perm(i).MenuID = Format(!MenuID, FORMAT_MENUID)
            g_perm(i).AddNew = IIf(!AddNewClss = "*", True, False)
            g_perm(i).Update = IIf(!UpdateClss = "*", True, False)
            g_perm(i).Delete = IIf(!DeleteClss = "*", True, False)
            g_perm(i).Output = IIf(!PrintClss = "*", True, False)

            If rs!Level = 0 Then
                If CInt(!MenuID) Mod 100 > 0 Then
                    trvMenu.Nodes.Add , , "K" & !MenuID, CStr(!Menu) & "(" & CStr(!MenuID) & ")", "Blank", "Check"
                Else
                    lSubHwnd = CreatePopupMenu()
                    AppendMenu lMenuHwnd, MF_POPUP, lSubHwnd, CStr(!Menu)

                    trvMenu.Nodes.Add , , "K" & !MenuID, !Menu, "Close", "Open"
                    m_nMenuCnt = m_nMenuCnt + 1
                End If
            ElseIf rs!Level = 1 Then
                If CInt(!MenuID) Mod 100 > 0 Then
                    AppendMenu lSubHwnd, MF_STRING, CLng(!MenuID), CStr(!MenuID) & ". " & CStr(!Menu)

                    trvMenu.Nodes.Add "K" & !ParentID, tvwChild, "K" & !MenuID, !Menu & "(" & !MenuID & ")", "Blank", "Check"
                Else
                    lPopHwnd = CreatePopupMenu()
                    AppendMenu lSubHwnd, MF_POPUP, lPopHwnd, CStr(!Menu)

                    trvMenu.Nodes.Add "K" & !ParentID, tvwChild, "K" & !MenuID, !Menu, "Unfolder", "Folder"
                End If
           ElseIf rs!Level = 2 Then
                If CInt(!MenuID) Mod 100 > 0 Then
                    AppendMenu lPopHwnd, MF_STRING, CLng(!MenuID), CStr(!MenuID) & ". " & CStr(!Menu)

                    trvMenu.Nodes.Add "K" & !ParentID, tvwChild, "K" & !MenuID, !Menu & "(" & !MenuID & ")", "Blank", "Check"
                Else
                    lPopHwnd = CreatePopupMenu()
                    AppendMenu lSubHwnd, MF_POPUP, lPopHwnd, CStr(!Menu)

                    trvMenu.Nodes.Add "K" & !ParentID, tvwChild, "K" & !MenuID, !Menu, "Unfolder", "Folder"
                End If
            ElseIf rs!Level = 3 Then
                AppendMenu lPopHwnd, MF_STRING, CLng(rs!MenuID), CStr(rs!MenuID) & ". " & CStr(rs!Menu)

                trvMenu.Nodes.Add "K" & rs!ParentID, tvwChild, "K" & rs!MenuID, rs!Menu & "(" & rs!MenuID & ")", "Blank", "Check"
            End If

            i = i + 1
            .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing

    lSubHwnd = CreatePopupMenu()

    AppendMenu lMenuHwnd, MF_POPUP, lSubHwnd, "����(&H)"
    AppendMenu lSubHwnd, MF_STRING, 1, "MRPPlus2�� ���Ͽ� ....(&A)"

    SetMenu Me.hWnd, lMenuHwnd

    Call cmdSize_Click(0)
    trvMenu.Nodes("K100 ").Expanded = False
    trvMenu.Nodes(1).Selected = True

    Call RunForm(1)
End Sub
'
'Public Sub MakeMenu(NewID As String)
'    Dim oMenu As PlusLib2.CMenu
'    Dim rs    As ADODB.Recordset
'    Dim lReturn&
'
'    On Error Resume Next
'
'    Set oMenu = New PlusLib2.CMenu
'        oMenu.Connection = g_adoCon
'        Set rs = oMenu.GetUserMenu(NewID)
'    Set oMenu = Nothing
'
'    Dim iLoop As Integer
'    Dim lMenuHwnd As Long, lPopHwnd As Long, lSubHwnd As Long
'
'    trvMenu.Nodes.Clear
'    lMenuHwnd = GetMenu(Me.hwnd)
'
'
'    Do While Not rs.EOF
'        DoEvents
'
'        '-------------------------------------------------------------------------------------------------------
'        If rs!Level = 0 Then
'            If (CInt(rs!MenuID) Mod 100) Then
'                trvMenu.Nodes.Add , , "K" & rs!MenuID, rs!Menu & "(" & rs!MenuID & ")", "Blank", "Check"
'            Else
'                lSubHwnd = CreatePopupMenu()
'
'                AppendMenu lMenuHwnd, MF_POPUP, lSubHwnd, CStr(rs!Menu)
'
'                trvMenu.Nodes.Add , , "K" & rs!MenuID, rs!Menu, "Close", "Open"
'                m_nMenuCnt = m_nMenuCnt + 1
'            End If
'        '-------------------------------------------------------------------------------------------------------
'        ElseIf rs!Level = 1 Then
'            If (CInt(rs!MenuID) Mod 100) Then
'                AppendMenu lSubHwnd, MF_STRING, CLng(rs!MenuID), CStr(rs!MenuID) & ". " & CStr(rs!Menu)
'                trvMenu.Nodes.Add "K" & rs!ParentID, tvwChild, "K" & rs!MenuID, rs!Menu & "(" & rs!MenuID & ")", "Blank", "Check"
'
'            Else
'                lPopHwnd = CreatePopupMenu()
'                AppendMenu lSubHwnd, MF_POPUP, lPopHwnd, CStr(rs!Menu)
'                trvMenu.Nodes.Add "K" & rs!ParentID, tvwChild, "K" & rs!MenuID, rs!Menu, "Unfolder", "Folder"
'            End If
'
'        '-------------------------------------------------------------------------------------------------------
'        ElseIf rs!Level = 2 Then
'            If (CInt(rs!MenuID) Mod 100) Then
'                AppendMenu lPopHwnd, MF_STRING, CLng(rs!MenuID), CStr(rs!MenuID) & ". " & CStr(rs!Menu)
'
'                trvMenu.Nodes.Add "K" & rs!ParentID, tvwChild, "K" & rs!MenuID, rs!Menu & "(" & rs!MenuID & ")", "Blank", "Check"
'            Else
'                lPopHwnd = CreatePopupMenu()
'                AppendMenu lSubHwnd, MF_POPUP, lPopHwnd, CStr(rs!Menu)
'
'                trvMenu.Nodes.Add "K" & rs!ParentID, tvwChild, "K" & rs!MenuID, rs!Menu, "Unfolder", "Folder"
'            End If
'        '-------------------------------------------------------------------------------------------------------
'
'        ElseIf rs!Level = 3 Then
'            AppendMenu lPopHwnd, MF_STRING, CLng(rs!MenuID), CStr(rs!MenuID) & ". " & CStr(rs!Menu)
'
'            trvMenu.Nodes.Add "K" & rs!ParentID, tvwChild, "K" & rs!MenuID, rs!Menu & "(" & rs!MenuID & ")", "Blank", "Check"
'        End If
'
'        rs.MoveNext
'
'    Loop
'    rs.Close
'    Set rs = Nothing
'
'    lSubHwnd = CreatePopupMenu()
'    AppendMenu lMenuHwnd, MF_POPUP, lSubHwnd, "����(&H)"
'    AppendMenu lSubHwnd, MF_STRING, 1, "MRP Plus�� ���Ͽ� ....(&A)"
'
'    SetMenu Me.hwnd, lMenuHwnd
'    Call cmdSize_Click(0)
'    trvMenu.Nodes("K100 ").Expanded = False
'
'    trvMenu.Nodes(1).Selected = True
'End Sub

Private Sub cmdSize_Click(Index As Integer)
    Dim i%

    For i = 1 To trvMenu.Nodes.Count - 1
        If trvMenu.Nodes(i).Children > 0 Then
            trvMenu.Nodes(i).Expanded = Index - 1
        End If
    Next i
End Sub

Private Sub ShowAbout()
    With frmSplash
        .cmdInformation.Visible = True
        .cmdOK.Visible = True
        
        .Show vbModal
    End With
End Sub

Private Sub mnuChange_Click()
    frmChange.Show vbModal
End Sub

Private Sub mnuDirect_Click()
    Me.PrintPreview = False
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 1.5, .Height - 20
    End With
    
    picSplitter.Visible = True
    m_bMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sPos As Single

    If m_bMoving Then
        sPos = X + imgSplitter.Left
        If sPos < MAX_SPLIT Then
            pnlMenu.Width = MAX_SPLIT + 50
            trvMenu.Width = MAX_SPLIT - 80
            picSplitter.Left = MAX_SPLIT - 10
        ElseIf sPos > Me.Width - MAX_SPLIT Then
            pnlMenu.Width = Me.Width - MAX_SPLIT + 50
            trvMenu.Width = Me.Width - MAX_SPLIT - 80
            picSplitter.Left = Me.Width - MAX_SPLIT - 10
        Else
            pnlMenu.Width = sPos + 85
            trvMenu.Width = sPos - 45
            picSplitter.Left = sPos + 10
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SizeControls(picSplitter.Left)
    
    picSplitter.Visible = False
    m_bMoving = False
End Sub

Sub SizeControls(X As Single)
    On Error Resume Next
    
    '�ʺ� �����մϴ�.
    If X < MAX_SPLIT Then X = MAX_SPLIT
    If X > (Me.Width - MAX_SPLIT) Then X = Me.Width - MAX_SPLIT

    imgSplitter.Left = trvMenu.Left + trvMenu.Width - 30
    imgSplitter.Height = pnlMenu.Height
End Sub

Private Sub mnuLogin_Click()
    Call ConnectDB
    
    frmLogin.Show vbModal 'Login Form�� Load�� UserID�� Passord�� Check ��.
    
    mnuLogin.Enabled = False
    mnuLogout.Enabled = True
    
    tbrMain.Enabled = True
End Sub

Private Sub mnuLogout_Click()
    Dim lReturn As Long
    Dim i%, iCount%
    Dim lMenuHwnd As Long
    Dim lSubHwnd As Long
    
    
    For iCount = Forms.Count - 1 To 1 Step -1
        Unload Forms(iCount)
    Next iCount
    g_adoCon.Close
    Set g_adoCon = Nothing
    
    mnuLogin.Enabled = True
    mnuLogout.Enabled = False
    
    tbrMain.Enabled = False
    
    trvMenu.Nodes.Clear
    
    '*****************************************************
    '*
    '*   �޴� �� (MenuBar) ����
    '*
    '*****************************************************
    lMenuHwnd = GetMenu(Me.hWnd)
    For i = 1 To m_nMenuCnt + 1
        lSubHwnd = GetSubMenu(lMenuHwnd, 1)
        lReturn = DeleteMenu(lMenuHwnd, lSubHwnd, MF_BYCOMMAND)
    Next i
    lReturn = DrawMenuBar(Me.hWnd)
    m_nMenuCnt = 0
    
    Call mnuLogin_Click
End Sub

Private Sub mnuMenu_Click()
    Call ShowMenu(Not mnuMenu.Checked)
End Sub


Private Sub ShowMenu(NewValue As Boolean)
    pnlMenu.Visible = NewValue
    mnuMenu.Checked = NewValue
    If NewValue Then
        tbrMain.Buttons("Menu").Value = tbrPressed
        txtName.SetFocus
    Else
        tbrMain.Buttons("Menu").Value = tbrUnpressed
    End If
    
End Sub


Private Sub mnuPreview_Click()
    Me.PrintPreview = True
End Sub

Private Sub mnuPrinterSet_Click()
    dlgDialog.ShowPrinter
End Sub

Private Sub mnuScreen_Click()
    txtName.SetFocus
End Sub

Private Sub mnuTools_Click()
    If tbrMain.Visible Then
        tbrMain.Visible = False
    Else
        tbrMain.Visible = True
    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i%
    Dim SI As STARTUPINFO

    Select Case Button.Key
        Case "Back"
            For i = Forms.Count - 1 To 1 Step -1
                If ActiveForm.Name = Forms(i).Name Then
                    Forms(i - 1).ZOrder vbBringToFront
                    Exit For
                End If
            Next i
        Case "Front"
            For i = Forms.Count - 1 To 1 Step -1
                If ActiveForm.Name = Forms(i).Name Then
                    Forms(Abs((i + 1) - Forms.Count) + 1).ZOrder vbBringToFront
                    Exit For
                End If
            Next i
        Case "Menu"
            If Button.Value = tbrPressed Then
                Call ShowMenu(True)
            Else
                Call ShowMenu(False)
            End If
        Case "Close"
            For i = Forms.Count - 1 To 1 Step -1
                Unload Forms(i)
            Next i
        Case "Upgrade"
            For i = Forms.Count - 1 To 1 Step -1
                Unload Forms(i)
            Next i
            Call CreateProcess(0&, App.Path & "\Upgrade.exe", 0&, 0&, 0, CREATE_SUSPENDED, 0&, 0&, SI, m_stUpgrade)
            Unload Me
        Case "Quit"
            For i = Forms.Count - 1 To 1 Step -1
                Unload Forms(i)
            Next i
        
            Unload Me
    End Select
End Sub


Private Sub trvMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim sTitle As String
    Dim sKey As Integer

    sKey = CInt(Mid(Node.Key, 2))
    If (sKey Mod 100) = 0 Then
        Node.Expanded = Not Node.Expanded
    Else
        Call RunForm(sKey)
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandler

    If KeyAscii = vbKeyReturn Then
        If Len(txtName) = 4 And IsNumeric(txtName) Then
            Call RunForm(txtName)
        End If
    End If
    Exit Sub
ErrHandler:
    If Err.Number = 35601 Then
        MsgBox LoadResString(113), vbCritical
        txtName = ""
        txtName.SetFocus
    Else
        MsgBox Err.Number & " : " & Err.Description
    End If
End Sub

Private Sub txtName_GotFocus()
    Call GotFocusText(txtName)
End Sub

'****************************************************************
'*Description:
'*  ADO�� �̿��Ͽ� Database�� �����ϱ�
'****************************************************************
Public Function ConnectDB() As Boolean
    Dim sConnect$

    On Error GoTo ErrHandler

    If g_adoCon Is Nothing Then
     
        sConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
         
       

'     sConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=" & _
'                        ";Initial Catalog=" & g_sDatabase & _
'                        ";Data Source=" & g_sServer & _
'                        ";Use Procedure for Prepare=1;Auto Translate=True;"

        Set g_adoCon = New ADODB.Connection
        With g_adoCon
            .CommandTimeout = 15
            .ConnectionString = sConnect
            .CursorLocation = adUseClient

            .Open sConnect
        End With
        ConnectDB = True
    ElseIf g_adoCon.State = adStateOpen Then
        ConnectDB = True
    Else
        ConnectDB = False
    End If
    
    Exit Function
ErrHandler:
    Unload frmSplash

    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection ����", True)

    ConnectDB = False
End Function

Public Sub SetFormCollection()
    If m_oForm Is Nothing Then
        Set m_oForm = New Collection
    Else
        Dim i%

        For i = 1 To m_oForm.Count
            Call m_oForm.Remove(1)
        Next i
    End If

    With m_oForm
        Call .Add(frmInfo, Format(1, FORMAT_MENUID))
        Call .Add(frmInfoSet, Format(2, FORMAT_MENUID))
'        Call .Add(frmLog, Format(3, FORMAT_MENUID))
        Call .Add(frmSetTerminal, Format(5, FORMAT_MENUID))
        'S_201312_��������_99 �� ���� �߰�
        Call .Add(frmSetting, Format(6, FORMAT_MENUID))     '�ڻ�����
        ' ----- [1000] �⺻ �ڵ� -----------------------------------------------------------------
        Call .Add(frmCustom, Format(1010, FORMAT_MENUID))            ' �ŷ�ó ����
        Call .Add(frmCustomUnit, Format(1020, FORMAT_MENUID))       ' �ŷ�ó �ܰ�
        
        ' ----- [1100] ǰ����� �ڵ� -----------------------------------------------------------------
        Call .Add(frmArticleUnit, Format(1110, FORMAT_MENUID))      ' ǰ�� �԰�
        Call .Add(frmArticleUnit, Format(1120, FORMAT_MENUID))      ' ǰ�� ����
        Call .Add(frmArticleUnit, Format(1130, FORMAT_MENUID))      ' ǰ�� ����
        Call .Add(frmArticleCode, Format(1140, FORMAT_MENUID))      ' ǰ�� ����
        Call .Add(frmArticleCodeAdmin, Format(1150, FORMAT_MENUID))      ' ǰ�� ����
        
        ' ----- [1200] ������� �ڵ� -----------------------------------------------------------------
        Call .Add(frmCommonCode, Format(1210, FORMAT_MENUID))       ' �μ� ����
        Call .Add(frmCommonCode, Format(1220, FORMAT_MENUID))       ' ��å ����
        Call .Add(frmPerson, Format(1230, FORMAT_MENUID))           ' ��� ����
        
        ' ----- [1300] ���ְ��� �ڵ� -----------------------------------------------------------------
        Call .Add(frmOrderCode, Format(1310, FORMAT_MENUID))        ' ����������
        Call .Add(frmOrderCode, Format(1320, FORMAT_MENUID))        ' �������� ����
        Call .Add(frmOrderCode, Format(1330, FORMAT_MENUID))        ' �������� ����
        Call .Add(frmOrderCode, Format(1340, FORMAT_MENUID))        ' ��屸�� ����
        Call .Add(frmOrderCode, Format(1350, FORMAT_MENUID))        ' �ֹ����� ����
        Call .Add(frmOrderCode, Format(1360, FORMAT_MENUID))        ' ���� ����
        
        ' ----- [1400] ���������� �ڵ� -----------------------------------------------------------------
        Call .Add(frmDyeAux, Format(1410, FORMAT_MENUID))             ' ������ ����
        Call .Add(frmDyeAuxGroup, Format(1420, FORMAT_MENUID))        ' ������ �׷�
        Call .Add(frmDyeAuxSubul, Format(1430, FORMAT_MENUID))
                
         '---- [1500] �������� �ڵ� -----------------------------------------------------------------
        Call .Add(frmProcessCode, Format(1510, FORMAT_MENUID))          ' �񰡵��ڵ�
        Call .Add(frmProcessCode, Format(1520, FORMAT_MENUID))          ' ���ְ����ڵ�
        Call .Add(frmProcessCode, Format(1530, FORMAT_MENUID))          ' �������� �ڵ�
        Call .Add(frmProcessCode, Format(1540, FORMAT_MENUID))          ' �۾��� �ڵ�
        Call .Add(frmMachineCode, Format(1550, FORMAT_MENUID))          ' ����/�����ڵ�
        Call .Add(frmPatternCode, Format(1560, FORMAT_MENUID))          ' �������� ����
        Call .Add(frmCodeCode, Format(1570, FORMAT_MENUID))          ' �ڵ����(Ư��,���,�ҷ�,����)
        
        ' ----- [1600] �˻���� �ڵ� -----------------------------------------------------------------
        Call .Add(frmInspectCode, Format(1610, FORMAT_MENUID))      ' �ҷ� ����
        Call .Add(frmInspectCode, Format(1620, FORMAT_MENUID))      ' �˻���� ����
        Call .Add(frmInspectCode, Format(1630, FORMAT_MENUID))      ' ��� ����
        
         '---- [1700] ������ �ڵ� -----------------------------------------------------------------
'        Call .Add(frmOutCode, Format(1710, FORMAT_MENUID))          ' ����� ����
'        Call .Add(frmOutCode, Format(1720, FORMAT_MENUID))          ' ��ǰ���� ����
        
        ' ----- [2000] ���ְ��� -------------------------------------------------------------------
        Call .Add(frmOrder, Format(2010, FORMAT_MENUID))                '���ֵ��
        Call .Add(frmOrderClose, Format(2020, FORMAT_MENUID))
        Call .Add(frmOrderAcptView, Format(2030, FORMAT_MENUID))     ' ���ں� Order���� ����
'        Call .Add(frmOrderCustom, Format(2030, FORMAT_MENUID))
'        Call .Add(frmOrderArticle, Format(2040, FORMAT_MENUID))
        
        ' ----- [2100] ���ܰ��� -------------------------------------------------------------------
        Call .Add(frmStuffIN, Format(2110, FORMAT_MENUID))
        Call .Add(frmStuffINView, Format(2120, FORMAT_MENUID))
        Call .Add(frmStuffINOrder, Format(2130, FORMAT_MENUID))
        Call .Add(frmStuffINList, Format(2140, FORMAT_MENUID))
        Call .Add(frmStuffINReturn, Format(2150, FORMAT_MENUID))
        
        ' ----- [3000] ������� -------------------------------------------------------------------
        Call .Add(frmPlanInput, Format(3010, FORMAT_MENUID))
        Call .Add(frmPlanInputView, Format(3020, FORMAT_MENUID))
        Call .Add(frmPlanCPB, Format(3050, FORMAT_MENUID))
        Call .Add(frmPlanCPBView, Format(3060, FORMAT_MENUID))
        
        Call .Add(frmInstRapid_NEW, Format(3130, FORMAT_MENUID))
'        Call .Add(frmInstRapid, Format(3130, FORMAT_MENUID))
        Call .Add(frmInstCondition, Format(3150, FORMAT_MENUID))
        
'        Call .Add(frmInstRapid, Format(3130, FORMAT_MENUID))
        
        Call .Add(frmResultDayByProcess, Format(3210, FORMAT_MENUID))       '������Ȳ
        Call .Add(frmResultSaleExpect, Format(3220, FORMAT_MENUID))       '������Ȳ
        Call .Add(frmOrderHistory, Format(3230, FORMAT_MENUID))
        Call .Add(frmCardHistory, Format(3240, FORMAT_MENUID))
        Call .Add(frmResultProdDyeing, Format(3250, FORMAT_MENUID))
        
        ' ----- [4000] ���� ���� ---------------------------------------------------------------------
        Call .Add(frmCardChange, Format(4010, FORMAT_MENUID))               ' ����ī�� ����
        Call .Add(frmCardDivide, Format(4020, FORMAT_MENUID))               ' ����ī�� �и�
        Call .Add(frmCardPattern, Format(4030, FORMAT_MENUID))              ' ����ī�� �и�
        
        Call .Add(frmWorkUnit, Format(4110, FORMAT_MENUID))                 ' �۾����� �׷����
        
        Call .Add(frmHold, Format(4210, FORMAT_MENUID))               ' ����ó��
        
        Call .Add(frmMatchColorView, Format(4310, FORMAT_MENUID))           ' �������
        Call .Add(frmProcessResultView, Format(4320, FORMAT_MENUID))        ' ��������
        Call .Add(frmProcessResultModify, Format(4330, FORMAT_MENUID))      ' �������� ��ȸ & ����
        Call .Add(frmDyeResultView, Format(4340, FORMAT_MENUID))            ' ��������
        Call .Add(frmProcessResultTenter, Format(4350, FORMAT_MENUID))            ' ��������

        Call .Add(frmProcessWait, Format(4410, FORMAT_MENUID))              ' ���� ��� ��ȸ
        Call .Add(frmProcWorking, Format(4420, FORMAT_MENUID))              ' ������ �۾� ��Ȳ
        Call .Add(frmProcWaiting, Format(4430, FORMAT_MENUID))              ' ������ ������ ��Ȳ
        Call .Add(frmProcessResultMgr, Format(4440, FORMAT_MENUID))         ' ����ī�� ������Ȳ
        ' ----- [5000] ����� ���� ---------------------------------------------------------------------
        Call .Add(frmBT, Format(5010, FORMAT_MENUID))               'B/T���
        Call .Add(frmBTCalc, Format(5020, FORMAT_MENUID))           'B/T ó���ۼ�
        Call .Add(frmBTView, Format(5030, FORMAT_MENUID))           'B/T��ȸ
        
        Call .Add(frmRecipe, Format(5110, FORMAT_MENUID))               'ó����
        Call .Add(frmRecipeView, Format(5120, FORMAT_MENUID))           'ó����
        Call .Add(frmModiRecipe, Format(5130, FORMAT_MENUID))           '���� ó����
        
        Call .Add(frmRecipeCalc, Format(5210, FORMAT_MENUID))           'ó����
        Call .Add(frmRecipeCalcView, Format(5220, FORMAT_MENUID))       'ó����
        
        ' ----- [7000] �˻���� -------------------------------------------------------------------
        Call .Add(frmInspectResultByOrder, Format(7010, FORMAT_MENUID)) ' ���ֺ� �˻� ��� ��ȸ
        Call .Add(frmInspectResultByDate, Format(7020, FORMAT_MENUID))  ' ���� �˻� ��� ��ȸ
        Call .Add(frmInspectDefectTotal, Format(7030, FORMAT_MENUID))   ' ���ں� �ҷ� ��Ȳ
        Call .Add(frmInspectResultByLot, Format(7040, FORMAT_MENUID))  ' LotNo�� �˻� ��� ��ȸ
        
        Call .Add(frmInspectDate, Format(7110, FORMAT_MENUID))          ' �˻� �Ϻ� ��ȸ
        Call .Add(frmInspect, Format(7120, FORMAT_MENUID))              ' �˻� ���� ��ȸ
        
        ' ----- [8000] ������ -------------------------------------------------------------------
        Call .Add(frmOutwareWork, Format(8010, FORMAT_MENUID))          ' ��� �۾�(Handy Terminal)
        Call .Add(frmOutwareIns, Format(8020, FORMAT_MENUID))           ' ��� ����
        Call .Add(frmOutware, Format(8030, FORMAT_MENUID))              ' ��� ����
        Call .Add(frmOutwareLot, Format(8040, FORMAT_MENUID))           ' Lot�� �����
        Call .Add(frmOutWareView, Format(8050, FORMAT_MENUID))          ' ��ǰ�����Ȳ
        
        Call .Add(frmOutwareDetail, Format(8110, FORMAT_MENUID))        ' ��� ���� ��ȸ

        ' ----- [8000] ������ ���� ---------------------------------------------------------------------
 '       Call .Add(frmDyeAux, Format(8010, FORMAT_MENUID))           ' ������ �ڵ� ����
'        Call .Add(frmDyeAuxIN, Format(8020, FORMAT_MENUID))         ' ������ ���� ����
'        Call .Add(frmDyeAuxOut, Format(8030, FORMAT_MENUID))        ' ������ ��� ����
        ' ----- [9000] �濵���� ---------------------------------------------------------------------
        Call .Add(frmControlOutWare, Format(9010, FORMAT_MENUID))           ' �������
        Call .Add(frmControlStock, Format(9020, FORMAT_MENUID))             ' ����Է�
        Call .Add(frmProcCost, Format(9030, FORMAT_MENUID))                 ' ���ó��
        
        Call .Add(frmSubulReport, Format(9110, FORMAT_MENUID))             ' ���Ҹ���
        Call .Add(frmStockReport, Format(9120, FORMAT_MENUID))             ' ������
        Call .Add(frmProcCostReport, Format(9130, FORMAT_MENUID))          ' û����
        Call .Add(frmProcCostCustom, Format(9140, FORMAT_MENUID))          ' ������ ����ǥ
        Call .Add(frmDeliverySaleReport, Format(9150, FORMAT_MENUID))      ' ����� ������ �ŵ� Ȯ�༭
'        Call .Add(frmContract, Format(9160, FORMAT_MENUID))               ' ���⹰ǰ�Ӱ�����༭
        Call .Add(frmOutWareReport, Format(9170, FORMAT_MENUID))           ' ���⹰ǰ�Ӱ�����༭
        Call .Add(frmSubulOrder, Format(9180, FORMAT_MENUID))              ' Order�� ���Ҹ���
        Call .Add(frmTaxList, Format(9190, FORMAT_MENUID))                 ' ��꼭 �߱� ��Ȳ

        
    End With
End Sub

Public Sub RunForm(Index As Integer)
    Dim sCaption$, sKey$, sNodeKey$, oForm As Form

    On Error Resume Next

    sKey = Format(Index, FORMAT_MENUID)
    txtName = sKey
    sNodeKey = "K" & sKey

    trvMenu.Nodes(sNodeKey).Selected = True

    Set oForm = m_oForm(sKey)

    oForm.Tag = sKey
    Call SetPermision(oForm)

    sCaption = "[" & sKey & "] " & Mid(trvMenu.Nodes(sNodeKey).Text, 1, Len(trvMenu.Nodes(sNodeKey).Text) - 6) & _
                " (�� " & oForm.Name & ")"

    oForm.optMain((Index Mod 100) / 10 - 1) = True
    oForm.tabForm.Tab = (Index Mod 100) / 10 - 1
    Call ShowForm(oForm, sCaption)

    Set oForm = Nothing
End Sub


'S_201312_��������_99 �� ���� �߰�
'****************************************************************
'*Description:
'*  ADO�� �̿��Ͽ� ������ �캯��ȣ Database�� �����ϱ�
'****************************************************************
Public Function ConnectWizDB() As Boolean
    
    Dim sWizConnect$

    On Error GoTo ErrHandler

''    If g_adoWizCon Is Nothing Then
        
        If Command() <> "" Then
            '//�׽���
           ' MsgBox "DB Test �ӽ�"
          '  g_sServer = "wizis.iptime.org,1433"
          '  g_sDatabase = "ZipDB"

            If g_sWizSQLAuthType = "1" Then
                
                                'SQL����
                sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & g_sWizSQLID & ";Password=" & g_sWizPassword & _
                            ";Initial Catalog=" & g_sWizDatabase & _
                            ";Data Source=" & g_sWizServer & _
                            ";Use Procedure for Prepare=1;Auto Translate=True;"
                
            Else
                '����������
                sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sWizServer & ";DATABASE=" & g_sWizDatabase & ";UID=sa;PWD=;"
            End If



        Else

            If g_sWizSQLAuthType = "1" Then
                'SQL����
                sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & g_sWizSQLID & ";Password=" & g_sWizPassword & _
                       ";Initial Catalog=" & g_sWizDatabase & _
                       ";Data Source=" & g_sWizServer & _
                       ";Use Procedure for Prepare=1;Auto Translate=True;"
            Else
         
                '����������
                sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sWizServer & ";DATABASE=" & g_sWizDatabase & ";UID=sa;PWD=;"
            End If
        End If

        Set g_adoWizCon = New ADODB.Connection
        With g_adoWizCon
            .CommandTimeout = 60
            .ConnectionString = sWizConnect
            .CursorLocation = adUseClient
            .Open sWizConnect
        End With


        If g_adoWizCon.State = adStateOpen Then
            ConnectWizDB = True
        Else
            
            ConnectWizDB = False
    ''        Set g_adoWizCon = Nothing
        End If
    
    Exit Function
ErrHandler:
''    Unload frmSplash
    Set g_adoWizCon = Nothing
''    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection ����", True)

    ConnectWizDB = False
End Function


