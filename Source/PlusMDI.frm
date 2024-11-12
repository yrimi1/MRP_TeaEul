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
   WindowState     =   2  '최대화
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   6495
      Top             =   2565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar MainStatus 
      Align           =   2  '아래 맞춤
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
            Object.ToolTipText     =   "작업일자"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "오후 5:34"
            Object.ToolTipText     =   "현재 시간"
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pnlMenu 
      Align           =   3  '왼쪽 맞춤
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
         Name            =   "굴림"
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
            Name            =   "굴림"
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
            Name            =   "굴림"
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
         Caption         =   "메뉴 목록"
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
         BorderStyle     =   0  '없음
         FillColor       =   &H00808080&
         Height          =   7515
         Left            =   3090
         ScaleHeight     =   3272.354
         ScaleMode       =   0  '사용자
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
         MousePointer    =   9  'W E 크기 조정
         Top             =   -15
         Width           =   105
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  '위 맞춤
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
         Caption         =   "  화면 번호 (F12)"
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
      Caption         =   "시스템(&S)"
      Begin VB.Menu mnuScreen 
         Caption         =   "화면 번호"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuMenu 
         Caption         =   "메뉴 목록"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTools 
         Caption         =   "도구 목록"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSP0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "로그 인"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "로그 아웃"
      End
      Begin VB.Menu mnuSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "임호 변경"
      End
      Begin VB.Menu mnuSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinterSet 
         Caption         =   "프린터 설정"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSP4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "프린터"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "미리보기"
      End
      Begin VB.Menu mnuSP5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDirect 
         Caption         =   "바로인쇄"
      End
   End
   Begin VB.Menu mnuErase 
      Caption         =   "파일 삭제"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "삭제"
      End
      Begin VB.Menu mnuSP6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "취소"
      End
   End
End
Attribute VB_Name = "PlusMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************************************
' 변경이력
'------------------------------------------------------------------------------
'
'요청ID : S_201203_태을염직_02
'요청일자 : 2012.03.05
'요청내용 : 오더별 명세 출력되게
'변경내용 : Gf_DB_CM_GetCompanyInfo 추가
'
'  요청사항 ID: S_201312_태을염직_99
'  요청자:
'  변경날짜 : 2013.12.12
'  작업자   : 오승욱
'  요청내용 : 지번주소에서 도로명 주소로 입력가능하게
'  변경내용 : 도로명,구 지번주소 옵션 버튼 추가
'******************************************************************************

Option Explicit

Private Const MAX_SPLIT As Integer = 2500
Private Const MAX_MDIH  As Integer = 12000
Private Const MAX_MDIV  As Integer = 9000

Private m_bMoving  As Boolean ' Splitter 사용에 사용 될 변수
Private m_bPreview As Boolean ' Print Preview에 사용 될 변수

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
    
    'S_201312_태을염직_99 에 의한 수정-Start.Bas에 정의함
''    'S_201203_태을염직_02 에 의한 추가
''    '-------------------------------------
''    '업체정보 Get
''    '-------------------------------------
''    If g_companyInfo.Company_Name = "" Then
''        If Gf_DB_CM_GetCompanyInfo(rs, "Y") = True Then
''
''            If rs.EOF = False Then
''                g_companyInfo.Company_Name = Trim(CheckNull(rs!KCompany))    '상호
''                g_companyInfo.Chief = Trim(CheckNull(rs!Chief))                  '대표자명
''                g_companyInfo.Address1 = Trim(CheckNull(rs!Address1))            '주소1
''                g_companyInfo.Address2 = Trim(CheckNull(rs!Address2))            '주소2
''                g_companyInfo.Company_type = Trim(CheckNull(rs!Condition))    '업태
''                g_companyInfo.Category = Trim(CheckNull(rs!Category))            '업종
''                g_companyInfo.Company_No = Trim(CheckNull(rs!CompanyNo))        '사업자번호
''            End If
''        End If
''    End If
    
    
    ' 이미지 리스트에 아이콘 생성 (트리메뉴에 사용)
    With imgTree
        .ListImages.Add Key:="Unfolder", Picture:=LoadResPicture("UNFOLDER", vbResIcon)
        .ListImages.Add Key:="Folder", Picture:=LoadResPicture("FOLDER", vbResIcon)
        .ListImages.Add Key:="Close", Picture:=LoadResPicture("CLOSE", vbResIcon)
        .ListImages.Add Key:="Open", Picture:=LoadResPicture("OPEN", vbResIcon)
        .ListImages.Add Key:="Blank", Picture:=LoadResPicture("BLANK", vbResIcon)
        .ListImages.Add Key:="Check", Picture:=LoadResPicture("CHECK", vbResIcon)
    End With
    
    '이미지 리스트에 아이콘 생성 (툴바에 사용)
    With imlTool
        .ListImages.Add Key:="Back", Picture:=LoadResPicture("BACK", vbResIcon)
        .ListImages.Add Key:="Front", Picture:=LoadResPicture("FRONT", vbResIcon)
        .ListImages.Add Key:="Monitor", Picture:=LoadResPicture("MONITOR", vbResIcon)
        .ListImages.Add Key:="Menu", Picture:=LoadResPicture("MENU", vbResIcon)
        .ListImages.Add Key:="Quit", Picture:=LoadResPicture("QUIT", vbResIcon)
        .ListImages.Add Key:="Close", Picture:=LoadResPicture("FOLDER", vbResIcon)
    End With

    '툴바(Toolbar) 설정
    With tbrMain
        .Buttons.Add Key:="Back", Caption:="뒤로", Style:=tbrDefault, Image:="Back"
        .Buttons.Add Key:="Front", Caption:="앞으로", Style:=tbrDefault, Image:="Front"
        .Buttons.Add Style:=tbrSeparator
        .Buttons.Add Key:="Upgrade", Caption:="자동업그레이드", Style:=tbrDefault, Image:="Monitor"
        .Buttons.Add Style:=tbrSeparator
        .Buttons.Add Key:="Menu", Caption:="메뉴목록", Style:=tbrCheck, Image:="Menu"
        .Buttons.Add Style:=tbrSeparator
        .Buttons.Add Key:="Close", Caption:="모두닫기", Style:=tbrDefault, Image:="Close"
        .Buttons.Add Key:="Quit", Caption:="종료 ", Style:=tbrDefault, Image:="Quit"
        .Buttons("Menu").Value = tbrPressed

        .MouseIcon = LoadResPicture("POINTER", vbResCursor)
    End With
    cmdSize(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdSize(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    trvMenu.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    Call SetFormCollection
    Call FirstFormLoad
    
    '//Hokk오류 발생 vbmode로 테스트
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
   
    
    '//Hokk오류 발생 vbmode로 테스트
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
'*Description: 동적 메뉴 구성
'*  TreeView Control과 상단 메뉴를 구성한다.
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

    AppendMenu lMenuHwnd, MF_POPUP, lSubHwnd, "도움말(&H)"
    AppendMenu lSubHwnd, MF_STRING, 1, "MRPPlus2에 대하여 ....(&A)"

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
'    AppendMenu lMenuHwnd, MF_POPUP, lSubHwnd, "도움말(&H)"
'    AppendMenu lSubHwnd, MF_STRING, 1, "MRP Plus에 대하여 ....(&A)"
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
    
    '너비를 설정합니다.
    If X < MAX_SPLIT Then X = MAX_SPLIT
    If X > (Me.Width - MAX_SPLIT) Then X = Me.Width - MAX_SPLIT

    imgSplitter.Left = trvMenu.Left + trvMenu.Width - 30
    imgSplitter.Height = pnlMenu.Height
End Sub

Private Sub mnuLogin_Click()
    Call ConnectDB
    
    frmLogin.Show vbModal 'Login Form을 Load후 UserID와 Passord를 Check 함.
    
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
    '*   메뉴 바 (MenuBar) 삭제
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
'*  ADO를 이용하여 Database에 접속하기
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

    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection 실패", True)

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
        'S_201312_태을염직_99 에 의한 추가
        Call .Add(frmSetting, Format(6, FORMAT_MENUID))     '자사정보
        ' ----- [1000] 기본 코드 -----------------------------------------------------------------
        Call .Add(frmCustom, Format(1010, FORMAT_MENUID))            ' 거래처 관리
        Call .Add(frmCustomUnit, Format(1020, FORMAT_MENUID))       ' 거래처 단가
        
        ' ----- [1100] 품명관련 코드 -----------------------------------------------------------------
        Call .Add(frmArticleUnit, Format(1110, FORMAT_MENUID))      ' 품명 규격
        Call .Add(frmArticleUnit, Format(1120, FORMAT_MENUID))      ' 품명 색상
        Call .Add(frmArticleUnit, Format(1130, FORMAT_MENUID))      ' 품명 사종
        Call .Add(frmArticleCode, Format(1140, FORMAT_MENUID))      ' 품명 관리
        Call .Add(frmArticleCodeAdmin, Format(1150, FORMAT_MENUID))      ' 품명 관리
        
        ' ----- [1200] 사원관련 코드 -----------------------------------------------------------------
        Call .Add(frmCommonCode, Format(1210, FORMAT_MENUID))       ' 부서 관리
        Call .Add(frmCommonCode, Format(1220, FORMAT_MENUID))       ' 직책 관리
        Call .Add(frmPerson, Format(1230, FORMAT_MENUID))           ' 사원 관리
        
        ' ----- [1300] 수주관련 코드 -----------------------------------------------------------------
        Call .Add(frmOrderCode, Format(1310, FORMAT_MENUID))        ' 원단폭관리
        Call .Add(frmOrderCode, Format(1320, FORMAT_MENUID))        ' 가공구분 관리
        Call .Add(frmOrderCode, Format(1330, FORMAT_MENUID))        ' 레벨구분 관리
        Call .Add(frmOrderCode, Format(1340, FORMAT_MENUID))        ' 밴드구분 관리
        Call .Add(frmOrderCode, Format(1350, FORMAT_MENUID))        ' 주문형태 관리
        Call .Add(frmOrderCode, Format(1360, FORMAT_MENUID))        ' 필장 관리
        
        ' ----- [1400] 염조제관련 코드 -----------------------------------------------------------------
        Call .Add(frmDyeAux, Format(1410, FORMAT_MENUID))             ' 염조제 관리
        Call .Add(frmDyeAuxGroup, Format(1420, FORMAT_MENUID))        ' 염조제 그룹
        Call .Add(frmDyeAuxSubul, Format(1430, FORMAT_MENUID))
                
         '---- [1500] 공정관련 코드 -----------------------------------------------------------------
        Call .Add(frmProcessCode, Format(1510, FORMAT_MENUID))          ' 비가동코드
        Call .Add(frmProcessCode, Format(1520, FORMAT_MENUID))          ' 외주가공코드
        Call .Add(frmProcessCode, Format(1530, FORMAT_MENUID))          ' 건조구분 코드
        Call .Add(frmProcessCode, Format(1540, FORMAT_MENUID))          ' 작업조 코드
        Call .Add(frmMachineCode, Format(1550, FORMAT_MENUID))          ' 공정/설비코드
        Call .Add(frmPatternCode, Format(1560, FORMAT_MENUID))          ' 공정패턴 관리
        Call .Add(frmCodeCode, Format(1570, FORMAT_MENUID))          ' 코드관리(특기,비고,불량,보류)
        
        ' ----- [1600] 검사관련 코드 -----------------------------------------------------------------
        Call .Add(frmInspectCode, Format(1610, FORMAT_MENUID))      ' 불량 관리
        Call .Add(frmInspectCode, Format(1620, FORMAT_MENUID))      ' 검사기준 관리
        Call .Add(frmInspectCode, Format(1630, FORMAT_MENUID))      ' 등급 관리
        
         '---- [1700] 출고관련 코드 -----------------------------------------------------------------
'        Call .Add(frmOutCode, Format(1710, FORMAT_MENUID))          ' 출고구분 관리
'        Call .Add(frmOutCode, Format(1720, FORMAT_MENUID))          ' 반품구분 관리
        
        ' ----- [2000] 수주관리 -------------------------------------------------------------------
        Call .Add(frmOrder, Format(2010, FORMAT_MENUID))                '수주등록
        Call .Add(frmOrderClose, Format(2020, FORMAT_MENUID))
        Call .Add(frmOrderAcptView, Format(2030, FORMAT_MENUID))     ' 일자별 Order접수 명세서
'        Call .Add(frmOrderCustom, Format(2030, FORMAT_MENUID))
'        Call .Add(frmOrderArticle, Format(2040, FORMAT_MENUID))
        
        ' ----- [2100] 원단관리 -------------------------------------------------------------------
        Call .Add(frmStuffIN, Format(2110, FORMAT_MENUID))
        Call .Add(frmStuffINView, Format(2120, FORMAT_MENUID))
        Call .Add(frmStuffINOrder, Format(2130, FORMAT_MENUID))
        Call .Add(frmStuffINList, Format(2140, FORMAT_MENUID))
        Call .Add(frmStuffINReturn, Format(2150, FORMAT_MENUID))
        
        ' ----- [3000] 생산관리 -------------------------------------------------------------------
        Call .Add(frmPlanInput, Format(3010, FORMAT_MENUID))
        Call .Add(frmPlanInputView, Format(3020, FORMAT_MENUID))
        Call .Add(frmPlanCPB, Format(3050, FORMAT_MENUID))
        Call .Add(frmPlanCPBView, Format(3060, FORMAT_MENUID))
        
        Call .Add(frmInstRapid_NEW, Format(3130, FORMAT_MENUID))
'        Call .Add(frmInstRapid, Format(3130, FORMAT_MENUID))
        Call .Add(frmInstCondition, Format(3150, FORMAT_MENUID))
        
'        Call .Add(frmInstRapid, Format(3130, FORMAT_MENUID))
        
        Call .Add(frmResultDayByProcess, Format(3210, FORMAT_MENUID))       '생산현황
        Call .Add(frmResultSaleExpect, Format(3220, FORMAT_MENUID))       '생산현황
        Call .Add(frmOrderHistory, Format(3230, FORMAT_MENUID))
        Call .Add(frmCardHistory, Format(3240, FORMAT_MENUID))
        Call .Add(frmResultProdDyeing, Format(3250, FORMAT_MENUID))
        
        ' ----- [4000] 공정 관리 ---------------------------------------------------------------------
        Call .Add(frmCardChange, Format(4010, FORMAT_MENUID))               ' 공정카드 관리
        Call .Add(frmCardDivide, Format(4020, FORMAT_MENUID))               ' 공정카드 분리
        Call .Add(frmCardPattern, Format(4030, FORMAT_MENUID))              ' 공정카드 분리
        
        Call .Add(frmWorkUnit, Format(4110, FORMAT_MENUID))                 ' 작업단위 그룹관리
        
        Call .Add(frmHold, Format(4210, FORMAT_MENUID))               ' 보류처리
        
        Call .Add(frmMatchColorView, Format(4310, FORMAT_MENUID))           ' 배색일지
        Call .Add(frmProcessResultView, Format(4320, FORMAT_MENUID))        ' 공정일지
        Call .Add(frmProcessResultModify, Format(4330, FORMAT_MENUID))      ' 공정일지 조회 & 발행
        Call .Add(frmDyeResultView, Format(4340, FORMAT_MENUID))            ' 염색일지
        Call .Add(frmProcessResultTenter, Format(4350, FORMAT_MENUID))            ' 가공일지

        Call .Add(frmProcessWait, Format(4410, FORMAT_MENUID))              ' 공정 대기 조회
        Call .Add(frmProcWorking, Format(4420, FORMAT_MENUID))              ' 공정별 작업 현황
        Call .Add(frmProcWaiting, Format(4430, FORMAT_MENUID))              ' 공정별 예상대기 현황
        Call .Add(frmProcessResultMgr, Format(4440, FORMAT_MENUID))         ' 공정카드 진행현황
        ' ----- [5000] 실험실 관리 ---------------------------------------------------------------------
        Call .Add(frmBT, Format(5010, FORMAT_MENUID))               'B/T등록
        Call .Add(frmBTCalc, Format(5020, FORMAT_MENUID))           'B/T 처방작성
        Call .Add(frmBTView, Format(5030, FORMAT_MENUID))           'B/T조회
        
        Call .Add(frmRecipe, Format(5110, FORMAT_MENUID))               '처방전
        Call .Add(frmRecipeView, Format(5120, FORMAT_MENUID))           '처방전
        Call .Add(frmModiRecipe, Format(5130, FORMAT_MENUID))           '수정 처방전
        
        Call .Add(frmRecipeCalc, Format(5210, FORMAT_MENUID))           '처방전
        Call .Add(frmRecipeCalcView, Format(5220, FORMAT_MENUID))       '처방전
        
        ' ----- [7000] 검사실적 -------------------------------------------------------------------
        Call .Add(frmInspectResultByOrder, Format(7010, FORMAT_MENUID)) ' 수주별 검사 결과 조회
        Call .Add(frmInspectResultByDate, Format(7020, FORMAT_MENUID))  ' 색상별 검사 결과 조회
        Call .Add(frmInspectDefectTotal, Format(7030, FORMAT_MENUID))   ' 일자별 불량 현황
        Call .Add(frmInspectResultByLot, Format(7040, FORMAT_MENUID))  ' LotNo별 검사 결과 조회
        
        Call .Add(frmInspectDate, Format(7110, FORMAT_MENUID))          ' 검사 일보 조회
        Call .Add(frmInspect, Format(7120, FORMAT_MENUID))              ' 검사 실적 조회
        
        ' ----- [8000] 출고관리 -------------------------------------------------------------------
        Call .Add(frmOutwareWork, Format(8010, FORMAT_MENUID))          ' 출고 작업(Handy Terminal)
        Call .Add(frmOutwareIns, Format(8020, FORMAT_MENUID))           ' 출고 관리
        Call .Add(frmOutware, Format(8030, FORMAT_MENUID))              ' 출고 관리
        Call .Add(frmOutwareLot, Format(8040, FORMAT_MENUID))           ' Lot별 출고내역
        Call .Add(frmOutWareView, Format(8050, FORMAT_MENUID))          ' 제품출고현황
        
        Call .Add(frmOutwareDetail, Format(8110, FORMAT_MENUID))        ' 출고 실적 조회

        ' ----- [8000] 염조제 관리 ---------------------------------------------------------------------
 '       Call .Add(frmDyeAux, Format(8010, FORMAT_MENUID))           ' 염조제 코드 관리
'        Call .Add(frmDyeAuxIN, Format(8020, FORMAT_MENUID))         ' 염조제 매입 관리
'        Call .Add(frmDyeAuxOut, Format(8030, FORMAT_MENUID))        ' 염조제 사용 내역
        ' ----- [9000] 경영정보 ---------------------------------------------------------------------
        Call .Add(frmControlOutWare, Format(9010, FORMAT_MENUID))           ' 출고조정
        Call .Add(frmControlStock, Format(9020, FORMAT_MENUID))             ' 재고입력
        Call .Add(frmProcCost, Format(9030, FORMAT_MENUID))                 ' 계상처리
        
        Call .Add(frmSubulReport, Format(9110, FORMAT_MENUID))             ' 수불명세서
        Call .Add(frmStockReport, Format(9120, FORMAT_MENUID))             ' 재고명세서
        Call .Add(frmProcCostReport, Format(9130, FORMAT_MENUID))          ' 청구서
        Call .Add(frmProcCostCustom, Format(9140, FORMAT_MENUID))          ' 가공료 집계표
        Call .Add(frmDeliverySaleReport, Format(9150, FORMAT_MENUID))      ' 수출용 원자재 매도 확약서
'        Call .Add(frmContract, Format(9160, FORMAT_MENUID))               ' 수출물품임가공계약서
        Call .Add(frmOutWareReport, Format(9170, FORMAT_MENUID))           ' 수출물품임가공계약서
        Call .Add(frmSubulOrder, Format(9180, FORMAT_MENUID))              ' Order불 수불명세서
        Call .Add(frmTaxList, Format(9190, FORMAT_MENUID))                 ' 계산서 발급 현황

        
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
                " (→ " & oForm.Name & ")"

    oForm.optMain((Index Mod 100) / 10 - 1) = True
    oForm.tabForm.Tab = (Index Mod 100) / 10 - 1
    Call ShowForm(oForm, sCaption)

    Set oForm = Nothing
End Sub


'S_201312_태을염직_99 에 의한 추가
'****************************************************************
'*Description:
'*  ADO를 이용하여 위저드 우변번호 Database에 접속하기
'****************************************************************
Public Function ConnectWizDB() As Boolean
    
    Dim sWizConnect$

    On Error GoTo ErrHandler

''    If g_adoWizCon Is Nothing Then
        
        If Command() <> "" Then
            '//테스용
           ' MsgBox "DB Test 임시"
          '  g_sServer = "wizis.iptime.org,1433"
          '  g_sDatabase = "ZipDB"

            If g_sWizSQLAuthType = "1" Then
                
                                'SQL인증
                sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & g_sWizSQLID & ";Password=" & g_sWizPassword & _
                            ";Initial Catalog=" & g_sWizDatabase & _
                            ";Data Source=" & g_sWizServer & _
                            ";Use Procedure for Prepare=1;Auto Translate=True;"
                
            Else
                '윈도우인증
                sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sWizServer & ";DATABASE=" & g_sWizDatabase & ";UID=sa;PWD=;"
            End If



        Else

            If g_sWizSQLAuthType = "1" Then
                'SQL인증
                sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & g_sWizSQLID & ";Password=" & g_sWizPassword & _
                       ";Initial Catalog=" & g_sWizDatabase & _
                       ";Data Source=" & g_sWizServer & _
                       ";Use Procedure for Prepare=1;Auto Translate=True;"
            Else
         
                '윈도우인증
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
''    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection 실패", True)

    ConnectWizDB = False
End Function


