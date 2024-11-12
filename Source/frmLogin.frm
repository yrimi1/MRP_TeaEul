VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   4710
   ClientLeft      =   6075
   ClientTop       =   1740
   ClientWidth     =   4755
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취 소 (&C)"
      Height          =   390
      Left            =   3210
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   3735
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "로그인 (&O)"
      Default         =   -1  'True
      Height          =   390
      Left            =   3210
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Frame fraGroup 
      Height          =   1005
      Left            =   30
      TabIndex        =   0
      Top             =   3150
      Width           =   3105
      Begin VB.TextBox txtPassWd 
         Height          =   300
         IMEMode         =   3  '사용 못함
         Left            =   1530
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   570
         Width           =   1470
      End
      Begin VB.TextBox txtUserID 
         Height          =   300
         IMEMode         =   8  '영문
         Left            =   1530
         TabIndex        =   1
         Top             =   210
         Width           =   1470
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   225
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "사용자 번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Top             =   585
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "비밀 번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  '단일 고정
      Height          =   3120
      Left            =   30
      Stretch         =   -1  'True
      Top             =   30
      Width           =   4695
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "Copyright (C) 2000 Wizard I.S. Corp."
      Height          =   180
      Index           =   0
      Left            =   795
      TabIndex        =   8
      Top             =   4230
      Width           =   3105
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "All Right are Reserved. "
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   4440
      Width           =   4590
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_nFailCount As Integer

Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Left = (Screen.Width - .Width) / 2

        .Caption = LoadResString(101)
        .Icon = LoadResPicture("KEY", vbResIcon)    ' 열쇠
    End With

    'imgLogo.Picture = LoadPicture(App.Path & "\Login.gif")

    cmdOK.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdCancel.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)

    txtUserID = GetSetting(LoadResString(100), "Login", "UserID")
End Sub

Private Sub Form_Activate()
    txtPassWd.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub cmdOK_Click()
    Dim oLogin As PlusLib2.CLogin
    Dim vDateTime As Variant

    cmdOK.Enabled = False

    On Error GoTo ErrHandle

    Set oLogin = New PlusLib2.CLogin
    g_sUserName = UCase(txtUserID)
    g_sPassword = txtPassWd

    oLogin.Connection = g_adoCon
    oLogin.UserName = g_sUserName

    g_sPersonName = oLogin.Login(g_sUserName, g_sPassword)

    Call SaveSetting(LoadResString(100), "Login", "UserID", g_sUserName)

    vDateTime = oLogin.GetNow()
    Date = vDateTime
    time = vDateTime

    PlusMDI.MainStatus.Panels(2).Text = g_sUserName & "(" & g_sPersonName & ")"
'    PlusMDI.MainStatus.Panels(2).Text = "사용자 : " & g_sUserName
    PlusMDI.MainStatus.Panels(3).Text = "작업일 : " & MakeDate(DF_LONG, Date)
    PlusMDI.MakeMenu g_sUserName

    Set oLogin = Nothing

    Unload Me

    On Error GoTo 0
    On Error Resume Next

    Screen.MousePointer = vbDefault
    PlusMDI.MainStatus.Panels(1).Text = ""

    Exit Sub

ErrHandle:
    Screen.MousePointer = vbDefault
    PlusMDI.MainStatus.Panels(1).Text = ""

    Set oLogin = Nothing

    m_nFailCount = m_nFailCount + 1
    If m_nFailCount > 3 Then
        Call ErrorBox(Err.Number, Err.Source, Err.Description, , True)
        End
    End If

    Call ErrorBox(Err.Number, Err.Source, Err.Description, "로그인 실패")
    txtUserID.SetFocus
    cmdOK.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub txtPassWd_GotFocus()
    Call GotFocusText(txtPassWd)
End Sub

Private Sub txtPassWd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtUserID_GotFocus()
    Call GotFocusText(txtUserID)
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub
