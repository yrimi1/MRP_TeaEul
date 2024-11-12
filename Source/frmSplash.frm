VERSION 5.00
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGif.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   4185
   ClientLeft      =   4650
   ClientTop       =   2385
   ClientWidth     =   7875
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00800000&
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame MainFrame 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H8000000C&
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   7800
      Begin AniGIFCtrl.AniGIF AniGIF 
         Height          =   945
         Left            =   5535
         TabIndex        =   10
         Top             =   210
         Width           =   2310
         BackColor       =   12632256
         PLaying         =   -1  'True
         Transparent     =   -1  'True
         Speed           =   1
         Stretch         =   0
         AutoSize        =   0   'False
         SequenceString  =   ""
         Sequence        =   0
         HTTPProxy       =   ""
         HTTPUserName    =   ""
         HTTPPassword    =   ""
         MousePointer    =   0
         ExtendWidth     =   4075
         ExtendHeight    =   1667
      End
      Begin VB.CommandButton cmdInformation 
         Caption         =   "시스템 정보 ..."
         Height          =   555
         Left            =   5745
         MousePointer    =   99  '사용자 정의
         TabIndex        =   7
         ToolTipText     =   "확인"
         Top             =   2805
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "확  인 (&O)"
         Default         =   -1  'True
         Height          =   555
         Left            =   5745
         MousePointer    =   99  '사용자 정의
         TabIndex        =   4
         ToolTipText     =   "확인"
         Top             =   3450
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  '평면
         Height          =   720
         Left            =   195
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '투명
         Caption         =   "Copyright : ⓒ2000~2013"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   3000
         TabIndex        =   11
         Tag             =   "Copyright"
         Top             =   3810
         Width           =   2355
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1275
         Width           =   180
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   285
         TabIndex        =   9
         Top             =   1290
         Width           =   180
      End
      Begin VB.Image imgBackGround 
         BorderStyle     =   1  '단일 고정
         Height          =   2985
         Left            =   90
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   5505
      End
      Begin VB.Label lblLicenseTo 
         BackStyle       =   0  '투명
         Caption         =   "(주) 위저드            정보시스템"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   6045
         TabIndex        =   6
         Tag             =   "LicenseTo"
         Top             =   1650
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "LicenseTo :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   5745
         TabIndex        =   5
         Tag             =   "LicenseTo"
         Top             =   1320
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProjectName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   21.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   1
         Left            =   1035
         TabIndex        =   3
         Tag             =   "Product"
         Top             =   270
         Width           =   255
      End
      Begin VB.Label lblProjectName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   21.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   0
         Left            =   1065
         TabIndex        =   2
         Tag             =   "Product"
         Top             =   285
         Width           =   255
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Tag             =   "Version"
         Top             =   750
         Width           =   885
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  '투명
         Caption         =   "Copyright : ⓒ2000~2013"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   3030
         TabIndex        =   12
         Tag             =   "Copyright"
         Top             =   3825
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const UPGRADE_FILE = "Upgrade.exe"

Private Sub cmdInformation_Click()
    Shell "C:\Program Files\Common Files\Microsoft Shared\MSinfo\Msinfo32.exe", vbNormalFocus
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    OneTime = True
    
    lblProjectName(0) = App.ProductName
    lblProjectName(1) = App.ProductName
    
    lblCompany(0) = g_companyInfo.Company_Name
    lblCompany(1) = g_companyInfo.Company_Name
    
    ' [1] Logo Image File----------------------------------------------------'
    'imgLogo.Picture = LoadResPicture("ACTIVE", vbResIcon)
    ' [2] BackGround Picture-------------------------------------------------'
    imgBackGround.Picture = LoadPicture(App.Path & "\Splash.gif")
    
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    ' [3] Animation Picture-------------------------------------------------'
    AniGIF.ReadGIF (App.Path & "\Welcome.gif")
    AniGIF.Play
    
    Call CheckUpgrade
    
    Call ExplodeForm(Me, 50)
    
    Dim lReturn As Long
    lReturn = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
    
    cmdOK.MouseIcon = LoadResPicture("POINTER", vbResCursor) 'Pointer Cursor
    cmdInformation.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
End Sub

Private Sub CheckUpgrade()
    Dim oFileSystem As Scripting.FileSystemObject
    Dim oFile As Scripting.File
    
    Dim sServerDate$
    
    Dim g_sUpgradePath$, sFilePath$
    
    On Error GoTo ErrHandler
    
    g_sUpgradePath = GetIniValue("Path", "Upgrade")
    If Len(g_sUpgradePath) > 0 Then
        'Call SetUpgradeFolder
        g_sUpgradePath = g_sUpgradePath & "\..\"
        
        Set oFileSystem = New Scripting.FileSystemObject
        Set oFile = oFileSystem.GetFile(g_sUpgradePath & UPGRADE_FILE)
        sServerDate = oFile.DateLastModified
        
        
        sFilePath = App.Path & "\" & UPGRADE_FILE
        If oFileSystem.FileExists(sFilePath) Then
            Set oFile = oFileSystem.GetFile(sFilePath)
            If oFile.DateLastModified < sServerDate Then
                oFileSystem.CopyFile g_sUpgradePath & UPGRADE_FILE, sFilePath, True
            End If
        Else
            oFileSystem.CopyFile g_sUpgradePath & UPGRADE_FILE, sFilePath, True
        End If
    End If
    
ErrHandler:
    Set oFile = Nothing
    Set oFileSystem = Nothing
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ImplodeForm(Me, 50)

    imgBackGround.Picture = LoadPicture()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash = Nothing
End Sub

