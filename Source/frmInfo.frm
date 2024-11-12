VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmInfo 
   ClientHeight    =   9255
   ClientLeft      =   3690
   ClientTop       =   3210
   ClientWidth     =   11850
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   0
      Top             =   8520
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlBorder 
      Height          =   2085
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   6405
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3678
      _Version        =   196609
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtInfo 
         Height          =   1455
         Index           =   1
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   2
         Top             =   540
         Width           =   11580
      End
      Begin Threed.SSPanel pnlName 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   105
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   741
         _Version        =   196609
         Caption         =   "사용자별 공지사항"
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlBorder 
      Height          =   6360
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   11218
      _Version        =   196609
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlName 
         Height          =   420
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   741
         _Version        =   196609
         Caption         =   "태을염직 알림사항"
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5760
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   5
         Top             =   510
         Width           =   11580
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Call GetInfo
'    Me.Move 0, 0, 11985, 9660
'
'    Call SetOperate(Me)
'    Call GetInfo
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)
    Call GetInfo
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub GetInfo()
    Dim oInfo  As PlusLib2.CInfo
    Dim rsMain As ADODB.Recordset
    Dim rsUser As ADODB.Recordset
    Dim sInfo$, i%

    On Error GoTo ErrHandler

    sInfo = "▶ 오늘은 " & MakeDate(DF_FULL, Date) & " 입니다."

    Set oInfo = New PlusLib2.CInfo
    oInfo.Connection = g_adoCon

    Set rsMain = oInfo.GetInfoByDate(MakeDate(DF_SHORT, Date), MakeDate(DF_SHORT, Date))
    Set rsUser = oInfo.GetInfoUserByUserID(MakeDate(DF_SHORT, Date), g_sUserName)

    Set oInfo = Nothing

    If Not rsMain.EOF Then
        txtInfo(0) = sInfo & vbCrLf & vbCrLf & "▶ " & rsMain!Info
    Else
        txtInfo(0) = sInfo
    End If
    rsMain.Close
    Set rsMain = Nothing

    If Not rsUser.EOF Then
        txtInfo(1) = "▶ " & rsUser!Info
    Else
        txtInfo(1) = ""
    End If
    rsUser.Close
    Set rsUser = Nothing

    Exit Sub

ErrHandler:
    Set oInfo = Nothing
    Set rsMain = Nothing
    Set rsUser = Nothing

    Call ErrorBox(Err.Number, "Info.GetInfo", Err.Description)
End Sub

