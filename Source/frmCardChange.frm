VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCardChange 
   ClientHeight    =   9315
   ClientLeft      =   5535
   ClientTop       =   6045
   ClientWidth     =   11850
   Icon            =   "frmCardChange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   11850
   Begin Crystal.CrystalReport cryReport 
      Left            =   360
      Top             =   8580
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "저장(&S)"
      Height          =   780
      Index           =   3
      Left            =   9420
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   42
      ToolTipText     =   "자료 저장"
      Top             =   1020
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "수정(&U)"
      Height          =   780
      Index           =   1
      Left            =   10995
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   47
      ToolTipText     =   "자료 수정"
      Top             =   1020
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "취소(&C)"
      Height          =   780
      Index           =   4
      Left            =   10215
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   46
      ToolTipText     =   "자료 취소"
      Top             =   1020
      Visible         =   0   'False
      Width           =   780
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10110
      TabIndex        =   24
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   480
      TabIndex        =   21
      Top             =   4770
      Visible         =   0   'False
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1535
      _Version        =   196609
      Alignment       =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin MSComctlLib.ProgressBar proProgress 
         Height          =   390
         Left            =   90
         TabIndex        =   22
         Top             =   375
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "180"
         Height          =   180
         Left            =   195
         TabIndex        =   23
         Top             =   120
         Width           =   270
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   5835
      Left            =   30
      TabIndex        =   20
      Top             =   2550
      Width           =   11835
      _cx             =   20876
      _cy             =   10292
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSFrame frmDetail 
      Height          =   1605
      Left            =   30
      TabIndex        =   19
      Top             =   930
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   2831
      _Version        =   196609
      Begin VB.TextBox txtStuffDensity 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   4860
         TabIndex        =   37
         Top             =   1230
         Width           =   1125
      End
      Begin VB.TextBox txtStuffWidth 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   4860
         TabIndex        =   36
         Top             =   870
         Width           =   1125
      End
      Begin VB.ComboBox cboStuffCustom 
         Height          =   300
         Left            =   7320
         TabIndex        =   39
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cboThreadName 
         Height          =   300
         Left            =   7320
         TabIndex        =   38
         Top             =   90
         Width           =   1935
      End
      Begin VB.TextBox txtOrderID 
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   90
         Width           =   1875
      End
      Begin VB.ComboBox cboUseClss 
         Height          =   300
         Left            =   1320
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   1230
         Width           =   2235
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   4860
         TabIndex        =   35
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox txtRoll 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   4860
         TabIndex        =   34
         Top             =   90
         Width           =   1125
      End
      Begin VB.TextBox txtOrderNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cboColor 
         Height          =   300
         Left            =   1320
         Style           =   2  '드롭다운 목록
         TabIndex        =   32
         Top             =   870
         Width           =   2235
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "관리번호"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "Order NO."
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   27
         Top             =   870
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "색상명"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   3
         Left            =   3630
         TabIndex        =   28
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "절수"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   4
         Left            =   3630
         TabIndex        =   29
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "수량"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   43
         Top             =   1230
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "사용구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   6
         Left            =   6090
         TabIndex        =   40
         Top             =   840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkReWorkClss 
            Caption         =   "재투입"
            Height          =   255
            Left            =   60
            TabIndex        =   44
            Top             =   30
            Width           =   885
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   7
         Left            =   6090
         TabIndex        =   41
         Top             =   1230
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkEmerClss 
            Caption         =   "긴급구분"
            Height          =   180
            Left            =   90
            TabIndex        =   45
            Top             =   60
            Width           =   1035
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   3
         Left            =   3240
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   90
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   8
         Left            =   6090
         TabIndex        =   54
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "사     종"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   9
         Left            =   6090
         TabIndex        =   55
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "제 직 처"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   10
         Left            =   3630
         TabIndex        =   56
         Top             =   870
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "생지폭"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   11
         Left            =   3630
         TabIndex        =   57
         Top             =   1230
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "생지밀도"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1614
      _Version        =   196609
      Begin VB.TextBox txtSearch 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   7830
         MaxLength       =   4
         TabIndex        =   52
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   6600
         MaxLength       =   8
         TabIndex        =   49
         Top             =   495
         Width           =   1185
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   5
         Top             =   75
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2820
         TabIndex        =   4
         Top             =   495
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   6600
         TabIndex        =   3
         Top             =   75
         Width           =   1905
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10980
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   2
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   780
      End
      Begin VB.ComboBox cboProcess 
         Height          =   300
         Left            =   8610
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   495
         Width           =   1335
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   795
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1402
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   480
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   10
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   4785
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   1440
         TabIndex        =   12
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
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
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품     명"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   4770
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   495
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   5220
         TabIndex        =   15
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
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
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   5220
         TabIndex        =   17
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
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
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "카드번호"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   8610
         TabIndex        =   50
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
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
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "대기공정"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   51
            Top             =   60
            Width           =   1185
         End
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8340
      TabIndex        =   48
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      발행(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmCardChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE   As String = "\Report\WorkCard.xls"
Private Const REPORTFILE1   As String = "\Report\TmpWorkCard.xls"

Private m_bloading As Boolean
Private m_iFlag As Integer

Private Sub cboUseClss_Click()
    With cboUseClss
        If cboUseClss = "보류" And cboUseClss.Tag = "대기" And m_iFlag = ID_UPDATE Then
            MsgBox "공정카드의 사용구분을 '보류'로 지정할 수 없습니다", vbInformation + vbOKOnly
            cboUseClss = cboUseClss.Tag
        End If
    End With
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index >= 1 And Index <= 3 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    ElseIf Index = 4 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(4).Enabled = True
            txtSearch(5).Enabled = True
            txtSearch(4).SetFocus
        Else
            txtSearch(4).Enabled = False
            txtSearch(5).Enabled = False
        End If
    Else
        If chkSearch(Index).Value = vbChecked Then
            cboProcess.Enabled = True
            cboProcess.SetFocus
        Else
            cboProcess.Enabled = False
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Dim sOrderID$
    
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    ElseIf Index = 3 Then
        sOrderID = txtOrderID.Tag
        Call ReturnCode(LG_ORDER, , False, txtOrderID)
        
        If Len(txtOrderID.Tag) > 0 Then
            txtOrderNO = txtOrderID.Text
            txtOrderID = txtOrderID.Tag
            txtOrderID.Tag = sOrderID
            
            If Not txtOrderID = sOrderID Then
                Call MakeColorCombo(txtOrderID)
            End If
        Else
            txtOrderID = sOrderID
            txtOrderID.Tag = sOrderID
        End If
    End If
End Sub

Private Sub cmdOperate_Click(Index As Integer)

    On Error GoTo ErrHandler
    
    Select Case Index
        '-------------------------------------------------------------------------------------'
        Case ID_UPDATE
            If grdData.Rows = grdData.FixedRows Then
                MsgBox LoadResString(111), vbInformation
                cmdSearch.SetFocus
                Exit Sub
            End If
            
            If grdData.TextMatrix(grdData.Row, 12) = "작업" Then
                MsgBox "작업중인 카드는 카드변경을 할 수 없습니다.", vbInformation + vbOKOnly
                Exit Sub
            End If
            
            m_iFlag = ID_UPDATE
            
            Call ChangeMode(Me, False)
            Call ModeChange(False)
            
            txtOrderID.SetFocus
        '-------------------------------------------------------------------------------------'
        Case ID_SAVE
            If SaveData() Then
                Call ChangeMode(Me, True)
                Call ModeChange(True)
                Call FillGridData
              
                m_iFlag = -1
            End If
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL
            m_iFlag = -1
            Call ChangeMode(Me, True)
            Call ModeChange(True)
            Call ShowData
    End Select

    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmCardChange.cmdOperate_Click", Err.Description)
End Sub

Private Sub cmdPrint_Click()
    Dim sCardID$, sSplitID$, sPatternID$
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    With grdData
        sCardID = MakeCardID(.TextMatrix(.Row, 6), OM_REDUCE)
        sSplitID = .TextMatrix(.Row, 7)
        sPatternID = .TextMatrix(.Row, 18)
    End With
    
    Call PrintWorkCard(CryReport, sCardID, sSplitID, sPatternID, PlusMDI.PrintPreview)
End Sub

Public Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    Call InitGrid
    Call MakeProcessCombo
    Call ModeChange(True)
    
    With cboUseClss
        .AddItem "대기"
        .AddItem "작업"
        .AddItem "완료"
        .AddItem "보류"
        
        .ListIndex = -1
    End With
    
    For i = 1 To 3
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(4).Enabled = False
    txtSearch(5).Enabled = False
    cboProcess.Enabled = False
    
    pnlProgress.Visible = False
        
End Sub

Private Sub grdData_RowColChange()
    If m_bloading Then Exit Sub
    
    Call ShowData
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
        If optOrder(0).Value Then
            .ColWidth(5) = 1350
            .ColWidth(4) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(5) = 0
            .ColWidth(4) = 1350
            chkSearch(3).Caption = "관리번호"
        End If
    End With
End Sub

Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    Dim sOrderID$
    If KeyAscii = vbKeyReturn Then
        sOrderID = txtOrderID.Tag
        Call ReturnCode(LG_ORDER, , False, txtOrderID)
        
        If Len(txtOrderID.Tag) > 0 Then
            txtOrderNO = txtOrderID.Text
            txtOrderID = txtOrderID.Tag
            txtOrderID.Tag = sOrderID
            
            If Not txtOrderID = sOrderID Then
                Call MakeColorCombo(txtOrderID)
            End If
        Else
            txtOrderID = sOrderID
            txtOrderID.Tag = sOrderID
        End If
    End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    ElseIf KeyAscii = vbKeyReturn And Index >= 3 Then
        Call NextFocus
    End If
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Redraw = flexRDNone
        .Cols = 23
        
        Call SetVSFlexGrid(grdData)
        .Rows = 1
        .RowHeightMin = 390
        
        .TextArray(0) = " ":
        .TextArray(1) = " ":            .ColWidth(1) = 250:     .ColHidden(1) = True
        .TextArray(2) = "거래처":       .ColWidth(2) = 1000:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":         .ColWidth(3) = 1800:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":     .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNo":      .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "카드번호":     .ColWidth(6) = 1000:               .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "분할" & vbCrLf & "번호":     .ColWidth(7) = 500:            .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":         .ColWidth(8) = 1300:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "절수":         .ColWidth(9) = 500:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "수량":         .ColWidth(10) = 600:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "완료공정":    .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "대기공정":    .ColWidth(12) = 900:           .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "카드상태":    .ColWidth(13) = 900:           .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "계획공정":    .ColWidth(14) = 7000:             .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "색상코드":    .ColHidden(15) = True   '.ColWidth(15) = 0
        .TextArray(16) = "재투입구분":  .ColHidden(16) = True '.ColWidth(16) = 0
        .TextArray(17) = "긴급구분":    .ColHidden(17) = True '.ColWidth(17) = 0
        .TextArray(18) = "공정패턴":    .ColHidden(18) = True
        .TextArray(19) = "사종":        .ColHidden(19) = True
        .TextArray(20) = "제직처":      .ColHidden(20) = True
        .TextArray(21) = "생지폭":      .ColHidden(21) = True
        .TextArray(22) = "생지밀도":    .ColHidden(22) = True
        
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub MakeProcessCombo()
    Dim oCard As PlusLib2.CCard
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon

    Set rs = oCard.GetProcess(1)
    Set oCard = Nothing

    With cboProcess
        .Clear

        Do Until rs.EOF
            .AddItem CStr(rs!Process)
            .ItemData(.NewIndex) = CLng(Left(rs!ProcessID, 2))
            
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing

        If .ListCount > 0 Then .ListIndex = 0
    End With

    m_bloading = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    m_bloading = False
    Call ErrorBox(Err.Number, "frmCardChange.MakeProcessCombo", Err.Description)
End Sub

Private Sub FillGridData()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    m_bloading = True
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetOrder(IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), txtSearch(4), txtSearch(5), _
                                 IIf(chkSearch(5) = vbChecked, 1, 0), Format(Left(cboProcess.ItemData(cboProcess.ListIndex), 2), "00"), 0)
    Set oCard = Nothing
        
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & _
                    rs!Color & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                    rs!UseClss & vbTab & CheckNull(rs!AfterProc) & vbTab & rs!OrderSeq & vbTab & rs!ReWorkClss & vbTab & _
                    rs!EmerClss & vbTab & rs!PatternID & vbTab & rs!ThreadName & vbTab & rs!StuffCustom & vbTab & _
                    rs!StuffWidth & vbTab & rs!StuffDensity
            
            If rs!UseClss = "보류" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbRed
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            ElseIf rs!UseClss = "작업" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbBlue
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            End If
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            Call ShowData
        Else
            .HighLight = flexHighlightNever
            
            Call ClearData
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bloading = False
    Exit Sub

ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    pnlProgress.Visible = False
    m_bloading = False
    Call ErrorBox(Err.Number, "frmCard.FillGridData", Err.Description)
End Sub

Private Sub MakeColorCombo(sOrderID As String)
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetOrderSub(sOrderID)
    Set oCard = Nothing
    
    With cboColor
        .Clear

        Do Until rs.EOF
            .AddItem rs!Color
            .ItemData(.NewIndex) = CLng(rs!OrderSeq)
            
            rs.MoveNext
        Loop

        If .ListCount > 0 Then .ListIndex = 0
    End With
    rs.Close

    Set rs = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmCardChange.MakeColorCombo", Err.Description)
End Sub

Private Sub ShowData()
    Call MakeColorCombo(MakeOrderID(grdData.TextMatrix(grdData.Row, 4), OM_REDUCE))
    Call MakeComboThreadName
    Call MakeComboStuffInCustom
    With grdData
        txtOrderID = MakeOrderID(.TextMatrix(.Row, 4), OM_REDUCE)
        txtOrderID.Tag = MakeOrderID(.TextMatrix(.Row, 4), OM_REDUCE)
        txtOrderNO = .TextMatrix(.Row, 5)
        cboColor.ListIndex = FindComboBox(cboColor, CLng(.TextMatrix(.Row, 15)))
        cboColor.Tag = CLng(.TextMatrix(.Row, 15))
        txtRoll = Format(.TextMatrix(.Row, 9), "#,##0")
        txtRoll.Tag = Format(.TextMatrix(.Row, 9), "#,##0")
        txtQty = Format(.TextMatrix(.Row, 10), "#,##0")
        txtQty.Tag = Format(.TextMatrix(.Row, 10), "#,##0")
        cboUseClss = .TextMatrix(.Row, 13)
        cboUseClss.Tag = .TextMatrix(.Row, 13)
        chkReWorkClss.Value = IIf(.TextMatrix(.Row, 16) = "*", vbChecked, vbUnchecked)
        chkEmerClss.Value = IIf(.TextMatrix(.Row, 17) = "*", vbChecked, vbUnchecked)
        cboThreadName.Text = .TextMatrix(.Row, 19)
        cboStuffCustom = .TextMatrix(.Row, 20)
        txtStuffWidth = .TextMatrix(.Row, 21)
        txtStuffDensity = .TextMatrix(.Row, 22)
    End With
End Sub

Private Sub ModeChange(bValue As Boolean)
    frmSearch.Enabled = bValue
    frmDetail.Enabled = Not bValue
    grdData.Enabled = bValue
    cmdFind(3).Enabled = Not bValue
End Sub

Private Function SaveData() As Boolean
    Dim tItem As PlusLib2.TCard
    Dim oCard As PlusLib2.CCard
    Dim i%
    
    On Error GoTo ErrHandler
    
    If CInt(txtQty) > CInt(txtQty.Tag) Then
        MsgBox "원래 수량보다 큰 값으로는 변경할 수 없습니다.", vbInformation + vbOKOnly
        txtQty = txtQty.Tag
        SaveData = False
        Exit Function
    End If

    
    If CInt(txtRoll) > CInt(txtRoll.Tag) Then
        MsgBox "원래 절수보다 큰 값으로는 변경할 수 없습니다.", vbInformation + vbOKOnly
        txtRoll = txtRoll.Tag
        SaveData = False
        Exit Function
    End If
    
    
    With tItem
        .sCardID = MakeCardID(grdData.TextMatrix(grdData.Row, 6), OM_REDUCE)
        .sSplitID = grdData.TextMatrix(grdData.Row, 7)
        .sOrderID = txtOrderID
        
        If txtOrderID = txtOrderID.Tag Then
            .nChkOrder = 0
        Else
            .nChkOrder = 1
        End If
        
        If cboColor.ListIndex = -1 Then
            .nOrderSeq = 0
        Else
            .nOrderSeq = cboColor.ItemData(cboColor.ListIndex)
            
            If cboColor.ItemData(cboColor.ListIndex) = cboColor.Tag Or .nChkOrder = 1 Then
                .nChkColor = 0
            Else
                .nChkColor = 1
            End If
        End If
        .nRoll = CheckNum(txtRoll)
        .nQty = CheckNum(txtQty)
        .sUseClss = cboUseClss
        .nChkUseClss = 0
        If cboUseClss <> cboUseClss.Tag And cboUseClss.Tag = "보류" Then
            .nChkUseClss = 1   '보류에서 대기로 변경될때 Hold Table 보류 취소 업데이트
        End If
        
        If cboUseClss = cboUseClss.Tag Then
            .nChkUseClss = 0
        Else
            .nChkUseClss = 1
        End If
        
        .sReWorkClss = IIf(chkReWorkClss.Value, "*", "")
        .sEmerClss = IIf(chkEmerClss.Value, "*", "")
        
        .sThreadName = cboThreadName
        .sStuffCustom = cboStuffCustom
        .nStuffWidth = CheckNum(txtStuffWidth)
        .nStuffDensity = CheckNum(txtStuffDensity)
        
        .sPersonID = g_sUserName
        .sModiClss = "카드변경"
    End With
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    oCard.UserName = g_sUserName
    
    If oCard.UpdateCardChange(tItem) Then
        SaveData = True
    Else
        SaveData = False
    End If
    Set oCard = Nothing
    
    Exit Function
ErrHandler:
    Set oCard = Nothing
    SaveData = False
    Call ErrorBox(Err.Number, "frmCardChange.SaveData", Err.Description)
End Function

Private Sub ClearData()
    txtOrderID = ""
    txtOrderNO = ""
    cboColor.ListIndex = -1
    cboColor.Tag = ""
    txtRoll = 0
    txtRoll.Tag = 0
    txtQty = 0
    txtQty.Tag = 0
    txtStuffWidth = 0
    txtStuffDensity = 0
    cboUseClss.ListIndex = -1
    cboUseClss.Tag = ""
    chkReWorkClss.Value = vbUnchecked
    chkEmerClss.Value = vbUnchecked
    cboThreadName.ListIndex = 0
    cboStuffCustom.ListIndex = 0
End Sub

Private Sub MakeComboThreadName()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetThreadName(MakeOrderID(grdData.TextMatrix(grdData.Row, 4), OM_REDUCE))
    Set oCard = Nothing
    
    With cboThreadName
        If rs.EOF Then
            .ListIndex = -1
            Exit Sub
        End If
        
        .Clear
        
        Do Until rs.EOF
            .AddItem rs(0)

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        .ListIndex = -1
    End With
    
    Exit Sub
ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmCardChange.MakeComboThreadName", Err.Description)
End Sub

Private Sub MakeComboStuffInCustom()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    Set rs = oCard.GetStuffCustom(MakeOrderID(grdData.TextMatrix(grdData.Row, 4), OM_REDUCE))
    Set oCard = Nothing
    
    With cboStuffCustom
        If rs.EOF Then
            .ListIndex = -1
            Exit Sub
        End If
        
        .Clear
        
        Do Until rs.EOF
            .AddItem rs(0)

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        .ListIndex = -1
    End With
       
    Exit Sub
ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmCardChange.MakeComboStuffInCustom", Err.Description)
End Sub

