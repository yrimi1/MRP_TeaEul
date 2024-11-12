VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetTerminal 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   6855
   ClientLeft      =   1935
   ClientTop       =   1425
   ClientWidth     =   11850
   Icon            =   "frmSetTerminal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   8280
      TabIndex        =   13
      Top             =   6120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      저장(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10080
      TabIndex        =   12
      Top             =   6120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab stbSetTerminal 
      Height          =   6465
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   11404
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   794
      TabCaption(0)   =   "  검사 시스템 운영 조건 설정  "
      TabPicture(0)   =   "frmSetTerminal.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeEdit(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmeEdit(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fmeEdit(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "  불량 등록 및 불량 배치 설정  "
      TabPicture(1)   =   "frmSetTerminal.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCopy"
      Tab(1).Control(1)=   "cmdUp"
      Tab(1).Control(2)=   "grdDefectList"
      Tab(1).Control(3)=   "grdButton"
      Tab(1).Control(4)=   "cmdDown"
      Tab(1).Control(5)=   "cmdDelete"
      Tab(1).ControlCount=   6
      Begin Threed.SSCommand cmdCopy 
         Height          =   510
         Left            =   -74940
         TabIndex        =   18
         Top             =   5370
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "불량 복사"
      End
      Begin Threed.SSCommand cmdUp 
         Height          =   510
         Left            =   -72735
         TabIndex        =   14
         Top             =   5370
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "위"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDefectList 
         Height          =   5250
         Left            =   -74970
         TabIndex        =   15
         Top             =   60
         Width           =   3915
         _cx             =   6906
         _cy             =   9260
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
         Rows            =   25
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
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
         OleDropMode     =   1
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdButton 
         Height          =   5865
         Left            =   -70980
         TabIndex        =   16
         Top             =   60
         Width           =   7650
         _cx             =   13494
         _cy             =   10345
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
         Rows            =   6
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
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
         AutoResize      =   0   'False
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
      Begin Threed.SSCommand cmdDown 
         Height          =   510
         Left            =   -71865
         TabIndex        =   17
         Top             =   5370
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "아래"
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   510
         Left            =   -73950
         TabIndex        =   19
         Top             =   5370
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "불량 삭제"
      End
      Begin Threed.SSFrame fmeEdit 
         Height          =   2220
         Index           =   6
         Left            =   210
         TabIndex        =   20
         Top             =   3570
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   3916
         _Version        =   196609
         Caption         =   "  3. 불량 버튼 설정  "
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   22
            Left            =   150
            TabIndex        =   21
            Top             =   255
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "불량 버튼 개수 X"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtButtonX 
            Height          =   315
            Left            =   2370
            TabIndex        =   8
            Top             =   255
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   23
            Left            =   150
            TabIndex        =   22
            Top             =   645
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "불량 버튼 개수 Y"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtButtonY 
            Height          =   315
            Left            =   2370
            TabIndex        =   9
            Top             =   645
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   24
            Left            =   150
            TabIndex        =   23
            Top             =   1035
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "사용 색상수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtColorCnt 
            Height          =   315
            Left            =   2370
            TabIndex        =   10
            Top             =   1035
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   25
            Left            =   150
            TabIndex        =   24
            Top             =   1425
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "반복 색상 행수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtRepeatCnt 
            Height          =   315
            Left            =   2370
            TabIndex        =   11
            Top             =   1425
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   34
            Top             =   1800
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "글자 크기"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtFontSize 
            Height          =   315
            Left            =   2370
            TabIndex        =   35
            Top             =   1800
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame fmeEdit 
         Height          =   2265
         Index           =   5
         Left            =   210
         TabIndex        =   25
         Top             =   1230
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   3995
         _Version        =   196609
         Caption         =   "  2. 세부검사 환경 설정  "
         Begin VB.ComboBox cboGradeClss 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":0044
            Left            =   2370
            List            =   "frmSetTerminal.frx":0046
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   300
            Width           =   2880
         End
         Begin VB.ComboBox cboDemeritClss 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":0048
            Left            =   2370
            List            =   "frmSetTerminal.frx":004A
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   690
            Width           =   2880
         End
         Begin VB.ComboBox cboLossClss 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":004C
            Left            =   2370
            List            =   "frmSetTerminal.frx":004E
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   1080
            Width           =   2880
         End
         Begin VB.ComboBox cboDefectClss 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":0050
            Left            =   2370
            List            =   "frmSetTerminal.frx":0052
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   1470
            Width           =   2880
         End
         Begin VB.ComboBox cboCutDefect 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":0054
            Left            =   2370
            List            =   "frmSetTerminal.frx":0056
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   1860
            Width           =   2880
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   14
            Left            =   150
            TabIndex        =   26
            Top             =   300
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "등급 결정 방법"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   15
            Left            =   150
            TabIndex        =   27
            Top             =   690
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "벌점 적용 방법"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   17
            Left            =   150
            TabIndex        =   28
            Top             =   1080
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "보상 적용 방법"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   20
            Left            =   150
            TabIndex        =   29
            Top             =   1470
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "대표 불량 등급"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   21
            Left            =   150
            TabIndex        =   30
            Top             =   1860
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "난단 대표 불량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSFrame fmeEdit 
         Height          =   1065
         Index           =   4
         Left            =   195
         TabIndex        =   31
         Top             =   75
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   1879
         _Version        =   196609
         Caption         =   "  1. 기본 환경 설정  "
         Begin VB.ComboBox cboRoundClss 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":0058
            Left            =   2370
            List            =   "frmSetTerminal.frx":005A
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   270
            Width           =   2880
         End
         Begin VB.ComboBox cboRollClss 
            Height          =   300
            ItemData        =   "frmSetTerminal.frx":005C
            Left            =   2370
            List            =   "frmSetTerminal.frx":005E
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   660
            Width           =   2880
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   13
            Left            =   135
            TabIndex        =   32
            Top             =   270
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "소숫점 관리"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   33
            Top             =   660
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "Roll No 관리"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmSetTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH As Integer = 11000
Private Const LIMIT_ROW = 16

Private m_bLoadingCombo  As Boolean
Private m_sFlag          As String * 1
Private m_nButtonX       As Integer
Private m_nButtonY       As Integer
Private m_nButtonSeqMax  As Integer
Private m_nColorCnt      As Integer
Private m_nRepeatCnt     As Integer
Private m_nFontSize      As Integer
Private m_nColor(1 To 7)

Private Sub Form_Load()
    Me.Move 0, 0, 11975, 7260
    Me.Caption = stbSetTerminal.Caption
     
'    Color(0) = &HFFFFFF
    m_nColor(1) = &HC0C0FF
    m_nColor(2) = &HC0E0FF
    m_nColor(3) = &HC0FFFF
    m_nColor(4) = &HC0FFC0
    m_nColor(5) = &HFFFFC0
    m_nColor(6) = &HFFC0C0
    m_nColor(7) = &HFFC0FF

    Call InitGrid
    Call SetCombo
    Call FillSetTerminal

    cmdSave.Picture = LoadResPicture("CHECK", vbResIcon)
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
End Sub

Private Sub stbSetTerminal_Click(PreviousTab As Integer)
    If stbSetTerminal.Tab = 1 Then
        Call MakeGridButton
        Call FillDefectList
        Call ButtonReArray
    End If
End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdDefectList)
   
    With grdDefectList
        .Redraw = flexRDNone

        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightAlways
        .ExplorerBar = flexExNone

        .Rows = 1
        .Cols = 9

        .TextArray(0) = "":             .ColWidth(0) = 0
        .TextArray(1) = "순서":         .ColWidth(1) = 700:     .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "불량명":       .ColWidth(2) = 1700:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "감점":         .ColWidth(3) = 700:     .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "보상":         .ColWidth(4) = 700:     .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "Display1":     .ColWidth(5) = 0
        .TextArray(6) = "Display2":     .ColWidth(6) = 0
        .TextArray(7) = "Display3":     .ColWidth(7) = 0
        .TextArray(8) = "불량종류":     .ColWidth(8) = 0

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetCombo()
    m_bLoadingCombo = True

    With cboRoundClss ' 소숫점 관리
        .AddItem "1. 절사"
        .AddItem "2. 반올림"
        .AddItem "3. 5사6입"
        .AddItem "4. 6사7입"
        .AddItem "5. 7사8입"
        .AddItem "6. 8사9입"
    End With
    
    With cboRollClss ' Roll NO 관리
        .AddItem "1. Order 별"
        .AddItem "2. Color 별"
        .AddItem "3. Order, Color, Lot 별"
        .AddItem "4. Order, Lot 별"
        .AddItem "5. Order, Color, 호기별"
    End With
    
    With cboDemeritClss ' 벌점 적용 방법
        .AddItem "1. 사용안함"
        .AddItem "2. 불량보상"
        .AddItem "3. 검사기준"
        .AddItem "4. 수동입력적용"
    End With
    
    With cboGradeClss ' 등급 결정 방법
        .AddItem "1. 사용안함"
        .AddItem "2. 검사자지정"
        .AddItem "3. 업체요구"
    End With

    With cboLossClss ' 보상 적용 방법
        .AddItem "1. 사용안함"
        .AddItem "2. 불량보상"
    End With

    With cboCutDefect ' 난단 대표 불량
        .AddItem "1.사용안함"
        .AddItem "2.사용함"
    End With

    Call MakeCodeCombo(cboDefectClss, CD_GRADE)

    m_bLoadingCombo = False
End Sub

Private Sub FillSetTerminal()
    Dim oSetTerminal As PlusLib2.CSetTerminal
    Dim rs As ADODB.Recordset

    On Error GoTo ErrHandler
    Set oSetTerminal = New PlusLib2.CSetTerminal
    oSetTerminal.Connection = g_adoCon
    oSetTerminal.UserName = g_sUserName
    
    Set rs = oSetTerminal.GetSetTerminal
    Set oSetTerminal = Nothing

    If Not rs.EOF Then
        cboRoundClss.ListIndex = rs!RoundClss
        cboGradeClss.ListIndex = rs!GradeClss
        cboDemeritClss.ListIndex = rs!DemeritClss
        cboLossClss.ListIndex = rs!LossClss
        cboDefectClss.ListIndex = rs!DefectClss
        cboCutDefect.ListIndex = rs!CutDefect
        txtButtonX = rs!ButtonX
        txtButtonY = rs!ButtonY
        txtColorCnt = rs!ColorCnt
        txtRepeatCnt = rs!RepeatCnt
        txtFontSize = rs!FontSize
        cboRollClss.ListIndex = rs!RollClss

        m_nButtonX = rs!ButtonX
        m_nButtonY = rs!ButtonY
        m_nButtonSeqMax = m_nButtonX * m_nButtonY
        m_nColorCnt = rs!ColorCnt
        m_nRepeatCnt = rs!RepeatCnt
        m_nFontSize = rs!FontSize
    End If

    rs.Close
    Set rs = Nothing

    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oSetTerminal = Nothing
    
    Call ErrorBox(Err.Number, "SetTerminal.FillSetTerminal", Err.Description)
End Sub

Private Sub MakeGridButton()
    Dim ColMin%, RowMin%, ColSum%, RowSum%
    Dim i%, iColor%, j%

    Call SetVSFlexGrid(grdButton)

    ColMin = CInt(grdButton.Width / m_nButtonX)
    RowMin = CInt(grdButton.Height / m_nButtonY)

    ColSum = 0
    RowSum = 0

    With grdButton
        .Redraw = flexRDNone

        .Rows = .FixedRows

        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarNone

        .Rows = m_nButtonY
        .Cols = m_nButtonX
        
        .FontSize = m_nFontSize
        For i = 0 To m_nButtonX - 2
            .ColWidth(i) = ColMin
            ColSum = ColSum + ColMin
        Next i
        .ColWidth(i) = grdButton.Width - ColSum

        For i = 0 To m_nButtonY - 2
            .RowHeight(i) = RowMin
            RowSum = RowSum + RowMin
        Next i
        .RowHeight(i) = grdButton.Height - RowSum

        .Redraw = flexRDDirect

        iColor = 1
        For i = 0 To m_nButtonY - 1
            If i <> 0 And (i Mod m_nRepeatCnt = 0) Then
                iColor = iColor + 1
            End If
            If iColor > m_nColorCnt Then
                iColor = 1
            End If

            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = m_nColor(iColor)
       Next i
    End With
End Sub

Private Sub FillDefectList()
    Dim oDefect As PlusLib2.CSetTerminal
    Dim rs As ADODB.Recordset
    Dim iNowRow&, iRow&, iCol&, RowCount%

    On Error GoTo ErrHandler

    Set oDefect = New PlusLib2.CSetTerminal
    oDefect.Connection = g_adoCon

    Set rs = oDefect.GetTerminalDefect()
    Set oDefect = Nothing

    With grdDefectList
        .Redraw = False

        iNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        Do Until rs.EOF
            .AddItem rs!DefectID & vbTab & rs!ButtonSeq & vbTab & rs!KDefect & vbTab & rs!Demerit & vbTab & _
                rs!Loss & vbTab & rs!Display1 & vbTab & rs!Display2 & vbTab & rs!Display3 & vbTab & rs!DefectClss

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = iNowRow
            
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If

        .Redraw = flexRDDirect
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oDefect = Nothing

    Call ErrorBox(Err.Number, "Defect.FillButton", Err.Description)
End Sub

Private Sub ButtonReArray()
    Dim i&, iRow&, iCol&, Position&

'    With grdDefectList
'        .Redraw = flexRDNone
'
'        For i = .FixedRows To .Rows - 1
'            Position = .TextMatrix(i, 1) - 1
'
'            iRow = Int(Position / m_nButtonX)
'            iCol = Position Mod m_nButtonX
'
'            grdButton.TextMatrix(iRow, iCol) = .TextMatrix(i, 5) & vbCrLf & .TextMatrix(i, 6) & vbCrLf & .TextMatrix(i, 7)
'        Next i
'
'        .Redraw = flexRDDirect
'    End With
    
    
    With grdDefectList
        .Redraw = flexRDNone

        For i = .FixedRows To .Rows - 1
            Position = .TextMatrix(i, 1) - 1

            iRow = Int(Position / m_nButtonX)
            iCol = Position Mod m_nButtonX

            grdButton.TextMatrix(iRow, iCol) = .TextMatrix(i, 5) & vbCrLf & .TextMatrix(i, 6) & vbCrLf & .TextMatrix(i, 7)
'            grdButton.TextMatrix(iRow, iCol) = .TextMatrix(i, 4)
            
            With grdButton
'                .Row = iRow
'                .Col = iCol
'                .ColSel = iCol
'                .CellBackColor = m_nColor(grdDefectList.TextMatrix(i, 8))
                .Cell(flexcpBackColor, iRow, iCol, iRow, iCol) = m_nColor(grdDefectList.TextMatrix(i, 8))
            End With
        Next i

        .Redraw = flexRDDirect
    End With
    

    Call SetFocusButton
End Sub

Private Sub SetFocusButton()
    Dim iRow&, iCol&, Position&

    With grdDefectList
        Position = .TextMatrix(.Row, 1) - 1
        iRow = Int(Position / m_nButtonX)
        iCol = Position Mod m_nButtonX
    End With

    With grdButton
        If grdDefectList.TextMatrix(grdDefectList.Row, 1) < m_nButtonSeqMax + 1 Then
            .Row = iRow
            .Col = iCol
        End If
    End With
End Sub

Private Sub cmdCopy_Click()
    Dim i&, MaxSeq%, sCopyRow$
    Dim Display1 As String
    Dim Display2 As String
    Dim Display3 As String

    On Error GoTo ErrHandler
    
    With grdDefectList
        For i = .FixedRows To .Rows - .FixedRows
            If MaxSeq < .TextMatrix(i, 1) Then
                MaxSeq = .TextMatrix(i, 1)
            End If
        Next i
        MaxSeq = MaxSeq + 1   '새로 추가될 버튼의 Seq
        
        If MaxSeq > m_nButtonSeqMax Then Exit Sub
        
        .Redraw = flexRDNone
        
        sCopyRow = .Cell(flexcpText, .Row, 0, .Row, 4)
        Display1 = .TextMatrix(.Row, 5)
        Display2 = .TextMatrix(.Row, 6)
        Display3 = .TextMatrix(.Row, 7)

        .AddItem "", .Row + 1

        .Cell(flexcpText, .Row + 1, 0, .Row + 1, 4) = sCopyRow
        .TextMatrix(.Row + 1, 5) = Display1
        .TextMatrix(.Row + 1, 6) = Display2
        .TextMatrix(.Row + 1, 7) = Display3
        
        .TextMatrix(.Row + 1, 1) = MaxSeq

        .Col = 1
        .Sort = flexSortNumericAscending
        .ColSel = .Cols - 1
        
        .Redraw = flexRDDirect
    End With

    Call ButtonReArray
    Exit Sub

ErrHandler:
    grdDefectList.Redraw = flexRDDirect

    Call ErrorBox(Err.Number, "SetTerminal.CopyDefect", Err.Description)
End Sub

Private Sub cmdDelete_Click()
    Dim oDefect As PlusLib2.CDefect
    Dim NewDefectSub As PlusLib2.tDefectSub
    Dim i%, nCurRow%, nCurCol%
    
    On Error GoTo ErrHandler
    
    With grdDefectList
        nCurRow = .Row
        
        .Redraw = flexRDNone
        .TextMatrix(.Row, 1) = 60
        
        For i = .Row + 1 To .Rows - 1
            .TextMatrix(i, 1) = .TextMatrix(i, 1) - 1
        Next i
        
        .Row = 1
        .Col = 1
        .Sort = flexSortNumericAscending
        .Rows = .Rows - 1
        
        .Redraw = flexRDDirect
        
        Call MakeGridButton
        Call ButtonReArray
        
        nCurRow = nCurRow - 1
        If nCurRow < 0 Then
            nCurRow = 1
        End If
        .Row = nCurRow
    End With
    
    Exit Sub
    
ErrHandler:
    Set oDefect = Nothing
    
    Call ErrorBox(Err.Number, "SetTerminal.DeleteDefect", Err.Description)
End Sub

Private Sub cmdUP_Click()
    Dim nCurrSeq&, nNewSeq&
    Dim sCopyRow(8) As String, sCopyRowNew(8) As String
    Dim i&
    Dim Flag As Boolean
    Dim nRow%, nCol%
    
    With grdDefectList
        If (.Row = 1 And .TextMatrix(.Row, 1) = 1) Or .TextMatrix(.Row, 1) = 1 Then
            Exit Sub
        End If
        
        If grdButton.TextMatrix(grdButton.Row, grdButton.Col) = "" Then
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        nRow = grdButton.Row
        nCol = grdButton.Col - 1
        
        If nCol = -1 Then
            nCol = grdButton.Cols - 1
            nRow = nRow - 1
            
            If nRow < 0 Then
                Exit Sub
            End If
        End If
        
        If grdButton.TextMatrix(nRow, nCol) = "" Then
            grdButton.TextMatrix(grdButton.Row, grdButton.Col) = ""
            .TextMatrix(.Row, 1) = .TextMatrix(.Row, 1) - 1
        Else
            nCurrSeq = .TextMatrix(.Row, 1)
            nNewSeq = .TextMatrix(.Row, 1) - 1
            
            For i = 0 To .Cols - 1
                sCopyRowNew(i) = .TextMatrix(.Row, i)
                
                sCopyRow(i) = .TextMatrix(.Row - 1, i)
            Next i
            
            For i = 0 To .Cols - 1
                .TextMatrix(.Row, i) = sCopyRow(i)
                
                .TextMatrix(.Row - 1, i) = sCopyRowNew(i)
            Next i
            
            .TextMatrix(.Row - 1, 1) = nNewSeq
            .TextMatrix(.Row, 1) = nCurrSeq
            .Row = .Row - 1
        End If
        
        .Redraw = flexRDDirect
    End With
    
    Call ButtonReArray
End Sub

Private Sub cmdDown_Click()
    Dim nCurrSeq&, nNewSeq&
    Dim sCopyRow(8) As String, sCopyRowNew(8) As String
    Dim i&
    Dim Flag As Boolean
    Dim nRow%, nCol%
    
    With grdDefectList
        If (.Row = .Rows - 1 And .TextMatrix(.Row, 1) = m_nButtonSeqMax) Or .TextMatrix(.Row, 1) = m_nButtonSeqMax Then
            Exit Sub
        End If
        
        If grdButton.TextMatrix(grdButton.Row, grdButton.Col) = "" Then
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        nRow = grdButton.Row
        nCol = grdButton.Col + 1
        
        If nCol > grdButton.Cols - 1 Then
            nCol = 0
            nRow = nRow + 1
            
            If nRow > grdButton.Rows - 1 Then
                Exit Sub
            End If
        End If
        
        If grdButton.TextMatrix(nRow, nCol) = "" Then
            grdButton.TextMatrix(grdButton.Row, grdButton.Col) = ""
            
            .TextMatrix(.Row, 1) = .TextMatrix(.Row, 1) + 1
        Else
        
          nCurrSeq = .TextMatrix(.Row, 1)
          nNewSeq = .TextMatrix(.Row, 1) + 1
          
          For i = 0 To .Cols - 1
              sCopyRowNew(i) = .TextMatrix(.Row, i)
              
              sCopyRow(i) = .TextMatrix(.Row + 1, i)
          Next i
          
          For i = 0 To .Cols - 1
              .TextMatrix(.Row, i) = sCopyRow(i)
              
              .TextMatrix(.Row + 1, i) = sCopyRowNew(i)
          Next i
          
          .TextMatrix(.Row, 1) = nCurrSeq
          .TextMatrix(.Row + 1, 1) = nNewSeq
          
           .Row = .Row + 1
          
        End If

        .Redraw = flexRDDirect
    End With
    
    Call ButtonReArray
End Sub

Private Sub grdDefectList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim nCol%
    Dim vValue As Variant
    Dim preRow%, preCol%

    With grdDefectList
        preRow = .Row
        preCol = .Col
        vValue = .TextMatrix(.Row, .Col)
        nCol = .Col + 1
        If nCol > 4 Then
            .Col = 3
            If .Row + 1 > .Rows - 1 Then
                Exit Sub
            Else
            .Row = .Row + 1
            End If

        Else
            .Col = nCol
        End If
    End With
End Sub

Private Sub grdDefectList_RowColChange()
    Call SetFocusButton

    With grdDefectList
        If .Col = 3 Or .Col = 4 Then
            .SelectionMode = flexSelectionFree
        Else
            .SelectionMode = flexSelectionByRow
        End If
        
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
    End With
End Sub

Private Sub grdButton_RowColChange()
    Dim i&, iRow&, iCol&, Position&
    
    With grdButton
        iRow = .Row
        iCol = .Col
        Position = (iRow * m_nButtonX) + iCol + 1
    End With

    With grdDefectList
        For i = .FixedRows To .Rows - .FixedRows
            If Position = IIf(.TextMatrix(i, 1) = "", 0, .TextMatrix(i, 1)) Then
                .Row = i
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub cmdSave_Click()
    Dim NewSetTerminal As PlusLib2.TSetTerminal
    Dim oSetTerminal As PlusLib2.CSetTerminal
    Dim NewDefectSub() As PlusLib2.tDefectSub

    Dim oDefectSub As PlusLib2.CDefect
    Dim i%, nDefectRow%

    On Error GoTo ErrHandler

    If stbSetTerminal.Tab = 0 Then
        Set oSetTerminal = New PlusLib2.CSetTerminal
        oSetTerminal.Connection = g_adoCon
        oSetTerminal.UserName = g_sUserName

        With NewSetTerminal
            .sRoundClss = cboRoundClss.ListIndex
            .sGradeClss = cboGradeClss.ListIndex
            .sDemeritClss = cboDemeritClss.ListIndex
            .sLossClss = cboLossClss.ListIndex
            .sDefectClss = cboDefectClss.ListIndex
            .sCutDefect = cboCutDefect.ListIndex
            .nButtonX = txtButtonX
            .nButtonY = txtButtonY
            .nColorCnt = txtColorCnt
            .nRepeatCnt = txtRepeatCnt
            .nFontSize = txtFontSize
            .nRollClss = cboRollClss.ListIndex
        End With
                
        If oSetTerminal.AddNewSetTerminal(NewSetTerminal) Then
            m_nButtonX = txtButtonX
            m_nButtonY = txtButtonY
            m_nButtonSeqMax = m_nButtonX * m_nButtonY
            m_nColorCnt = txtColorCnt
            m_nRepeatCnt = txtRepeatCnt
            m_nFontSize = txtFontSize
            Set oSetTerminal = Nothing
            Exit Sub
        End If
    Else
        Set oDefectSub = New PlusLib2.CDefect
        oDefectSub.Connection = g_adoCon
        oDefectSub.UserName = g_sUserName
    
        With grdDefectList
            nDefectRow = .Rows - .FixedRows - 1
            
            ReDim NewDefectSub(nDefectRow)
            
            For i = 0 To nDefectRow
                NewDefectSub(i).DefectID = .TextMatrix(.FixedRows + i, 0)       '[1] 불량 코드
                NewDefectSub(i).ButtonSeq = .TextMatrix(.FixedRows + i, 1)       '[2] 버튼위치
                NewDefectSub(i).Demerit = .TextMatrix(.FixedRows + i, 3)
                NewDefectSub(i).Loss = .TextMatrix(.FixedRows + i, 4)
            Next i
        End With
        
        If oDefectSub.AddNewDefectSub(NewDefectSub) Then
            Set oDefectSub = Nothing
            Exit Sub
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    Set oSetTerminal = Nothing
    
    Call ErrorBox(Err.Number, "SetTerminal.SaveData", Err.Description)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

