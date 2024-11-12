VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfoSet 
   ClientHeight    =   9255
   ClientLeft      =   3330
   ClientTop       =   2940
   ClientWidth     =   11865
   Icon            =   "frmInfoSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   690
      Left            =   2190
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   75
      Width           =   840
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   330
      Index           =   1
      Left            =   60
      MousePointer    =   99  '사용자 정의
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   435
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   330
      Index           =   0
      Left            =   60
      MousePointer    =   99  '사용자 정의
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   615
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   375
      Left            =   7335
      TabIndex        =   1
      Top             =   75
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "☜  날짜를 선택하십시오"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid grdInfoUser 
      Height          =   3555
      Left            =   15
      TabIndex        =   2
      Top             =   4920
      Width           =   3015
      _cx             =   5318
      _cy             =   6271
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
   Begin VSFlex7LCtl.VSFlexGrid grdInfo 
      Height          =   4065
      Left            =   15
      TabIndex        =   3
      Top             =   810
      Width           =   3015
      _cx             =   5318
      _cy             =   7170
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   2
      Left            =   4680
      TabIndex        =   4
      Top             =   90
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23724032
      CurrentDate     =   37096
   End
   Begin Threed.SSPanel pnlName 
      Height          =   300
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
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
      Caption         =   "공지 일자"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10185
      TabIndex        =   8
      Top             =   8535
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
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
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlBorder 
      Height          =   4170
      Index           =   0
      Left            =   3120
      TabIndex        =   9
      Top             =   4320
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   7355
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
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdMove 
         Caption         =   "<<"
         Height          =   615
         Index           =   1
         Left            =   4050
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2355
         Width           =   615
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">>"
         Height          =   615
         Index           =   0
         Left            =   4050
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1590
         Width           =   615
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPerson 
         Height          =   3495
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   585
         Width           =   3915
         _cx             =   6906
         _cy             =   6165
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
      Begin Threed.SSPanel pnlName 
         Height          =   390
         Index           =   1
         Left            =   165
         TabIndex        =   13
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         _Version        =   196609
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사용자 설정 (공지사항)"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPerson 
         Height          =   3495
         Index           =   1
         Left            =   4740
         TabIndex        =   14
         Top             =   585
         Width           =   3915
         _cx             =   6906
         _cy             =   6165
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
   End
   Begin Threed.SSPanel pnlBorder 
      Height          =   3750
      Index           =   2
      Left            =   3105
      TabIndex        =   15
      Top             =   510
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   6615
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
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtInfo 
         Height          =   3105
         Index           =   0
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   17
         Top             =   540
         Width           =   4830
      End
      Begin VB.TextBox txtInfo 
         Height          =   3105
         Index           =   1
         Left            =   4935
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   16
         Top             =   540
         Width           =   3720
      End
      Begin Threed.SSPanel pnlName 
         Height          =   390
         Index           =   2
         Left            =   60
         TabIndex        =   18
         Top             =   135
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
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
         Caption         =   "알림 사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   390
         Index           =   0
         Left            =   4935
         TabIndex        =   19
         Top             =   135
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
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
         Caption         =   "사용자별 공지사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   8460
      TabIndex        =   20
      Top             =   8535
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
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
      Caption         =   "      확인(&O)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   735
      TabIndex        =   21
      Top             =   75
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23724033
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   735
      TabIndex        =   22
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23724033
      CurrentDate     =   36871
   End
   Begin Threed.SSCommand cmdNew 
      Height          =   420
      Left            =   10170
      TabIndex        =   23
      Top             =   45
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   741
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
      Caption         =   "새 공지사항 (&N)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblCount 
      Caption         =   "검색건수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   8805
      Width           =   2520
   End
End
Attribute VB_Name = "frmInfoSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''Private Const LIMIT_WIDTH2 = 2350
''Private Const LIMIT_WIDTH3 = 3750
''Private Const LIMIT_WIDTH1 = 2040
''
''Private Const LIMIT_ROW1 = 12
''Private Const LIMIT_ROW2 = 28

Private m_bFlag As Boolean

Private Sub cmdSearch_Click()
    Call FillGridInfo
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11985, 9660

    m_bFlag = False

    Call SetOperate(Me)
    Call InitGrid
    Call FillGridPerson
    dtpDate(2) = Now
    
    Me.Show
    
    Call cmdTerm_Click(1)   ' 금월로 설정

    pnlMsg.Visible = False
    cmdSave.Picture = LoadResPicture("CHECK", vbResIcon)
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

  '  Call FillGridInfo
End Sub

Private Sub InitGrid()
    With grdInfo
        .Cols = 3
        Call SetVSFlexGrid(grdInfo)

        .Rows = .FixedRows

        .TextArray(0) = ""
        .TextArray(1) = "공지일자":     .ColWidth(1) = 2350:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "알림사항":     .ColWidth(2) = 0
    End With

    With grdInfoUser
        .Cols = 5
        Call SetVSFlexGrid(grdInfoUser)

        .Rows = .FixedRows

        .TextArray(1) = "공지일자":                 .ColWidth(1) = 0
        .TextArray(2) = "일련번호":                 .ColWidth(2) = 0
        .TextArray(3) = "사용자별 공지사항 목록":   .ColWidth(3) = 0
        .TextArray(4) = "사용자별 공지사항 목록":   .ColWidth(4) = 2350:    .ColAlignment(4) = flexAlignLeftCenter
    End With

    With grdPerson(0)
        .Cols = 5
        Call SetVSFlexGrid(grdPerson(0))

        .Redraw = flexRDNone

        .FixedCols = 0
        .FixedRows = 1
        .Rows = .FixedRows

        .GridLines = flexGridNone
        .BackColorBkg = vbWhite
        .SheetBorder = vbWhite
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1

        .TextArray(0) = "":         .ColWidth(0) = 255
        .TextArray(1) = "부서명":   .ColWidth(1) = 1500:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "사원명":   .ColWidth(2) = 3750:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "부서ID":   .ColWidth(3) = 0
        .TextArray(4) = "사원ID":   .ColWidth(4) = 0

        .Redraw = flexRDDirect
    End With

    With grdPerson(1)
        .Cols = 3
        Call SetVSFlexGrid(grdPerson(1))

        .Rows = .FixedRows

        .TextArray(1) = "사원명":       .ColWidth(1) = 2040:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "사원ID":       .ColWidth(2) = 0
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        
        Call CheckedPerson
    Else
        With grdPerson(1)
            If .Rows = .FixedRows Or .Row = 0 Then Exit Sub ' row가 없으면 종료
                
            .RemoveItem .Row ' 해당 row를 삭제
                  
        End With
    End If
End Sub

Private Function SaveData() As Boolean
    Dim oInfo As PlusLib2.CInfo
    Dim NewInfo As PlusLib2.TInfo
    Dim NewInfoUser As PlusLib2.TInfoUser  '사용자별 공지사항
    Dim PersonID() As String
    Dim iLoop%, nSeq%
    Dim InfoSeq  ''개인별 공지사항 일련번호 설정..
 
    On Error GoTo ErrHandler
    
    If (Len(txtInfo(0)) = 0 And Len(txtInfo(1)) > 0) Then
        txtInfo(0).Text = "설정된 공지사항이 없습니다."
    End If
    
    Set oInfo = New PlusLib2.CInfo
    With NewInfo  ' 전체 공지
        .sInfoDate = MakeDate(DF_SHORT, dtpDate(2))
        .sInfo = txtInfo(0)
    End With
    
    nSeq = CheckCount() - 1 ' 개인 공지사항에 선택된 사람 수
 
    If nSeq > -1 Then
        ReDim PersonID(nSeq) '선택된 사용자 ID 저장용 배열..
        
        For iLoop = 0 To nSeq '선택된 사용자 ID를 저장
            PersonID(iLoop) = grdPerson(1).TextMatrix(iLoop + 1, 2)
        Next iLoop
    End If
    
    If CheckDate() Then  '같은 날짜 있음.
        If m_bFlag Then  '새공지 추가..
            oInfo.Connection = g_adoCon
            InfoSeq = oInfo.GetNewInfoSeq("[InfoUser]", "[InfoSeq]", "InfoDate = " & NewInfo.sInfoDate)
                ' 현재 입력된 개인공지 목록들의 일련번호중 가장 큰 번호
        Else  '기존 공지 업데이트
            InfoSeq = IIf(grdInfoUser.Row = 0, 1, grdInfoUser.TextMatrix(grdInfoUser.Row, 2))
        End If
    Else  ' 같은 날짜 없고 새로 입력시..
        InfoSeq = 1
        m_bFlag = True
    End If
     
    With NewInfoUser '개인별 공지 저장 구조체
        .sInfoDate = MakeDate(DF_SHORT, dtpDate(2))
        .nInfoseq = InfoSeq
        .sInfoUser = txtInfo(1).Text
    End With
    
    oInfo.Connection = g_adoCon
    oInfo.UserName = g_sUserName
    
    SaveData = oInfo.AddInfo(NewInfo, NewInfoUser, PersonID(), nSeq)
    
    m_bFlag = False
    
    Exit Function
ErrHandler:
    Call ErrorBox(Err.Number, "InfoSet.SaveData", Err.Description)
    
End Function

Private Sub cmdNew_Click()
    Dim iLoop As Integer
    
    m_bFlag = Not m_bFlag

    pnlMsg.Visible = m_bFlag
    grdInfo.Enabled = Not m_bFlag
    grdInfoUser.Enabled = Not m_bFlag

    If m_bFlag Then
        cmdNew.Caption = "입력 취소(&N)"
        
        Call ClearText(txtInfo)
    Else
        cmdNew.Caption = "새 공지사항(&N)"
        
        Call ShowData
    End If
    
    dtpDate(2) = Now
    grdPerson(1).Rows = grdPerson(1).FixedRows
    
    ' 해당 날짜에 기존 공지사항 있을경우..
    With grdInfo
        If CheckDate() Then
            For iLoop = 0 To .Rows - 1
                If (MakeDate(DF_LONG, dtpDate(2)) = .TextMatrix(iLoop, 1)) Then
                    txtInfo(0) = .TextMatrix(iLoop, 2)
                    
                    Exit Sub
                End If
            Next iLoop
        End If
    End With
End Sub

Private Sub cmdSave_Click()
    
    If (MsgBox("변경된 내용을 저장하시겠습니까?", vbYesNo + vbQuestion, "공지사항 저장") = vbYes) Then
      'Yes 누른 경우...
        If SaveData() Then
            Call FillGridInfo
        End If
    
    Else
        Call FillGridInfo
        ' no 선택.
        
    End If
    
    grdInfo.Enabled = True
    grdInfoUser.Enabled = True
    
    m_bFlag = False
    pnlMsg.Visible = False
    cmdNew.Caption = "새 공지사항(&N)"

End Sub

Private Sub dtpDate_Change(Index As Integer)
    Dim iLoop As Integer
    
'    If (Index = 0 Or Index = 1) Then
'        Call FillGridInfo
'
    If Index = 2 Then
        With grdInfo
            If m_bFlag Then
                dtpDate(1) = MakeDate(DF_LONG, dtpDate(2))
                Call FillGridInfo
                dtpDate(2) = MakeDate(DF_LONG, dtpDate(1))
                For iLoop = 0 To .Rows - 1 '기존 입력된 날짜와 같은 날짜가 있다면..
                    If (MakeDate(DF_LONG, dtpDate(2)) = .TextMatrix(iLoop, 1)) Then
                        
                        txtInfo(0) = grdInfo.TextMatrix(iLoop, 2)
                        grdInfo.Select iLoop, 1  '해당 날짜 row에 Select...
                        grdPerson(1).Rows = grdPerson(1).FixedRows
                        txtInfo(1) = ""
                    
                        Exit Sub
                    
                    End If
                        
                Next iLoop
                
                Call ClearText(txtInfo)  '같은 날짜 없으면 입력창 지우기..
                grdInfoUser.Rows = grdInfoUser.FixedRows
                grdPerson(1).Rows = grdPerson(1).FixedRows
                
            End If
        End With
        
    End If
End Sub


Private Function CheckDate() As Boolean
    Dim oInfo As PlusLib2.CInfo
    Dim rs As ADODB.Recordset
    
    CheckDate = True
    If Not m_bFlag Then Exit Function
    
    Set oInfo = New PlusLib2.CInfo
    oInfo.Connection = g_adoCon
    Set rs = oInfo.CheckDate(MakeDate(DF_SHORT, dtpDate(2)))  '같은 날짜 있는지 확인..
    Set oInfo = Nothing
    
    If rs.RecordCount <> 0 Then  '같은 날짜 데이터가 있으면
        CheckDate = True
    Else
        CheckDate = False
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub FillGridInfo()
    Dim oInfo  As PlusLib2.CInfo
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    
    On Error GoTo ErrHandler
    
    Set oInfo = New PlusLib2.CInfo
    oInfo.Connection = g_adoCon
    Set rs = oInfo.GetInfoByDate(MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)))
    Set oInfo = Nothing

    If rs.RecordCount = 0 Then
        grdInfo.Rows = grdInfo.FixedRows
        grdInfo.HighLight = flexHighlightNever
        Call ClearText(txtInfo)
        lblCount.Caption = LoadResString(250)
        
        Exit Sub
    End If
    
    With grdInfo
        .Redraw = False
        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & MakeDate(DF_LONG, rs!InfoDate) & vbTab & CheckNull(rs!Info)
                '' 수정... 개인공지 컬럼 삭제
            rs.MoveNext
        Loop
    
        lblCount.Caption = LoadResString(250) & grdInfo.Rows - 1 & " 건"
        rs.Close
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
           .Col = .FixedCols
           .ColSel = .Cols - 1

           Call ShowData  '' 개인공지 설정 목록 출력
        End If
        
        .Redraw = True
        .Row = .Rows - 1 ' 가장 마지막 row를 선택..
    End With
    
    Exit Sub

ErrHandler:
    Set oInfo = Nothing

    Call ErrorBox(Err.Number, "InfoSet.FillGridInfo", Err.Description)
    Err.Clear
End Sub
    
Private Sub ShowData()
    Dim oInfo As PlusLib2.CInfo
    Dim rs As ADODB.Recordset
    
    Dim content As String
    
    On Error GoTo ErrHandler
    
    If grdInfo.Rows = grdInfo.FixedRows Then
        Exit Sub
    End If
    
    With grdInfo
        dtpDate(2) = .TextMatrix(.Row, 1)
        txtInfo(0) = .TextMatrix(.Row, 2)
    End With

    Set oInfo = New PlusLib2.CInfo
    oInfo.Connection = g_adoCon
    Set rs = oInfo.GetPersonInfoList(MakeDate(DF_SHORT, dtpDate(2))) '개인 공지사항 목록
    Set oInfo = Nothing
   
    If rs.RecordCount = 0 Then '데이터가 없는 경우 각 그리드와 텍스트창 초기화..
        txtInfo(1) = ""
        grdInfoUser.Rows = grdInfoUser.FixedRows
        grdPerson(1).Rows = grdPerson(1).FixedRows
        
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    ' 데이터가 있는 경우...
    With grdInfoUser '해당 날짜의 개인별 공지사항 목록 출력..
        .Redraw = False
        .Rows = .FixedRows
        
        Do Until rs.EOF  ' 그리드에 출력...
            content = rs!Info
            If Len(content) > 15 Then
                content = Left(content, 15) & "..."
                If (InStr(content, vbCrLf)) > 0 Then
                    content = Left(content, InStr(content, vbCrLf))
                End If
            End If
            
            .AddItem .Rows & vbTab & rs!InfoDate & vbTab & rs!InfoSeq & vbTab & rs!Info & vbTab & content
            rs.MoveNext
        Loop
        .Redraw = True
        
     '   .Select 0, 0 'grdPerson에 초기 정보 주자...
        
        .Select 1, 3
        
    End With
    
    rs.Close
    Set rs = Nothing
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "SetInfo.ShowData", Err.Description)
    
End Sub

Private Sub grdPerson_DblClick(Index As Integer)
    With grdPerson(0)
        If .Row < 1 Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub


Private Sub grdInfo_RowColChange()
    Call ShowData
End Sub


Private Sub DoFlexGridGroup(iRow As Integer, iLvl As Integer)
    With grdPerson(0)
        ' Set the row as a group
        .IsSubtotal(iRow) = True

        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = vbBlue
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = COLOR_GRIDROW
            '.Cell(flexcpChecked, iRow, 0) = flexUnchecked
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        End Select
    End With
End Sub

Private Sub CheckedPerson()
    Dim iRow%, iNowRow%
    Dim i As Integer
    Dim itemCheck As Boolean ' 동일 ID 있는지 확인..
    Dim temp1, temp2 As String
    
   
    With grdPerson(0)
        If .IsSubtotal(.Row) Then  '부서 row 일 경우 부서내의 모든 사람을 이동시킴..
            For iRow = .Row + 1 To .Rows - 1
                itemCheck = False
                For i = 0 To grdPerson(1).Rows - 1 ' 기존 사원ID와 새로 입력할 사원ID 비교 같으면 삽입 안함.
                    If (.TextMatrix(iRow, 4) = grdPerson(1).TextMatrix(i, 2)) Then
                        itemCheck = True
                        Exit For
                    End If
                Next i
                
                If Not itemCheck Then
                        If .IsSubtotal(iRow) Then Exit For
                        grdPerson(1).AddItem grdPerson(1).Rows & vbTab & .TextMatrix(iRow, 2) & vbTab & .TextMatrix(iRow, 4)
                End If
                
            Next iRow
        
        Else '부서 row가 아닌 사원 선택시..
            For i = 0 To grdPerson(1).Rows - 1 ' 기존 사원ID와 새로 입력할 사원ID 비교 같으면 삽입 안함.
                If (.TextMatrix(.Row, 4) = grdPerson(1).TextMatrix(i, 2)) Then
                    itemCheck = True
                    Exit For
                End If
            Next i
            If Not itemCheck Then
                grdPerson(1).AddItem grdPerson(1).Rows & vbTab & .TextMatrix(.Row, 2) & vbTab & .TextMatrix(.Row, 4)
            End If
            
        End If
            
    End With
End Sub

Private Sub FillGridPerson()
    Dim oPerson As PlusLib2.CPerson
    Dim rs As ADODB.Recordset
    Dim iLoop%, iTop%, iRow%
    
    Set oPerson = New PlusLib2.CPerson
    
    oPerson.Connection = g_adoCon
    Set rs = oPerson.GetPerson()
    
    Screen.MousePointer = flexHourglass
    With grdPerson(0)
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        Do Until rs.EOF
            If rs!DepartID <> .TextMatrix(.Rows - 1, 3) Then
                .AddItem "" & vbTab & rs!Depart & vbTab & "" & vbTab & _
                    rs!DepartID & vbTab & ""
                
                Call DoFlexGridGroup(.Rows - 1, 1)
                iTop = .Rows - 1
            End If
             
             ' 사원이름, 부서명, 개인ID, 개인 공지사항..
            .AddItem "" & vbTab & "" & vbTab & rs!Name & vbTab & rs!DepartID & vbTab & rs!PersonID
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
   '     Call ChangeScroll(0)
        
        .Redraw = flexRDDirect
    End With
    Screen.MousePointer = flexDefault
End Sub

Private Function CheckCount() As Integer
    CheckCount = grdPerson(1).Rows - 1
End Function

Private Sub grdInfoUser_RowColChange()
    Dim oPerson As PlusLib2.CPerson
    Dim oInfo As PlusLib2.CInfo  ''개인공지
    Dim rs As ADODB.Recordset
    Dim InfoNum As String
    
    If grdInfoUser.Rows = grdInfoUser.FixedRows Then
        Exit Sub
    End If

    Set oInfo = New PlusLib2.CInfo '개인 공지
    
    oInfo.Connection = g_adoCon
    
    InfoNum = grdInfoUser.TextMatrix(grdInfoUser.Row, 2)
   
    Set rs = oInfo.GetPersonInfoID(MakeDate(DF_SHORT, dtpDate(2)), val(InfoNum))
    '해당 개인 공지별 설정된 사람들 이름과 ID
    
    With grdInfoUser
        txtInfo(1) = .TextMatrix(.Row, 3)
    End With
    
    With grdPerson(1)
        .Redraw = flexRDNone
        .Rows = .FixedRows
        Do Until rs.EOF
            .AddItem .Rows & vbTab & rs!Name & vbTab & rs!PersonID
            rs.MoveNext
        Loop
        .Redraw = True
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorHandler:
    Call ErrorBox(Err.Number, "SetInfo.grdInfoUser_RowColChange", Err.Description)
    
End Sub

