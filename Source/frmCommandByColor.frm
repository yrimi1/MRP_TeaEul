VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCommandByColor 
   Caption         =   "Order별 진행현황"
   ClientHeight    =   9255
   ClientLeft      =   3435
   ClientTop       =   2535
   ClientWidth     =   15240
   Icon            =   "frmCommandByColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15240
   Begin Threed.SSCommand cmdClose 
      Height          =   1305
      Left            =   12930
      TabIndex        =   48
      Top             =   6675
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   2302
      _Version        =   196609
      Caption         =   "투입마감"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   6075
      TabIndex        =   43
      Top             =   450
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optClose 
         Caption         =   "구분 안함"
         Height          =   180
         Index           =   1
         Left            =   1350
         TabIndex        =   45
         Top             =   75
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optClose 
         Caption         =   "마감건 구분"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   44
         Top             =   75
         Width           =   1365
      End
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   11055
      TabIndex        =   41
      Top             =   90
      Width           =   1965
   End
   Begin Threed.SSFrame frmCommand 
      Height          =   2910
      Left            =   15
      TabIndex        =   27
      Top             =   5580
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   5133
      _Version        =   196609
      Caption         =   "    지시 내역    "
      Begin Threed.SSCommand cmdSave 
         Height          =   645
         Left            =   11055
         TabIndex        =   31
         Top             =   405
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1138
         _Version        =   196609
         Caption         =   "저장"
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   645
         Left            =   9735
         TabIndex        =   30
         Top             =   405
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1138
         _Version        =   196609
         Caption         =   "추가"
      End
      Begin Threed.SSPanel pnlInput 
         Height          =   1665
         Left            =   6915
         TabIndex        =   29
         Top             =   1110
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   2937
         _Version        =   196609
         Enabled         =   0   'False
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtLoss 
            Height          =   315
            Left            =   1410
            TabIndex        =   46
            Top             =   1215
            Width           =   3870
         End
         Begin VB.TextBox txtPerson 
            Height          =   315
            Left            =   1410
            TabIndex        =   37
            Top             =   840
            Width           =   3555
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Index           =   0
            Left            =   105
            TabIndex        =   34
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "색       상"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.TextBox txtQty 
            Height          =   315
            Left            =   1410
            TabIndex        =   33
            Top             =   480
            Width           =   3885
         End
         Begin VB.ComboBox cboColor 
            Height          =   300
            Left            =   1410
            TabIndex        =   32
            Top             =   120
            Width           =   3900
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Index           =   1
            Left            =   105
            TabIndex        =   35
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "준비 수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Index           =   2
            Left            =   105
            TabIndex        =   36
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "작  성  자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   1
            Left            =   5010
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   840
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            Enabled         =   0   'False
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Index           =   3
            Left            =   105
            TabIndex        =   47
            Top             =   1215
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "적용 Loss"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDetail 
         Height          =   2520
         Left            =   135
         TabIndex        =   28
         Top             =   255
         Width           =   6675
         _cx             =   11774
         _cy             =   4445
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
   Begin VSFlex7LCtl.VSFlexGrid grdColor 
      Height          =   1770
      Left            =   0
      TabIndex        =   26
      Top             =   3735
      Width           =   15225
      _cx             =   26855
      _cy             =   3122
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
   Begin VB.Frame fraOrder 
      Height          =   765
      Left            =   60
      TabIndex        =   22
      Top             =   -15
      Width           =   1410
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   210
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   1200
      End
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   2190
      TabIndex        =   19
      Top             =   2175
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
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   120
         Width           =   270
      End
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   11070
      TabIndex        =   11
      Top             =   435
      Width           =   1965
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   7335
      TabIndex        =   10
      Top             =   90
      Width           =   1965
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   720
      Left            =   14430
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   13
      ToolTipText     =   "자료 저장"
      Top             =   60
      Width           =   780
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2205
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   405
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1560
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   405
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2205
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1560
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   4185
      TabIndex        =   4
      Top             =   90
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   4185
      TabIndex        =   5
      Top             =   435
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2895
      TabIndex        =   8
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "수주일자"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   9825
      TabIndex        =   14
      Top             =   435
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   15
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   6090
      TabIndex        =   16
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "Order No"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   45
         Width           =   1095
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   0
      Left            =   13080
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      Enabled         =   0   'False
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13545
      TabIndex        =   18
      Top             =   8550
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   2850
      Left            =   0
      TabIndex        =   25
      Top             =   825
      Width           =   15225
      _cx             =   26855
      _cy             =   5027
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
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   3
      Left            =   9825
      TabIndex        =   39
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   40
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   3
      Left            =   13080
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      Enabled         =   0   'False
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   0
      Left            =   5505
      TabIndex        =   7
      Top             =   165
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   1
      Left            =   5505
      TabIndex        =   6
      Top             =   510
      Width           =   360
   End
End
Attribute VB_Name = "frmCommandByColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH1 = 2000
Private Const LIMIT_WIDTH2 = 1355
Private Const LIMIT_ROW1 = 8
Private Const LIMIT_ROW2 = 18

Private Const REPORTFILE = "\Report\ResultByOrder.rpt"

Dim m_bSkipColor As Boolean
Dim m_bSkipOrder As Boolean
Dim m_sMode As Boolean



Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then '[0] 수주일자 선택
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else '[1, 2] 거래처, 관리번호 선택
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 2 Then
                cmdFind(0).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 2 Then
                cmdFind(1).Enabled = False
            End If
        End If
    End If

End Sub

Private Sub ClearData()
    txtQty = ""
    txtPerson = ""
    txtPerson.Tag = ""
    txtLoss = ""
    pnlInput.Enabled = False
    cmdSave.Caption = "추가"

End Sub

Private Sub cmdAdd_Click()
    If m_sMode = False Then
        pnlInput.Enabled = True
        m_sMode = True
        cmdAdd.Caption = "취소"
        cmdFind(1).Enabled = True
        txtQty.SetFocus
    Else
        pnlInput.Enabled = False
        m_sMode = False
        cmdAdd.Caption = "추가"
        cmdFind(1).Enabled = False
        Call ClearData
    End If
End Sub



Private Sub cmdClose_Click()
    Dim oCommand As PlusLib2.CCommand
    Dim bResult As Boolean
    Dim sOrderID$, nChkClose%
    
    On Error GoTo ErrHandler
    
    sOrderID = MakeOrderID(grdData.TextMatrix(grdData.Row, 4), OM_REDUCE)
    
    ' 현재 종료구분 설정값
    If grdData.TextMatrix(grdData.Row, 4) = "■" Then
        nChkClose = 1
    Else
        nChkClose = 0
    End If
    
    Set oCommand = New PlusLib2.CCommand
    
    bResult = oCommand.UpdateStuffClose(sOrder, MakeDate(DF_SHORT, Now), nChkClose)
    
    Set oCommand = Nothing
    
    If bResult = True Then
        If nChkClose = 0 Then
            MessageBox "투입마감 처리되었습니다"
        Else
            MessageBox "투입마감이 해제되었습니다"
        End If
    
    Else
        If nChkClose = 0 Then
            MessageBox "투입마감이 실패했습니다"
        Else
            MessageBox "투입마감 해제에 실패했습니다"
        End If
    End If
    
Exit Sub

ErrHandler:
    Set oCommand = Nothing
    
    Call ErrorBox(Err.Number, "frmCommandByColor.cmdClose_Click", Err.Description)
End Sub

Private Sub cmdExit_Click()
    PlusMDI.pnlMenu.Visible = True
    
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 0 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(2))
    Else
        Call ReturnCode(LG_PERSON, , False, txtPerson)
    End If
End Sub



Private Sub cmdSave_Click()
    Dim oCommand As PlusLib2.CCommand
    Dim NewCommand As TCommand
    Dim bSave As Boolean
    
    On Error GoTo ErrHandler

   
    With NewCommand
        .sOrderID = MakeOrderID(grdData.TextMatrix(grdData.Row, 4), OM_REDUCE)
        .nOrderIDSeq = 0
        .nOrderSeq = cboColor.ItemData(cboColor.ListIndex)
        .Instdate = MakeDate(DF_SHORT, Now)
        .InstQty = txtQty
        .nApplyLoss = txtLoss
        .sPersonID = txtPerson.Tag
    
    End With
   
   
    Set oCommand = New PlusLib2.CCommand
    oCommand.Connection = g_adoCon
    oCommand.UserName = g_sUserName


    bSave = oCommand.AddNewCommand(NewCommand)
    
    If bSave = True Then
        MessageBox "저장되었습니다"
        Call cmdAdd_Click
        Call FillGridData
    Else
        MessageBox "저장에 실패했습니다"
    End If
    
    
    Set oCommand = Nothing

    Exit Function

ErrHandler:
    Set oCommand = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub



Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then       ' 전월
        dtpDate(0) = DateSerial(Year(Date), Month(Date) - 1, 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date), 1 - 1)
    ElseIf Index = 1 Then   ' 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = Date
    ElseIf Index = 2 Then   ' 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 3 Then   ' 금년
        dtpDate(0) = DateSerial(Year(Date), 1, 1)
        dtpDate(1) = Now
    End If

    cmdSearch.SetFocus

End Sub

Private Sub Form_Load()
    PlusMDI.pnlMenu.Visible = False

    Me.Move 0, 0, 15420, 9660
    
    Call InitGrid
    
    Call SetOperate(Me)
        
    dtpDate(0) = Now
    dtpDate(1) = Now
    
    chkSearch(0) = vbChecked
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    pnlProgress.Visible = False
    m_bSkipOrder = False
End Sub


Private Sub FillGridData()
    Dim oCommand As PlusLib2.CCommand
    Dim rs As ADODB.Recordset
    Dim lNowRow&, i%
    Dim nChkDate%, sDate$, eDate$
    Dim nChkCustom%, sCustomID$
    Dim nChkArticle%, sArticle$
    Dim nChkOrder%, sOrder$
    Dim nChkClose$
    Dim sClose$
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    nChkDate = IIf(chkSearch(0).Value = vbChecked, 1, 0)
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkCustom = IIf(chkSearch(2).Value = vbChecked, 1, 0)
    sCustomID = txtSearch(2).Tag
    nChkArticle = IIf(chkSearch(3).Value = vbChecked, 1, 0)
    sArticle = txtSearch(3).Tag
    nChkOrder = IIf(chkSearch(1).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0)
    sOrder = txtSearch(1)
    nChkClose = IIf(optClose(1).Value = True, 1, 0)
    
    Set oCommand = New PlusLib2.CCommand
    oCommand.Connection = g_adoCon
    
    Set rs = oCommand.GetCommandByOrder(nChkDate, sDate, eDate, nChkCustom, sCustomID, nChkArticle, sArticleID, _
                                            nChkOrder, sOrder, nChkClose)
    
    Set oCommand = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        grdData.Rows = 1
        grdData.HighLight = flexHighlightNever
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
    m_bSkipOrder = True
    
    nRow = rs.RecordCount
    
    With grdData
        .Redraw = False
        
        If .Rows > .FixedRows Then
            lNowRow = .Row
            .Rows = 1
        Else
            lNowRow = 1
        End If
        
        For i = 1 To nRow
            DoEvents
            
            If IsNull(rs!StuffCloseClss) Then
                sClose = " "
            Else
                sClose = "■"
            End If
            
            .AddItem CStr(i) & vbTab & sClose & vbTab & rs!KCustom & vbTab & rs!Article & vbTab & rs!OrderID & vbTab & _
                    rs!OrderNO & vbTab & MakeDate(DF_LONG, rs!DvlyDate) & vbTab & rs!ColorCnt & vbTab & SetCurrency(rs!OrderQty) & vbTab & _
                    rs!LossRate & vbTab & "  " & vbTab & "  " & vbTab & SetCurrency(rs!InstQty) & vbTab & " " & vbTab & _
                    SetCurrency(rs!WorkQty) & vbTab & SetCurrency(rs!InstQty - rs!WorkQty) & vbTab & SetCurrency(rs!StuffINQty) & vbTab & _
                    SetCurrency(rs!SetQty) & vbTab & SetCurrency(rs!StuffINQty - rs!SetQty)
                           
            lblCount = CStr(i) & " / " & CStr(nRow) & "  (" & Format((i / nRow) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / nRow) * 100)

            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If

            rs.MoveNext
        Next i
        
        If .Rows > .FixedRows Then
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            .HighLight = flexHighlightAlways
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            Call FillGridColor
        End If
        .Redraw = True
    End With
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    pnlProgress.Visible = False
    m_bSkipOrder = False
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oCommand = Nothing
    
    pnlProgress.Visible = False
    Screen.MousePointer = vbDefault
    m_bSkipOrder = False
    
    Call ErrorBox(Err.Number, "frmCommandByColor.FillGridData", Err.Description)
End Sub




Private Sub FillGridColor()
    Dim oCommand As PlusLib2.CCommand
    Dim rs As ADODB.Recordset
    Dim lNowRow&, i%
    Dim sOrder$
    
    On Error GoTo ErrHandler
    
    If m_bSkipOrder = True Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    sOrder = grdData.TextMatrix(grdData.Row, 4)
    
    Set oCommand = New PlusLib2.CCommand
    oCommand.Connection = g_adoCon
    
    Set rs = oCommand.GetCommandByColor(sOrder)
    
    Set oCommand = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        grdColor.Rows = 1
        grdColor.HighLight = flexHighlightNever
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    m_bSkipColor = True
    
    With grdColor
        .Redraw = False
        
        cboColor.Clear
        
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!Color & SetCurrency(rs!ColorQty) & vbTab & " " & vbTab & " " & vbTab & _
                        SetCurrency(rs!InstQty) & vbTab & " " & vbTab & SetCurrency(rs!WorkQty) & vbTab & _
                        SetCurrency(rs!InstQty - rs!WorkQty) & vbTab & "-" & vbTab & "-" & vbTab & SetCurrency(rs!SetQty) & vbTab & _
                        rs!OrderSeq
            
            ' 색상 콤보 설정
            cboColor.AddItem rs!Color
            cboColor.ItemData(cboColor.NewIndex) = rs!OrderSeq
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If

            rs.MoveNext
        Next i
        
        If .Rows > .FixedRows Then
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            .HighLight = flexHighlightAlways
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
        End If
        .Redraw = True
    End With
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    m_bSkipColor = False
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oCommand = Nothing
    
    pnlProgress.Visible = False
    Screen.MousePointer = vbDefault
    m_bSkipColor = False
    
    Call ErrorBox(Err.Number, "frmCommandByColor.FillGridColor", Err.Description)
End Sub


Private Sub FillGridDetail()
    Dim oCommand As PlusLib2.CCommand
    Dim rs As ADODB.Recordset
    Dim lNowRow&, i%
    Dim sOrder$, sColorID$
    
    On Error GoTo ErrHandler
    
    If m_bSkipColor = True Then Exit Sub
    
    If grdColor.Row = grdColor.FixedRows Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    sOrder = grdData.TextMatrix(grdData.Row, 4)
    sColorID = grdColor.TextMatrix(grdColor.Row, 13)
    
    Set oCommand = New PlusLib2.CCommand
    oCommand.Connection = g_adoCon
    
    Set rs = oCommand.GetCommandByDate(sOrder, sColorID)
    
    Set oCommand = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        grdColor.Rows = 1
        grdColor.HighLight = flexHighlightNever
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With grdDetail
        .Redraw = False
        'A.OrderID, A.OrderIDSeq, A.OrderSeq, B.Color, A.InstDate, A.InstQty, ApplyLoss, A.PersonID, C.Name
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & MakeDate(DF_LONG, rs!Instdate) & vbTab & rs!Color & vbTab & SetCurrency(rs!InstQty) & vbTab & _
                        rs!ApplyLoss & vbTab & rs!Name
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If

            rs.MoveNext
        Next i
        
        If .Rows > .FixedRows Then
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            .HighLight = flexHighlightAlways
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
        End If
        .Redraw = True
    End With
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oCommand = Nothing
    
    pnlProgress.Visible = False
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmCommandByColor.FillGridDetail", Err.Description)
End Sub



Private Sub ShowData()
    Dim i%
    Dim oCommand As PlusLib2.CCommand
    Dim rs As Recordset
    Dim nRoll(9) As Integer, nQty(9) As Long
    
    'On Error GoTo ErrHandler
    
    Set oCommand = New PlusLib2.CCommand
    oCommand.Connection = g_adoCon
    Set rs = oCommand.GetResultOrder(Replace(grdData.TextMatrix(grdData.Row, 1), "-", ""))
    Set oCommand = Nothing
    
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        grdColor.Rows = grdData.FixedRows
        cmdPrint.Enabled = False
       
        Call ChangeScroll(1)
        Exit Sub
    End If
    cmdPrint.Enabled = True
    With grdColor
        .Redraw = False
        .Rows = 1
        For i = 1 To rs.RecordCount
            .Rows = i + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = MakeDate(DF_LONG, rs!TotalDate)
            .TextMatrix(i, 2) = IIf(rs!R1 = 0, "", rs!R1 & " *" & SetCurrency(rs!Q1))
            .TextMatrix(i, 3) = IIf(rs!R2 = 0, "", rs!R2 & " *" & SetCurrency(rs!Q2))
            .TextMatrix(i, 4) = IIf(rs!R3 = 0, "", rs!R3 & " *" & SetCurrency(rs!Q3))
            .TextMatrix(i, 5) = IIf(rs!R4 = 0, "", rs!R4 & " *" & SetCurrency(rs!Q4))
            .TextMatrix(i, 6) = IIf(rs!R5 = 0, "", rs!R5 & " *" & SetCurrency(rs!Q5))
            .TextMatrix(i, 7) = IIf(rs!Q6 = 0, "", "? *" & SetCurrency(rs!Q6))
            .TextMatrix(i, 8) = IIf(rs!R7 = 0, "", rs!R7 & " *" & SetCurrency(rs!Q7))
            .TextMatrix(i, 9) = IIf(rs!R8 = 0, "", rs!R8 & " *" & SetCurrency(rs!Q8))
            .TextMatrix(i, 10) = IIf(rs!R9 = 0, "", rs!R9 & " *" & SetCurrency(rs!Q9))

            nRoll(0) = nRoll(0) + rs!R1
            nRoll(1) = nRoll(1) + rs!R2
            nRoll(2) = nRoll(2) + rs!R3
            nRoll(3) = nRoll(3) + rs!R4
            nRoll(4) = nRoll(4) + rs!R5
            nRoll(5) = nRoll(5) + rs!R6
            nRoll(6) = nRoll(6) + rs!R7
            nRoll(7) = nRoll(7) + rs!R8
            nRoll(8) = nRoll(8) + rs!R9

            nQty(0) = nQty(0) + rs!Q1
            nQty(1) = nQty(1) + rs!Q2
            nQty(2) = nQty(2) + rs!Q3
            nQty(3) = nQty(3) + rs!Q4
            nQty(4) = nQty(4) + rs!Q5
            nQty(5) = nQty(5) + rs!Q6
            nQty(6) = nQty(6) + rs!Q7
            nQty(7) = nQty(7) + rs!Q8
            nQty(8) = nQty(8) + rs!Q9
            
            rs.MoveNext
        Next i
        .Redraw = True
    End With
    rs.Close
    Set rs = Nothing
    
    
    
    Exit Sub

ErrHandler:
 
    Set rs = Nothing
    Set oCommand = Nothing
    
    Call ErrorBox(Err.Number, "frmCommandByColor.ShowData", Err.Description)
End Sub


Private Sub InitGrid()
    Dim nWidth%, i%
    
    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = False
        .Rows = 1
        .Cols = 20
        .FixedCols = 8
        .ScrollBars = flexScrollBarBoth

        .TextArray(0) = " "
        .TextArray(1) = "마감구분":                     .ColWidth(1) = 300:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "거래처":                       .ColWidth(2) = 1300:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":                         .ColWidth(3) = 1200:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "수주번호":                     .ColWidth(4) = 1200:            .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "OrderNO":                      .ColWidth(5) = 1200:            .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "납기":                         .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "색상수":                       .ColWidth(7) = 550:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "수주량":                       .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "Loss":                         .ColWidth(9) = 600:             .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "실제" & vbCrLf & "Loss":      .ColWidth(10) = 800:            .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(11) = "수주량" & vbCrLf & "*Loss":   .ColWidth(11) = 800:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "준비" & vbCrLf & "수량":      .ColWidth(12) = 800:            .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "과부족":                      .ColWidth(13) = 800:            .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "생산" & vbCrLf & "지시량":    .ColWidth(14) = 800:            .ColAlignment(14) = flexAlignCenterCenter
        .TextArray(15) = "예비량":                      .ColWidth(15) = 800:            .ColAlignment(15) = flexAlignCenterCenter
        .TextArray(16) = "입고량":                      .ColWidth(16) = 800:            .ColAlignment(16) = flexAlignCenterCenter
        .TextArray(17) = "생지" & vbCrLf & "재고":      .ColWidth(17) = 800:            .ColAlignment(17) = flexAlignCenterCenter
        .TextArray(18) = "연폭량":                      .ColWidth(18) = 800:            .ColAlignment(18) = flexAlignCenterCenter
        .TextArray(19) = "실재고":                      .ColWidth(19) = 800:            .ColAlignment(19) = flexAlignCenterCenter
        
        .Redraw = True
    End With
    
    Call SetVSFlexGrid(grdColor)
    With grdColor
        .Redraw = False
        .Rows = 1
        .Cols = 14
        .FixedCols = 2
        .ScrollBars = flexScrollBarBoth

        .TextArray(0) = " "
        .TextArray(1) = "색 상 명":                     .ColWidth(1) = 5500:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "수주량":                       .ColWidth(2) = 800:             .ColAlignment(3) = flexAlignRightCenter
        .TextArray(3) = " ":                            .ColWidth(3) = 1200:            .ColAlignment(2) = flexAlignRightCenter
        .TextArray(4) = "수주량" & vbCrLf & "*Loss":    .ColWidth(4) = 800:             .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "준비" & vbCrLf & "수량":       .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "과부족":                       .ColWidth(6) = 800:              .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "생산" & vbCrLf & "지시량":     .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "예비량":                       .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = " - ":                          .ColWidth(9) = 800:             .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = " - ":                         .ColWidth(10) = 800:            .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(11) = "연폭량":                      .ColWidth(11) = 800:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = " - ":                         .ColWidth(12) = 800:            .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "ColorID":                     .ColWidth(13) = 0

        .Redraw = True
    End With
            
    Call SetVSFlexGrid(grdDetail)
    With grdDetail
        .Redraw = False
        .Rows = 1
        .Cols = 6

        .TextArray(0) = " "
        .TextArray(1) = "지시일자":                     .ColWidth(1) = 1300:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":                       .ColWidth(2) = 1900:            .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "계획량":                       .ColWidth(3) = 900:             .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "적용Loss":                     .ColWidth(4) = 900:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "작성자":                       .ColWidth(5) = 900:             .ColAlignment(5) = flexAlignCenterCenter
        
        .Redraw = True
    End With
        
End Sub


Private Sub grdColor_Click()
    If Not m_bSkipColor Then
        Call FillGridDetail
    End If
    
End Sub

Private Sub grdData_RowColChange()
    If Not m_bSkipOrder Then

        Call FillGridColor
    End If
End Sub


Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkSearch(1).Caption = "Order No"
    Else
        chkSearch(1).Caption = "관리 번호"
    End If
End Sub


Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call ReturnCode(LG_PERSON, , False, txtPerson)
    End If
End Sub


Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub


Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Index = 2 Then Call cmdFind_Click
    KeyAscii = KeyPress(txtSearch(Index), KeyAscii)
End Sub

