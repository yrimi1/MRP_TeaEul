VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderClose 
   ClientHeight    =   9255
   ClientLeft      =   105
   ClientTop       =   765
   ClientWidth     =   15240
   Icon            =   "frmOrderClose.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15240
   Begin VB.ComboBox cboFiber 
      Height          =   300
      Left            =   10470
      Style           =   2  '드롭다운 목록
      TabIndex        =   41
      Top             =   480
      Width           =   1545
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   2
      Left            =   7245
      TabIndex        =   37
      Top             =   480
      Width           =   1395
   End
   Begin VB.ComboBox cboOrderClss 
      Height          =   300
      Left            =   13140
      Style           =   2  '드롭다운 목록
      TabIndex        =   35
      Top             =   480
      Width           =   1125
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   1800
      Top             =   8625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11845
      TabIndex        =   11
      Top             =   8550
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   345
      Left            =   30
      TabIndex        =   33
      Top             =   8115
      Width           =   15135
      _cx             =   26696
      _cy             =   609
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
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   3
      Left            =   10470
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1470
      MousePointer    =   99  '사용자 정의
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2130
      MousePointer    =   99  '사용자 정의
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   90
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1470
      MousePointer    =   99  '사용자 정의
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2130
      MousePointer    =   99  '사용자 정의
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   795
      Left            =   14370
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   9
      ToolTipText     =   "자료 저장"
      Top             =   45
      Width           =   780
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   7245
      TabIndex        =   6
      Top             =   120
      Width           =   1395
   End
   Begin VB.ComboBox cboSearch 
      Height          =   300
      Left            =   13140
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.Frame fraOrder 
      Height          =   810
      Left            =   60
      TabIndex        =   23
      Top             =   -15
      Width           =   1305
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   510
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "선택 해제"
      Height          =   315
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   8910
      Width           =   1140
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "전체 선택"
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   8535
      Width           =   1140
   End
   Begin Threed.SSCommand cmdClose 
      Height          =   690
      Left            =   4680
      TabIndex        =   10
      Top             =   8520
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      확인(&C)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13530
      TabIndex        =   12
      Top             =   8550
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   15195
      _cx             =   26802
      _cy             =   12726
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
      Begin Threed.SSPanel pnlProgress 
         Height          =   870
         Left            =   2235
         TabIndex        =   20
         Top             =   2550
         Visible         =   0   'False
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   1535
         _Version        =   196609
         Alignment       =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin MSComctlLib.ProgressBar proProgress 
            Height          =   390
            Left            =   90
            TabIndex        =   21
            Top             =   375
            Width           =   10485
            _ExtentX        =   18494
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
            TabIndex        =   22
            Top             =   120
            Width           =   270
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   2640
         Left            =   6165
         TabIndex        =   17
         Top             =   4275
         Width           =   7920
         _cx             =   13970
         _cy             =   4657
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
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   5865
      TabIndex        =   26
      Top             =   120
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
         TabIndex        =   5
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   8700
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   150
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   9105
      TabIndex        =   28
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "관리번호"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   7
         Top             =   45
         Width           =   1185
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   4050
      TabIndex        =   3
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116785153
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   4050
      TabIndex        =   4
      Top             =   465
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116785153
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2805
      TabIndex        =   29
      Top             =   120
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
         TabIndex        =   2
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   12120
      TabIndex        =   30
      Top             =   120
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "수주 상태"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdGridPrint 
      Height          =   690
      Left            =   10160
      TabIndex        =   34
      Top             =   8550
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "화면인쇄"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   3
      Left            =   12135
      TabIndex        =   36
      Top             =   480
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "수주 구분"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   4
      Left            =   5865
      TabIndex        =   38
      Top             =   480
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   39
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   2
      Left            =   8700
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   480
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   5
      Left            =   9090
      TabIndex        =   42
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "원단 구분"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   8460
      TabIndex        =   43
      Top             =   8550
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   3
      Left            =   11730
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   5
      Left            =   5340
      TabIndex        =   32
      Top             =   540
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   4
      Left            =   5340
      TabIndex        =   31
      Top             =   210
      Width           =   360
   End
End
Attribute VB_Name = "frmOrderClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************************************************
'** System 명 : MRRPLUS2
'** Author    : Wizard
'** 작성자    :
'** 내용      :
'** 생성일자  :
'** 변경일자  :
'**------------------------------------------------------------------------------------------------
''*************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID                    , 요청 및 작업내용
'**************************************************************************************************
' 2014.12.12, 오승욱, 김대진대리 S_2014012_태을염직_01               색상이 2~11개 사이의 데이터 상세 그리드 높이가 잘림
'**************************************************************************************************


Option Explicit

Private Const LIMIT_WIDTH1 = 1140
Private Const LIMIT_ROW1 = 26

Private Const LIMIT_WIDTH2 = 1450
Private Const LIMIT_ROW2 = 8                'S_2014012_태을염직_01 에 의한 수정(OLD:9)

'---------------------------------------------------'
Private Const REPORTFILE_1 = "\Report\OrderClose.rpt"
'------------------------------------------------------------
Private m_nSelected As Integer ' 수주 선택갯수
Private m_bSkipEvent As Boolean
Private m_bloading As Boolean


Private Sub cboSearch_Click()
    If m_bloading Then Exit Sub
    
    If cboSearch.ListIndex = 1 Then
        cmdClose.Caption = "완료처리"
    Else
        cmdClose.Caption = "진행처리"
    End If
    
    Call FillGridOrder
End Sub

Private Sub cmdExcel_Click()
    If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub

    Call MakeExcelGrid(grdOrder)

End Sub

Private Sub cmdGridPrint_Click()
    Dim i%
    
    With grdOrder
        .Redraw = False
        .RowHidden(.Rows - 1) = False
        .RowHeight(.Rows - 1) = 400
                
        .FontSize = 7
        .ColWidth(3) = GetFlexColWidth(11)
        .ColWidth(4) = 1200
        .ColWidth(5) = 2500
        .ColWidth(6) = GetFlexColWidth(7)
        
'        .ColWidth(7) = 900
'        .ColWidth(8) = 900
        
        For i = 12 To 18
            .ColWidth(i) = 700
        Next i
        .PrintGrid "태을염직", True, 2, 100, 300
        
        .FontSize = 9
        .ColWidth(3) = GetFlexColWidth(13)
        .ColWidth(4) = 1140
        .ColWidth(5) = 1610
        .ColWidth(6) = GetFlexColWidth(5)
        .ColWidth(7) = 640
        .ColWidth(8) = 640

        For i = 12 To 18
            .ColWidth(i) = 900
        Next i

        .RowHidden(.Rows - 1) = True
        .Redraw = True
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
            
    Set rs = oOrder.PrintOrderClose(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
                                IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
                                IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0), txtSearch(3), _
                                cboSearch.ListIndex, cboOrderClss.ListIndex, cboFiber.ListIndex)
    Set oOrder = Nothing
    
    ReDim sParam(4)
    sParam(0) = "수주 진행 현황"
    sParam(1) = CompanyName
    If dtpDate(0) = dtpDate(1) Then
            sParam(2) = "수주일자  : " & IIf(chkSearch(0), MakeDate(DF_LONG, dtpDate(0)), "")
        Else
            sParam(2) = "수주일자  : " & MakeDate(DF_LONG, dtpDate(0)) & " ~ " & MakeDate(DF_LONG, dtpDate(1))
        End If
    sParam(3) = "거 래 처   : " & IIf(chkSearch(1), txtSearch(1), "(전체)")
    
    If optOrder(0).Value Then
        sParam(4) = "OrderNO : " & IIf(chkSearch(2), txtSearch(2), "(전체)")
    Else
        sParam(4) = "관리번호  : " & IIf(chkSearch(2), txtSearch(2), "(전체)")
    End If
        
    Call PrintReport(REPORTFILE_1, rs, sParam, PlusMDI.PrintPreview)
       
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmOrderClose.cmdPrint_Click", Err.Description)
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%

    m_bloading = True

    Me.Move 0, 0, 15360, 9660

    Call InitGrid
    Call SetOperate(Me)
    
    Show
    
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
    
    For i = 1 To 3
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
    
    cmdClose.Picture = LoadResPicture("CHECK", vbResIcon)
    cmdClose.Enabled = False
    
    m_bSkipEvent = True
    m_bloading = True
    With cboSearch
        .AddItem "전체"
        .AddItem "진행건"
        .AddItem "완료건"
        
        .ListIndex = 0
    End With
    
    With cboOrderClss
        .AddItem "전체"
        .AddItem "본작업"
        .AddItem "시가공"
        
        .ListIndex = 0
    End With
    
    With cboFiber
        .AddItem "전체"
        .AddItem "면"
        .AddItem "화섬"
        
        .ListIndex = 0
    End With
        
    grdColor.Visible = False
    m_bloading = False
    
    chkSearch(0).Value = vbChecked
    Call SetDtpDate(1, dtpDate(0), dtpDate(1))
End Sub

'Private Sub cboSearch_Click()
'    If m_bLoading Then Exit Sub
'
'    cmdClose.Caption = Space(6) & cboSearch.List(Abs(cboSearch.ListIndex - 1))
'
'    Call FillGridOrder
'End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(Index).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
            
            dtpDate(0).SetFocus
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            cmdFind(Index).Enabled = True
        Else
            txtSearch(Index).Enabled = False
            cmdFind(Index).Enabled = False
        End If
    End If
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Dim SetValue, i%
    
    If Index = 0 Then   '[0] 전체선택
        SetValue = flexChecked
        If cboSearch.ListIndex > 0 Then
            cmdClose.Enabled = True
        End If
    Else                '[1] 선택 해제
        SetValue = flexUnchecked
        If cboSearch.ListIndex > 0 Then
            cmdClose.Enabled = False
        End If
    End If

    m_nSelected = 0
    With grdOrder
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, 1) = SetValue
            
            If SetValue = flexChecked Then
                m_nSelected = m_nSelected + 1
            Else
                m_nSelected = m_nSelected - 1
            End If
        Next i
    End With
End Sub

Private Sub cmdClose_Click()
    Dim sOrderID() As String, nPoint%, i%

    Dim oOrder As PlusLib2.COrder
    
    On Error GoTo ErrHandler
    
    If m_nSelected < 1 Then Exit Sub
    
    If MsgBox(LoadResString(cboSearch.ListIndex + 224), vbQuestion + vbYesNo) = vbYes Then
        ReDim sOrderID(m_nSelected - 1)
        With grdOrder
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then
                    sOrderID(nPoint) = Replace(.TextMatrix(i, 3), "-", "")
                    
                    nPoint = nPoint + 1
                End If
            Next i
        End With
        
        Set oOrder = New PlusLib2.COrder
        oOrder.Connection = g_adoCon
        If oOrder.UpdateOrderClose(sOrderID, cboSearch.ListIndex) Then
            
        Else
            MsgBox LoadResString(cboSearch.ListIndex + 227), vbCritical
        End If
        Set oOrder = Nothing
        
        Call FillGridOrder
        
        MsgBox LoadResString(161), vbInformation
    End If
    Exit Sub
ErrHandler:
    Set oOrder = Nothing

    Call ErrorBox(Err.Number, "frmOrderClose.cmdClose_Click", Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    ElseIf Index = 3 Then
        Call ReturnCode(LG_ORDER, , False, txtSearch(3))
        
        If optOrder(1).Value Then
            txtSearch(3).Text = txtSearch(3).Tag
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

End Sub

Private Sub InitGrid()
    Dim i%, nWidth&

    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 10
        .ExtendLastCol = True
        
        .RowHeight(0) = 280
        
        .TextArray(0) = "합계 (YDS)":         .ColWidth(0) = 6550
        .TextArray(1) = "0 건":         .ColWidth(1) = 1020
        
        .ColWidth(2) = 1010
        .ColWidth(3) = 900
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 900
        .ColWidth(9) = 900
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        
        .ColFormat(2) = "#,##0"
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0"
        .ColFormat(6) = "#,##0"
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        
        .Redraw = flexRDDirect
    End With

    With grdOrder
        .Redraw = flexRDNone
        .Cols = 19
        Call SetVSFlexGrid(grdOrder)
        
        .TextArray(1) = "선택":                     .ColWidth(1) = 300:                     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order No.":                .ColWidth(2) = 0:                       .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호":                 .ColWidth(3) = GetFlexColWidth(13):     .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "거래처":                   .ColWidth(4) = 1140:            .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "품명":                     .ColWidth(5) = 1610:                    .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "납기일자":                 .ColWidth(6) = GetFlexColWidth(5):      .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "가공구분":                 .ColWidth(7) = 900:                     .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "가공폭":                   .ColWidth(8) = 640:                     .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "축율" & vbCrLf & "Loss":   .ColWidth(9) = 600:                     .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "수주" & vbCrLf & "수량":  .ColWidth(10) = GetFlexColWidth(8):     .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "단위":                    .ColWidth(11) = GetFlexColWidth(4):     .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "원단입고":  .ColWidth(12) = 900:                    .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "염색투입":  .ColWidth(13) = 900:                    .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "제품검사":  .ColWidth(14) = 900:                    .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "제직불량":  .ColWidth(15) = 900:                    .ColAlignment(15) = flexAlignRightCenter
        .TextArray(16) = "가공불량":  .ColWidth(16) = 900:                    .ColAlignment(16) = flexAlignRightCenter
        .TextArray(17) = "제품출고":  .ColWidth(17) = 900:                    .ColAlignment(17) = flexAlignRightCenter
        .TextArray(18) = "과부족":                  .ColWidth(18) = 900:                    .ColAlignment(18) = flexAlignRightCenter
        
        .ColDataType(1) = flexDTBoolean
        
        .Redraw = flexRDDirect
        .ColFormat(10) = "#,##0"
        .ColFormat(12) = "#,##0"
        .ColFormat(13) = "#,##0"
        .ColFormat(14) = "#,##0"
        .ColFormat(15) = "#,##0"
        .ColFormat(16) = "#,##0"
        .ColFormat(17) = "#,##0"
        .ColFormat(18) = "#,##0"
        
        .RowHeightMin = 390
    End With
    

    With grdColor
        .Redraw = flexRDNone
        .Move 3875, 4275
        
        .Cols = 12
        Call SetVSFlexGrid(grdColor)
    
        .TextArray(1) = "색상" & vbCrLf & "순위":   .ColWidth(1) = 450:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":                   .ColWidth(2) = 1800:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "Design No.":               .ColWidth(3) = 1100:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "수주수량":   .ColWidth(4) = 900:             .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "원단입고":   .ColWidth(5) = 900:             .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "염색투입":   .ColWidth(6) = 900:             .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "제품검사":   .ColWidth(7) = 900:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "제직불량":   .ColWidth(8) = 900:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "가공불량":   .ColWidth(9) = 900:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "제품출고":   .ColWidth(10) = 900:             .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "과부족":                   .ColWidth(11) = 900:             .ColAlignment(11) = flexAlignRightCenter
        
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0"
        .ColFormat(6) = "#,##0"
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        .ColFormat(11) = "#,##0"
        
        .ColHidden(1) = True
        For i = 1 To .Cols - 1
            nWidth = nWidth + .ColWidth(i)
        Next i
        .Width = nWidth
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        With grdOrder
            If .Rows = .FixedRows Then Exit Sub
        End With
        
        Call CheckCount
    End If
End Sub

Private Sub grdOrder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdOrder
        If .Rows = .FixedRows Or .MouseRow < .FixedRows Or .MouseRow >= .Rows Or .MouseCol <> 1 Then Exit Sub
    End With

    Call CheckCount
End Sub

Private Sub CheckCount()
    With grdOrder
        If .Cell(flexcpChecked, .Row, 1) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, 1) = flexChecked
            m_nSelected = m_nSelected + 1
        Else
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
            m_nSelected = m_nSelected - 1
        End If
    End With
    
    If cboSearch.ListIndex > 0 Then
        cmdClose.Enabled = IIf(m_nSelected > 0, True, False)
    End If
End Sub


Private Sub grdOrder_RowColChange()
    Dim sOrderID$
    
    If Not m_bSkipEvent Then
        With grdOrder
            sOrderID = Replace(.TextMatrix(.Row, 3), "-", "")
        End With
    
        If (grdOrder.Row - grdOrder.TopRow + 2) > (LIMIT_ROW1 / 2) Then
            grdColor.Move 3875, 600
        Else
            grdColor.Move 3875, 4275
        End If

        Call FillGridColor(sOrderID)
    End If
End Sub

Private Sub optOrder_Click(Index As Integer)
    If optOrder(0).Value Then '[0] 관리번호
        chkSearch(3).Caption = "Order No."
        grdOrder.ColWidth(2) = GetFlexColWidth(13)
        grdOrder.ColWidth(3) = 0
    Else '[1] Order No.
        chkSearch(3).Caption = "관리 번호"
        grdOrder.ColWidth(2) = 0
        grdOrder.ColWidth(3) = GetFlexColWidth(13)
    End If
End Sub


Private Sub txtSearch_GotFocus(Index As Integer)
    With txtSearch(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FillGridOrder()
    Dim oOrder As PlusLib2.COrder
    Dim rs As Recordset
    
    Dim i%, nNowRow%, nRowCount%
    Dim nSum(7) As Long

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    proProgress.Value = 0
    lblCount = LoadResString(304)
    
    pnlProgress.Visible = True

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon

    Set rs = oOrder.GetOrderTotal(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
                                IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
                                IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0), txtSearch(3), _
                                cboSearch.ListIndex, cboOrderClss.ListIndex, cboFiber.ListIndex)
    Set oOrder = Nothing
    
    m_bSkipEvent = True
    With grdOrder
        .Redraw = flexRDNone

        nNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        nRowCount = rs.RecordCount
        For i = 1 To nRowCount
            DoEvents
            lblCount = CStr(i) & " / " & CStr(nRowCount)
            proProgress.Value = CInt((i / nRowCount) * 100)
            
            .AddItem CStr(i) & vbTab & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!kCustom & vbTab & _
                rs!Article & vbTab & Format(MakeDate(DF_LONG, rs!DvlyDate), "MM-DD") & vbTab & rs!WorkName & vbTab & _
                rs!WorkWidth & vbTab & CStr(rs!ChunkRate) & "+" & CStr(rs!LossRate) & vbTab & rs!OrderQty & vbTab & IIf(rs!UnitClss = 0, "YD", "MT") & vbTab & _
                rs!InQty & vbTab & rs!SetQty & vbTab & _
                rs!PassQty & vbTab & rs!WeavQty & vbTab & rs!DyeQty & vbTab & rs!OutQty & vbTab & (rs!OrderQty - rs!OutQty)

            nSum(0) = nSum(0) + IIf(rs!UnitClss = 0, rs!OrderQty, Int(rs!OrderQty / 0.9144))
            nSum(1) = nSum(1) + IIf(rs!UnitClss = 0, rs!InQty, Int(rs!InQty / 0.9144))
            nSum(2) = nSum(2) + IIf(rs!UnitClss = 0, rs!SetQty, Int(rs!SetQty / 0.9144))
            nSum(3) = nSum(3) + IIf(rs!UnitClss = 0, rs!PassQty, Int(rs!PassQty / 0.9144))
            nSum(4) = nSum(4) + IIf(rs!UnitClss = 0, rs!WeavQty, Int(rs!WeavQty / 0.9144))
            nSum(5) = nSum(5) + IIf(rs!UnitClss = 0, rs!DyeQty, Int(rs!DyeQty / 0.9144))
            nSum(6) = nSum(6) + IIf(rs!UnitClss = 0, rs!OutQty, Int(rs!OutQty / 0.9144))
            nSum(7) = nSum(7) + IIf(rs!UnitClss = 0, (rs!OrderQty - rs!OutQty), Int((rs!OrderQty - rs!OutQty) / 0.9144))

            If (i Mod 2) = 0 Then
                .Row = i
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If
            If rs!CloseClss = "*" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, 1) = vbRed
            End If
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .Rows = .Rows + 1
            .RowHidden(.Rows - 1) = True
            
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = " "
'            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            
            .TextMatrix(.Rows - 1, 10) = nSum(0)
            .TextMatrix(.Rows - 1, 12) = nSum(1)
            .TextMatrix(.Rows - 1, 13) = nSum(2)
            .TextMatrix(.Rows - 1, 14) = nSum(3)
            .TextMatrix(.Rows - 1, 15) = nSum(4)
            .TextMatrix(.Rows - 1, 16) = nSum(5)
            .TextMatrix(.Rows - 1, 17) = nSum(6)
            .TextMatrix(.Rows - 1, 18) = nSum(7)
            
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
            
            .HighLight = flexHighlightAlways
            If .Rows <= nNowRow Then
                .Row = .Rows - 1
            Else
                .Row = nNowRow
            End If
            
            .Col = .FixedCols
            .ColSel = .Cols - 1
            grdColor.Visible = True
            
            Call FillGridColor(MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE))
        Else
            .HighLight = flexHighlightNever
            grdColor.Visible = False
            MsgBox LoadResString(203), vbInformation
        End If

'        Call ChangeScroll(0)
        
        
        .Redraw = flexRDDirect
    End With
    
    With grdTotal
        .TextMatrix(0, 1) = SetCurrency(nRowCount) & " 건"
        For i = 0 To 7
            .TextMatrix(0, i + 2) = nSum(i)
        Next i
    End With
    
    m_nSelected = 0
    m_bSkipEvent = False
    
    If cboSearch.ListIndex = 0 Then cmdClose.Enabled = False
    
    pnlProgress.Visible = False
    
    Screen.MousePointer = vbArrow
    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing
    
    pnlProgress.Visible = False
    Screen.MousePointer = vbArrow
    
    Call ErrorBox(Err.Number, "OrderClose.FillGridOrder", Err.Description)
End Sub


Private Sub FillGridColor(OrderID As String)
    Dim nNowRow%

    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    'If m_bSkipEvent Then Exit Sub
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    Set rs = oOrder.GetOrderSubTotal(OrderID, 0, 0)
    Set oOrder = Nothing
    
    With grdColor
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            nNowRow = .Row
            .Rows = .FixedRows
        Else
            nNowRow = 1
        End If
        
        Do Until rs.EOF
            If rs!OrderSeq <> 0 Then
                .AddItem CStr(.Rows) & vbTab & rs!OrderSeq & vbTab & _
                        rs!Color & vbTab & CheckNull(rs!DesignNO) & vbTab & rs!ColorQty & vbTab & _
                        0 & vbTab & rs!SetQty & vbTab & rs!PassQty & vbTab & rs!WeavQty & vbTab & rs!DyeQty & vbTab & _
                        rs!OutQty & vbTab & (rs!ColorQty - rs!OutQty)
            End If
            rs.MoveNext
        Loop
        rs.Close

        If .Rows < LIMIT_ROW2 Then
            If .Rows = .FixedRows Then
                .ScrollBars = flexScrollBarNone
            Else
                'S_2014012_태을염직_01 에 의한 수정-OLD소스
''                .Height = (.RowHeight(.FixedRows) + 45) * .Rows + 240
                'S_2014012_태을염직_01 에 의한 수정-NEW소스
                
                .Height = (.RowHeight(.FixedRows) + 90) * .Rows + 330
                
                .ScrollBars = flexScrollBarNone
            End If
        Else
            .Height = 2640
            .ScrollBars = flexScrollBarVertical
        End If

'        Call ChangeScroll(1)
        If .Rows > .FixedRows Then
            If .Rows > nNowRow Then
                .Row = nNowRow
            Else
                .Row = .Rows - 1
            End If
            
            .HighLight = flexHighlightAlways
            
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If
        .Redraw = flexRDDirect
    End With

    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing
    
    grdColor.Redraw = flexRDDirect
    
    Call ErrorBox(Err.Number, "Order.FillGridColor", Err.Description)
End Sub

Private Sub ChangeScroll(Index As Integer)
    On Error GoTo ErrHandler

    If Index = 0 Then '[1] 수주 내용
        With grdOrder
            If .Rows > LIMIT_ROW1 Then
                .ColWidth(4) = LIMIT_WIDTH1 - 240
            Else
                .ColWidth(4) = LIMIT_WIDTH1
            End If
        End With
    Else '[2] 색상 내용
        With grdColor
            If .Rows > LIMIT_ROW2 Then
                .ColWidth(2) = LIMIT_WIDTH2 - 240
            Else
                .ColWidth(2) = LIMIT_WIDTH2
            End If
        End With
    End If
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "OrderClose.ChangeScroll", Err.Description)
    
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
        ElseIf Index = 3 Then
            Call ReturnCode(LG_ORDER, , False, txtSearch(3))
            
            If optOrder(1).Value Then
                txtSearch(3).Text = txtSearch(3).Tag
            End If
        End If
        cmdSearch.SetFocus
    End If
End Sub

