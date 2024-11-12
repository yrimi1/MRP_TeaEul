VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSubulReport 
   Caption         =   "수불명세서"
   ClientHeight    =   9390
   ClientLeft      =   4275
   ClientTop       =   2910
   ClientWidth     =   15225
   Icon            =   "frmSubulReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   15225
   Begin Threed.SSPanel pnlPrn 
      Height          =   3225
      Left            =   5070
      TabIndex        =   25
      Top             =   2700
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   5689
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   870
         TabIndex        =   30
         Top             =   1590
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄범위"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.ComboBox cboCustom 
         Height          =   300
         Left            =   2160
         Style           =   2  '드롭다운 목록
         TabIndex        =   29
         Top             =   1590
         Width           =   2715
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   3000
         TabIndex        =   28
         Top             =   2400
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   1500
         TabIndex        =   27
         Top             =   2400
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "인쇄"
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   60
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   767
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "수불 명세서 인쇄"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   735
         Left            =   2160
         TabIndex        =   46
         Top             =   810
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1296
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optPrn 
            Caption         =   "전체현황"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   48
            Top             =   120
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optPrn 
            Caption         =   "개별인쇄"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   47
            Top             =   420
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Index           =   0
         Left            =   870
         TabIndex        =   49
         Top             =   810
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   3165
         Left            =   30
         Top             =   30
         Width           =   5805
      End
   End
   Begin Threed.SSPanel pnlSub 
      Height          =   4245
      Left            =   5520
      TabIndex        =   12
      Top             =   1980
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7488
      _Version        =   196609
      BevelWidth      =   2
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   609
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "Order내역"
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdOutWare 
         Height          =   465
         Left            =   6090
         TabIndex        =   21
         Top             =   1620
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "출고등록 화면으로"
      End
      Begin Threed.SSCommand cmdStuffIN 
         Height          =   465
         Left            =   2010
         TabIndex        =   20
         Top             =   1620
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "입고등록 화면으로"
      End
      Begin Threed.SSCommand cmdOrder 
         Height          =   435
         Left            =   6120
         TabIndex        =   19
         Top             =   90
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
         _Version        =   196609
         Caption         =   "수주등록 화면으로"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOrder 
         Height          =   1095
         Left            =   90
         TabIndex        =   13
         Top             =   510
         Width           =   8085
         _cx             =   14261
         _cy             =   1931
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
      Begin VSFlex7LCtl.VSFlexGrid grdStuffIN 
         Height          =   1575
         Left            =   90
         TabIndex        =   17
         Top             =   2100
         Width           =   4005
         _cx             =   7064
         _cy             =   2778
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
      Begin VSFlex7LCtl.VSFlexGrid grdOutWare 
         Height          =   1575
         Left            =   4170
         TabIndex        =   18
         Top             =   2100
         Width           =   4005
         _cx             =   7064
         _cy             =   2778
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1710
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "입고내역"
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   315
         Left            =   4200
         TabIndex        =   24
         Top             =   1710
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "출고내역"
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   4
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  '단색
         Height          =   3735
         Left            =   30
         Top             =   30
         Width           =   8235
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7515
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1110
      Width           =   15195
      _cx             =   26802
      _cy             =   13256
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
      RowHeightMin    =   350
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
      Begin Threed.SSPanel pnlMemo 
         Height          =   3915
         Left            =   4740
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6906
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdClose 
            Height          =   465
            Left            =   5130
            TabIndex        =   41
            Top             =   3390
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   820
            _Version        =   196609
            Caption         =   "닫기"
         End
         Begin Threed.SSCommand cmdSave 
            Height          =   465
            Left            =   3930
            TabIndex        =   40
            Top             =   3390
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   820
            _Version        =   196609
            Caption         =   "저장"
         End
         Begin VB.TextBox txtMemo 
            Height          =   2895
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   39
            Top             =   450
            Width           =   6285
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   405
            Left            =   30
            TabIndex        =   43
            Top             =   30
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   714
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   16711680
            Caption         =   "메모항목 등록"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1931
      _Version        =   196609
      Begin VB.CheckBox chkNotIncSampleOutware 
         Caption         =   "Sample 출고 제외"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3930
         TabIndex        =   54
         Top             =   120
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkKG 
         Caption         =   "KG재고 조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5580
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cboSubulWidth 
         Height          =   300
         Left            =   3930
         Style           =   2  '드롭다운 목록
         TabIndex        =   51
         Top             =   720
         Width           =   1035
      End
      Begin VB.ComboBox CboOrderFlag 
         Height          =   300
         Left            =   5040
         Style           =   2  '드롭다운 목록
         TabIndex        =   44
         Top             =   390
         Width           =   945
      End
      Begin Threed.SSCommand cmdMemo 
         Height          =   465
         Left            =   14430
         TabIndex        =   42
         Top             =   570
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "메모"
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   34
         Top             =   720
         Width           =   2235
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   825
         Left            =   6900
         TabIndex        =   31
         Top             =   210
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   1455
         _Version        =   196609
         BorderWidth     =   1
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Label Label2 
            Caption         =   "■  상세내역 에서 수주등록, 입고등록, 출고등록 하면으로 이동하여  수정 할 수 있습니다."
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   7245
         End
         Begin VB.Label Label1 
            Caption         =   "■  입출고의 상세내역을 조회: 상세내역 보임 선택 하십시오."
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   150
            Width           =   7035
         End
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   405
         Width           =   2235
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   6000
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 저장"
         Top             =   270
         Width           =   840
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "수불일자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   405
         Width           =   1185
         _ExtentX        =   2090
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
         Left            =   3570
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   405
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   2
         Left            =   2640
         TabIndex        =   11
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
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
            TabIndex        =   36
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   3570
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
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
         Left            =   3990
         TabIndex        =   45
         Top             =   390
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "사용구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11850
      TabIndex        =   8
      Tag             =   "PERM_ADDNEW"
      Top             =   8670
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13590
      TabIndex        =   9
      Top             =   8670
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   30
      TabIndex        =   14
      Top             =   8700
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
      _Version        =   196609
      Caption         =   "상세내역 "
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "보임"
      End
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   1
         Left            =   1035
         TabIndex        =   16
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "숨김"
         Value           =   -1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   510
      Index           =   1
      Left            =   6810
      TabIndex        =   50
      Top             =   8700
      Visible         =   0   'False
      Width           =   3030
      _cx             =   5345
      _cy             =   900
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   11.25
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
   Begin VSFlex7LCtl.VSFlexGrid grdDataOrder 
      Height          =   510
      Left            =   3600
      TabIndex        =   52
      Top             =   8700
      Visible         =   0   'False
      Width           =   3030
      _cx             =   5345
      _cy             =   900
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   11.25
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
Attribute VB_Name = "frmSubulReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
' 변경이력
'-----------------------------------------------------------------------------------------------------
'요청ID : S_201211_태을염직_03
'요청일자 : 2012.11.22
'요청내용 : 수불명세서 엑셀로 출력되게
'변경내용 : 엑셀 양식으로 변경-기존 그리드인쇄
'
'요청ID : S_201212_태을염직_07
'요청일자 : 2012.12.20
'요청내용 : 수불명세서 상단에 거래처 코드 안나오게,출력일자 삭제,페이지 표시, 생지입고 절수 삭제하고 오더NO 칸 넓게
'변경내용 :
'******************************************************************************************************
Option Explicit

Private m_bloading As Boolean
Dim sPrinter As String

'S_201203_태을염직_02 에 의한 추가
Private Const EXCEL_ROW As Integer = 42             '엑셀 한 페이지 총 행수(프린트 여백 내)

'S_201211_태을염직_03 에 의한 추가
Private Const REPORTFILE = "\Report\원본SubulReport.xls"            '2011.09.30, old remark , old:Private Const REPORTFILE = "\Report\SubulReport.rpt"
'태을염직 KG 수불 없음
''Private Const REPORTFILE_KG = "\Report\원본SubulReport_KG.xls"      '2011.09.30, old remark , old:Private Const REPORTFILE = "\Report\SubulReport.rpt"
Private Const REPORTFILE1 = "tmp_SubulReport.xls"                   '2011.09.30, old remark , old:Private Const REPORTFILE = "\Report\SubulReport.rpt"
Private Const EXCEL_ROLL_ROW    As Integer = 49


Private Sub cmdClose_Click()
    pnlMemo.Visible = False
End Sub

Private Sub cmdMemo_Click()
    pnlMemo.Visible = True
    Dim IOClss As String
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim vKey As Variant
    Dim RetBool As Boolean, sMemo As String

    On Error GoTo ErrHandler

    Dim sWorkID$

    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    With grdData(0)
        IOClss = Trim(.TextMatrix(.Row, 16))
        Select Case Trim(.TextMatrix(.Row, 16))
            Case "1"
                vKey = Split(.TextMatrix(.Row, 17), "-")
                    RetBool = oStuffIn.GetSubulMemo(IOClss, vKey(0), vKey(1), vKey(2), "", 0, sMemo)
            Case "2"
                vKey = Split(.TextMatrix(.Row, 17), "-")
                    RetBool = oStuffIn.GetSubulMemo(IOClss, "", "", 0, vKey(0), vKey(1), sMemo)
        End Select
    End With

    Set oStuffIn = Nothing
    
    If RetBool Then
        txtMemo = sMemo
    End If

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmSubulReport.SaveData", Err.Description)
    Set oStuffIn = Nothing
    
End Sub

Private Sub cmdOrder_Click()
    If grdOrder.Rows > grdOrder.FixedRows Then
        Call frmOrder.LoadOrder(grdOrder.TextMatrix(grdOrder.Row, 8))
    End If
End Sub

Private Sub cmdOutWare_Click()
    Dim OutWareKey As Variant
    With grdOutWare
        If .Rows > .FixedRows Then
            OutWareKey = Split(.TextMatrix(.Row, 5), "-")
            If UBound(OutWareKey) = 1 Then
                If .TextMatrix(.Row, 4) = "1" Then
                    Call frmOutwareIns.LoadOutWareIns(OutWareKey(0), OutWareKey(1))
                Else
                
                    Call frmOutware.LoadOutWare(OutWareKey(0), Int(OutWareKey(1)))
                End If
            End If
        End If
    End With

End Sub

'S_201211_태을염직_03 에 의한 OLD소스
'Private Sub cmdPrint_Click()
'    pnlPrn.Visible = True
'End Sub

'''201211_태을염직_03 에 의한 NEW
''Private Sub cmdPrint_Click()
''
''    cmdPrint.Enabled = False
''
''    If Trim(txtSearch(1).Tag) = "" Then
''        If grdData(0).FixedRows >= grdData(0).Row Or grdData(0).TextMatrix(grdData(0).Row, 19) = "" Then        '거래처 코드
''            MsgBox "수불현황은 거래처 전체로 출력할수 없습니다.", vbOKOnly, "출력 불가"
''            Exit Sub
''        End If
''
''    End If
''    Call ExcelPrintSubul(PlusMDI.PrintPreview)         '1개의 선택업체 출력
''
''
''    cmdPrint.Enabled = True
''End Sub

'S_201211_태을염직_03 에 의한 NEW소스
Private Sub cmdPrint_Click()

    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim sDate$, eDate$
    
    On Error GoTo ErrHandler
    
'    If Len(txtSearch(1).Tag) = 0 And (grdData(0).FixedRows >= grdData(0).Row Or grdData(0).TextMatrix(grdData(0).Row, 19) = "") Then
    If Len(txtSearch(1).Tag) = 0 Then
        
        MsgBox "수불 명세서는 거래처를 선택한후에 발행이 됩니다." & vbCrLf & "먼저 거래처를 선택하여주십시오.", vbOKOnly
        Exit Sub
    End If
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass
    
''    'KG수불 없음
''    If chkKG.Value = vbChecked Then
''        Call MakeExcelSubulReport_Kg
''
''    Else
        Call MakeExcelSubulReport
''    End If

    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmSubulReport.cmdPrint_Click", Err.Description)

End Sub

'S_201211_태을염직_03 에 의한 추가
Private Sub MakeExcelSubulReport()
    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oRange                          As Excel.Range
    Dim oFs                             As FileSystemObject
    Dim oCustom                         As PlusLib2.CCustom
    Dim oOutware                        As PlusLib2.COutWare
    Dim rs                              As ADODB.Recordset
    Dim lssql                           As String
    Dim lstempQty                       As String
    Dim lsTempLossQty                   As String
    Dim i%, j%, nRow%, nLimitRow%, nBaseRow%, nCurRow%, nPage%, sReport$
    Dim nOrderSeq%, sLotNo$
    Dim nColorRoll%, nColorQty#, nColorLossQty#
    Dim sWorkWidth                      As String
    Dim EXCEL_1PageData_ROW             As Integer
    Dim vColorSum()                     As Double
    Dim sUnit                           As String
    Dim nSeq                            As Integer
    Dim sDate                           As String
    Dim sArticleID                      As String
    
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    Set oExcel = New Excel.Application
    
    '원본파일 open
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & REPORTFILE)

    '//디버깅시 아래 주석 해제
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
 
    EXCEL_1PageData_ROW = 48
 
    
    With oExcel
       ' Make Sum
        .Worksheets("Form").Activate
        
      
        .Cells(2, 1) = g_companyInfo.Company_Name           '공급자
        
        'S_201212_태을염직_07 에 의한 수정
''        .Cells(3, 6) = txtSearch(1).Tag & " -▶ " & txtSearch(1)                                '거래처
        .Cells(3, 6) = txtSearch(1)                                '거래처
        .Cells(4, 6) = MakeDate(DF_FULL, dtpDate(1)) & " - " & MakeDate(DF_FULL, dtpDate(2))    '기간
        
        'S_201212_태을염직_07 에 의한 수정
''        .Cells(4, 44) = MakeDate(DF_LONG, Date)         '출력일자
          
        .Worksheets("Print").Activate
        
        nPage = 1
        nBaseRow = GetExcelRollBaseRow(nPage)
        Call InsertExcelForm(oExcel, nPage)
        nCurRow = nBaseRow + 7
            
        nOrderSeq = 0
        For i = grdData(0).FixedRows To grdData(0).Rows - 1
            If grdData(0).RowHidden(i) = True Then GoTo Next_i
            If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
                nPage = nPage + 1
                nBaseRow = GetExcelRollBaseRow(nPage)
                Call InsertExcelForm(oExcel, nPage)
                nCurRow = nBaseRow + 7
                nRow = 0
            End If
            
            
            '-----------------------------------------
            '* 품명이 달라졌으면
            '-----------------------------------------
''            If sArticleID <> grdData(0).TextMatrix(i, 3) Then                                      '품명
''                If i > grdData(0).FixedRows Then
''                End If
''
''                If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
''                    nPage = nPage + 1
''                    nBaseRow = GetExcelRollBaseRow(nPage)
''                    Call InsertExcelForm(oExcel, nPage)
''                    nCurRow = nBaseRow + 7
''                    nRow = 0
''                End If
''
''                If nRow > 0 Then
''                    Set oRange = .Worksheets("Print").Range(GF_Excel_CA(1) & (nCurRow + nRow), GF_Excel_CA(45) & (nCurRow + nRow))
''                    oRange.Borders(xlEdgeTop).LineStyle = xlContinuous
''                    oRange.Borders(xlEdgeTop).Weight = xlHairline
''                    oRange.Borders(xlEdgeTop).ColorIndex = xlAutomatic
''                End If
''                .Cells(nCurRow + nRow, 4) = grdData(0).TextMatrix(i, 3)                        '품명
''
''                sDate = ""
''            End If
''
''            '4:가공구분,7:생지입고수량,10:출고수량,
''            '36:가공지출고-소요량, 40:재고량,45:OrderNo
''            If grdData(0).TextMatrix(i, 4) = "" And grdData(0).TextMatrix(i, 7) = "" And .Cells(nCurRow + nRow, 4) = "" And .Cells(nCurRow + nRow, 10) = "" And _
''                grdData(0).TextMatrix(i, 11) = "" And grdData(0).TextMatrix(i, 13) = "" Then
''                GoTo Next_i
''            End If
            
            
''            If sDate <> grdData(0).TextMatrix(i, 2) Then                                      '일자
                .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, 2)
''             End If
              
            .Cells(nCurRow + nRow, 4) = grdData(0).TextMatrix(i, 3)                        '품명
            .Cells(nCurRow + nRow, 13) = grdData(0).TextMatrix(i, 4)                           '가공구분
            .Cells(nCurRow + nRow, 18) = grdData(0).TextMatrix(i, 5)                           '전기이월
            
            'S_201212_태을염직_07 에 의한 주석
''            .Cells(nCurRow + nRow, 22) = grdData(0).TextMatrix(i, 6)                           '생지입고-절수
            .Cells(nCurRow + nRow, 22) = grdData(0).TextMatrix(i, 7)                           '생지입고-수량
            .Cells(nCurRow + nRow, 26) = grdData(0).TextMatrix(i, 9)                           '출고 - 절수
            .Cells(nCurRow + nRow, 29) = grdData(0).TextMatrix(i, 10)                          '출고 - 출고량
            .Cells(nCurRow + nRow, 33) = grdData(0).TextMatrix(i, 11)                          '출고 -소요량
            .Cells(nCurRow + nRow, 37) = grdData(0).TextMatrix(i, 13)                          '재고량
            .Cells(nCurRow + nRow, 42) = grdData(0).TextMatrix(i, 14)                          'Order NO
 
''            If grdData(0).TextMatrix(i, 3) = "**소계**" Then
''
''                Set oRange = .Worksheets("Print").Range(GF_Excel_CA(6) & (nCurRow + nRow), GF_Excel_CA(45) & (nCurRow + nRow)) '슷자를 Excel Column 영문자로 변경 Range 설정
''                'oRange.Interior.ColorIndex = 15
''                oRange.Font.Bold = True
''
''            End If
            
            nRow = nRow + 1
Next_i:
            sDate = grdData(0).TextMatrix(i, 2)                                                '일자
            sArticleID = grdData(0).TextMatrix(i, 3)                                           '품명
        
        
        Next i
        
        If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
            nPage = nPage + 1
            nBaseRow = GetExcelRollBaseRow(nPage)
            Call InsertExcelForm(oExcel, nPage)
            nCurRow = nBaseRow + 7
            nRow = 0
        End If
     
    End With

    
    Set oFs = New FileSystemObject
    
    '수불명세서 폴더 없을 경우 생성
    If Not oFs.FolderExists(CStr(App.Path) & "\수불명세서\") Then
        oFs.CreateFolder (CStr(App.Path) & "\수불명세서\")           '없을경우 폴더 생성
    End If
    
    '저장할 파일명
''    sReport = App.Path & "\" & REPORTFILE1
    sReport = App.Path & "\수불명세서\수불명세서_" & Left(MakeDate(DF_SHORT, dtpDate(1)), 6) & "_" & txtSearch(1) & ".xls"
    
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
    
    Set oFs = Nothing

    Call oExcelBook.SaveAs(sReport)

    oExcel.WindowState = xlMaximized
    oExcel.Application.Visible = True
    oExcel.ActiveWindow.SelectedSheets.PrintPreview

    Screen.MousePointer = vbDefault

    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing
    
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
''    Resume Next
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[MakeExcelSubulReport]"
    End If
     
    
    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing

End Sub

'''KG 수불 태을염직에는 없음
''Private Sub MakeExcelSubulReport_Kg()
'''KG 이 있는 수불명세서 발행, 2011.12.03, S_201111_조일_03 에 의한 신규 생성
''    Dim oExcel                          As Excel.Application
''    Dim oExcelBook                      As Excel.Workbook
''    Dim oExcelSheet                     As Excel.Worksheet
''    Dim oRange                          As Excel.Range
''    Dim oFs                             As FileSystemObject
''    Dim oCustom                         As PlusLib2.CCustom
''    Dim oOutware                        As PlusLib2.COutWare
''    Dim rs                              As ADODB.Recordset
''    Dim lssql                           As String
''    Dim lstempQty                       As String
''    Dim lsTempLossQty                   As String
''    Dim i%, j%, nRow%, nLimitRow%, nBaseRow%, nCurRow%, nPage%, sReport$
''    Dim nOrderSeq%, sLotNo$
''    Dim nColorRoll%, nColorQty#, nColorLossQty#
''    Dim sWorkWidth                      As String
''    Dim EXCEL_1PageData_ROW             As Integer
''    Dim vColorSum()                     As Double
''    Dim sUnit                           As String
''    Dim nSeq                            As Integer
''    Dim sDate                           As String
''    Dim sArticleID                      As String
''
''
''    On Error GoTo ErrHandler
''
''    Screen.MousePointer = vbHourglass
''    Set oExcel = New Excel.Application
''
''    '원본파일 open
''    Set oExcelBook = oExcel.Workbooks.Open(App.Path & REPORTFILE_KG)
''
''    '//디버깅시 아래 주석 해제
'''    oExcel.WindowState = xlMaximized
'''    oExcel.Application.Visible = True
''
''    EXCEL_1PageData_ROW = 48
''
''
''    With oExcel
''       ' Make Sum
''        .Worksheets("Form").Activate
''
''
''        .Cells(2, 1) = g_companyInfo.Company_Name
''        .Cells(3, 5) = txtSearch(1).Tag & " -▶ " & txtSearch(1)                                '거래처
''        .Cells(4, 5) = MakeDate(DF_FULL, dtpDate(1)) & " - " & MakeDate(DF_FULL, dtpDate(2))    '기간
''
''        .Cells(4, 53) = MakeDate(DF_LONG, Date)
''
''        .Worksheets("Print").Activate
''
''        nPage = 1
''        nBaseRow = GetExcelRollBaseRow(nPage)
''        Call InsertExcelForm(oExcel, nPage)
''        nCurRow = nBaseRow + 7
''
''        nOrderSeq = 0
''        For i = grdData(0).FixedRows + 1 To grdData(0).Rows - 1 ' Step 2
''            If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
''                nPage = nPage + 1
''                nBaseRow = GetExcelRollBaseRow(nPage)
''                Call InsertExcelForm(oExcel, nPage)
''                nCurRow = nBaseRow + 7
''                nRow = 0
''            End If
''
''
''            '-----------------------------------------
''            '* 품명이 달라졌으면
''            '-----------------------------------------
''            If sArticleID <> grdData(0).TextMatrix(i, 2) Then                                      '품명
''                If i > grdData(0).FixedRows Then
''                End If
''
''                If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
''                    nPage = nPage + 1
''                    nBaseRow = GetExcelRollBaseRow(nPage)
''                    Call InsertExcelForm(oExcel, nPage)
''                    nCurRow = nBaseRow + 7
''                    nRow = 0
''                End If
''
''                If nRow > 0 Then
''                    Set oRange = .Worksheets("Print").Range(GF_Excel_CA(1) & (nCurRow + nRow), GF_Excel_CA(57) & (nCurRow + nRow)) '슷자를 Excel Column 영문자로 변경 Range 설정
''                    oRange.Borders(xlEdgeTop).LineStyle = xlContinuous
''                    oRange.Borders(xlEdgeTop).Weight = xlHairline
''                    oRange.Borders(xlEdgeTop).ColorIndex = xlAutomatic
''                End If
''                .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, 2)                        '품명
''
''                sDate = ""
''            End If
''
''            If grdData(0).TextMatrix(i, 4) = "" And grdData(0).TextMatrix(i, 6) = "" And grdData(0).TextMatrix(i, 8) = "" And .Cells(nCurRow + nRow, 9) = "" And .Cells(nCurRow + nRow, 11) = "" And _
''                grdData(0).TextMatrix(i, 12) = "" And grdData(0).TextMatrix(i, 13) = "" And grdData(0).TextMatrix(i, 15) = "" And grdData(0).TextMatrix(i, 18) = "" Then
''                GoTo Next_i
''            End If
''
''            If sDate <> grdData(0).TextMatrix(i, 3) Then                                       '일자
''                .Cells(nCurRow + nRow, 6) = grdData(0).TextMatrix(i, 3)
''             End If
''
''
''            .Cells(nCurRow + nRow, 8) = grdData(0).TextMatrix(i, 4)                            '입출고처
''            .Cells(nCurRow + nRow, 11) = grdData(0).TextMatrix(i, 6)                           '생지입고-수량
''
''            .Cells(nCurRow + nRow, 14) = grdData(0).TextMatrix(i, 8)                           '출고 - Order NO
''            .Cells(nCurRow + nRow, 19) = grdData(0).TextMatrix(i, 24)                          '출고 - 수주량  'S_201201_조일_15 에 의한 추가
''            .Cells(nCurRow + nRow, 23) = grdData(0).TextMatrix(i, 9)                           '출고 - 가공구분
''            .Cells(nCurRow + nRow, 25) = grdData(0).TextMatrix(i, 25)                          '출고 - 폭      'S_201201_조일_15 에 의한 추가
''
''            .Cells(nCurRow + nRow, 28) = grdData(0).TextMatrix(i, 11)                          '출고 - 출고량
''            .Cells(nCurRow + nRow, 32) = grdData(0).TextMatrix(i, 12)                          '출고 -소요량
''
''            .Cells(nCurRow + nRow, 36) = grdData(0).TextMatrix(i, 13)                          '출고 - KG출고량
''            .Cells(nCurRow + nRow, 40) = grdData(0).TextMatrix(i, 14)                          '출고 - KG소요량
''            .Cells(nCurRow + nRow, 44) = grdData(0).TextMatrix(i, 15)                          'Loss
''            .Cells(nCurRow + nRow, 46) = grdData(0).TextMatrix(i, 17)                          '재고량
''            .Cells(nCurRow + nRow, 50) = grdData(0).TextMatrix(i, 18)                          '재고량
''            .Cells(nCurRow + nRow, 54) = grdData(0).TextMatrix(i, 21)                          '비고
''
''
''            If grdData(0).TextMatrix(i, 3) = "**소계**" Then
''
''                Set oRange = .Worksheets("Print").Range(GF_Excel_CA(6) & (nCurRow + nRow), GF_Excel_CA(57) & (nCurRow + nRow)) '슷자를 Excel Column 영문자로 변경 Range 설정
''                'oRange.Interior.ColorIndex = 15
''                oRange.Font.Bold = True
''
''            End If
''
''            nRow = nRow + 1
''Next_i:
''            sDate = grdData(0).TextMatrix(i, 3)                                                '일자
''            sArticleID = grdData(0).TextMatrix(i, 2)                                           '품명ID
''
''
''        Next i
''
''        If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
''            nPage = nPage + 1
''            nBaseRow = GetExcelRollBaseRow(nPage)
''            Call InsertExcelForm(oExcel, nPage)
''            nCurRow = nBaseRow + 7
''            nRow = 0
''        End If
''
''    End With
''
''
''
''
''    Set oFs = New FileSystemObject
''
''    '수불명세서 폴더 없을 경우 생성
''    If Not oFs.FolderExists(CStr(App.Path) & "\수불명세서\") Then
''        oFs.CreateFolder (CStr(App.Path) & "\수불명세서\")           '없을경우 폴더 생성
''    End If
''
''    '저장할 파일명
''''    sReport = App.Path & "\" & REPORTFILE1
''    sReport = App.Path & "\수불명세서\수불명세서_kg_" & Left(MakeDate(DF_SHORT, dtpDate(1)), 6) & "_" & txtSearch(1) & ".xls"
''
''    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
''    Set oFs = Nothing
''
''    Call oExcelBook.SaveAs(sReport)
''
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
''    oExcel.ActiveWindow.SelectedSheets.PrintPreview
''
''    Screen.MousePointer = vbDefault
''
''    Set oExcelSheet = Nothing
''    Set oExcelBook = Nothing
''    Set oExcel = Nothing
''    Set oFs = Nothing
''
''    Exit Sub
''
''ErrHandler:
''    Screen.MousePointer = vbDefault
''
''    If Err.Number <> 0 Then
''        MsgBox Err.Number & "," & Err.Description, vbCritical, "[MakeExcelSubulReport_Kg]"
''    End If
''
''
''
''    Set oExcelSheet = Nothing
''    Set oExcelBook = Nothing
''    Set oExcel = Nothing
''    Set oFs = Nothing
''
''End Sub
'S_201211_태을염직_03 에 의한 추가
Private Function GetExcelRollBaseRow(nPage)
    GetExcelRollBaseRow = (nPage - 1) * EXCEL_ROLL_ROW
End Function

'S_201211_태을염직_03 에 의한 추가
Private Function InsertExcelForm(oExcel As Excel.Application, nPage As Integer)
    Dim i%, nBaseRow%

    nBaseRow = GetExcelRollBaseRow(nPage)
    With oExcel
        .Sheets("Form").Select

        .Rows("1:" & CStr(EXCEL_ROLL_ROW)).Select
        .Selection.Copy

        .Sheets("Print").Select
        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
        .Selection.Insert Shift:=xlDown
        
       'S_201212_태을염직_07 에 의한 추가-현재 페이지 표시
        .Cells(nBaseRow + 49, 42) = "PAGE : " & nPage
    End With
End Function


Private Sub cmdPrnCancel_Click()
    pnlPrn.Visible = False
End Sub

Private Sub cmdPrnOK_Click()
    Dim II%, vCustom As Variant
    
    If optPrn(0).Value = True Then
        Call FillGrdList
    Else
        If cboCustom.Text = AllStr Then
           
            For II = 1 To cboCustom.ListCount - 1
                Call SetDataToPrn(cboCustom.List(II))
                
            Next II
        Else
            Call SetDataToPrn(cboCustom.Text)
        End If
    End If
    Call ReturnPrinter(sPrinter)
    pnlPrn.Visible = False
    
End Sub

Sub FillGrdList()
    Dim i%, nRows As Integer, II As Integer, JJ As Integer
    Dim sDate As String, eDate As String
    
    sDate = MakeDate(DF_SHORT, dtpDate(1))
    eDate = MakeDate(DF_SHORT, dtpDate(2))
       
    With grdData(0)
'        .Rows = grdData(0).FixedRows
'        .Cols = grdData(0).Cols
'        .FixedRows = grdData(0).FixedRows
'        .Redraw = flexRDBuffered
'        .ExtendLastCol = False
'
'        .GridLines = flexGridInset

        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        nRows = 0
        .Cell(flexcpText, nRows, 0, nRows, .Cols - 1) = "수 불  명 세 서  현 황"
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
        .RowHeight(nRows) = 800
        
        nRows = 1
        .RowHeight(nRows) = 500
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "▶ 거 래 처 : 전거래처 "
        
        
        nRows = 2
        .RowHeight(nRows) = 500
        
        .Cell(flexcpText, nRows, 2, nRows, 2) = "태을염직(주)"
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "▶ 정산일자 : " & sDate & " ~ " & eDate
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To .FixedRows - 1
           .MergeRow(i) = True
        Next i
        
        .ColHidden(16) = True
'
'        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter
'        .Cell(flexcpBackColor, 3, 1, 4, .Cols - 1) = &HF5F5F5
'
'        .ExtendLastCol = False
'        .Redraw = flexRDDirect
'
'        .GridLinesFixed = flexGridInset
'        .ColHidden(0) = True
'        .ColHidden(1) = False
'        .ColHidden(5) = True
'        .ColHidden(9) = True
'        .ColHidden(14) = True
'        .ColHidden(15) = True
'
'        .ColWidth(1) = 1800
'        .ColWidth(2) = 2200
'        .ColWidth(3) = 1300
'        .ColWidth(8) = 1600
'
'        nRows = .Rows
'        For II = grdData(0).FixedRows To grdData(0).Rows - 1
'                .AddItem ""
'                .RowHeight(.Rows - 1) = 400
'                For JJ = 0 To .Cols - 1
'                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
'                Next JJ
'                .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
'        Next II
'
'        .MergeCells = flexMergeFree
'        .ExtendLastCol = False
        
        
        Call SetPrintMode(grdData(0), 2, True)
        
        .PrintGrid "태을염직", True, 2, 100, 500
        Call SetPrintMode(grdData(0), 2, False)
        .ColHidden(16) = False
        
'        .ColWidth(1) = 1500
'        .ColWidth(2) = 2200
'        .ColWidth(4) = 1300
'        .ColWidth(4) = 1400
        
'        .ExtendLastCol = True
    End With

End Sub

Sub SetDataToPrn(ByVal kCustom As String)
    Dim II%, JJ%, sRows As Integer

    
    Call FillGrdPrintHeader(kCustom)
    With grdData(1)
        sRows = .Rows
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            If grdData(0).TextMatrix(II, 1) = kCustom Then
                .AddItem ""
                For JJ = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
                .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
            End If
        Next II
        
        
        For II = 0 To .Cols - 1
            .Cell(flexcpAlignment, sRows, II, .Rows - 1, II) = grdData(0).ColAlignment(II)
        Next II
        
        .ColWidth(3) = 2800    ' + 500   '품명
        .ColWidth(4) = 1200    ' + 400   '가공구분
        .ColWidth(5) = 1200    ' + 400   '가공구분
        .ColWidth(7) = 1000    ' + 300   '입고수량
        .ColWidth(10) = 1000   ' + 300   '입고수량
        .ColWidth(11) = 1000   ' + 300   '입고수량
        .ColWidth(13) = 1000   ' + 300   '입고수량
        .ColWidth(14) = 2100
        
        
        .ColWidth(1) = 0
        .ColWidth(6) = 0
        .ColWidth(9) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        
        .ExtendLastCol = False
        
        Call SetPrintMode(grdData(1), 2, True)
        
        For II = .FixedRows To .Rows - 1
            If .TextMatrix(II, 17) = "3" Then
                .Cell(flexcpFontBold, .Rows - 1, 2, .Rows - 1, .Cols - 2) = True
                .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 2) = PRNHeaderColor
            End If
        Next II
        
        .PrintGrid "태을염직", True, 2, 700, 500
        
        Call SetPrintMode(grdData(1), 2, False)
        
    End With
End Sub

Private Function SaveData() As Boolean
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim vKey As Variant

    On Error GoTo ErrHandler

    SaveData = False

    Dim sWorkID$

    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    With grdData(0)
        Select Case Trim(.TextMatrix(.Row, 16))
            Case "1"
                vKey = Split(.TextMatrix(.Row, 17), "-")
                SaveData = oStuffIn.UpdateStuffINMemo(vKey(0), vKey(1), vKey(2), Trim(txtMemo))
                                
            Case "2"
                vKey = Split(.TextMatrix(.Row, 17), "-")
                SaveData = oStuffIn.UpdateOutWareMemo(vKey(0), vKey(1), Trim(txtMemo))
        End Select
        
    End With

    Set oStuffIn = Nothing

    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, "frmSubulReport.SaveData", Err.Description)
    Set oStuffIn = Nothing

End Function

Private Sub cmdSave_Click()
    If SaveData Then
        MsgBox "입력한 내용이 저장 되었습니다.", vbInformation
        Call FillGridData
    End If
End Sub

Private Sub cmdStuffIN_Click()
    With grdStuffIN
        If .Rows > .FixedRows Then
            frmStuffIN.ZOrder
  '          frmStuffIN.WindowState = vbNormal
            
            Call frmStuffIN.LoadStuffIN(.TextMatrix(.Row, 4))
        End If
    End With
    
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 15360, 9840
    
    pnlSub.Move 6540, 4620, 8265, 3765

    Call SetOperate(Me)
    Call ChangeMode(Me, True)
    
    dtpDate(1) = DateSerial(Year(Date), Month(Date), 1)
    dtpDate(2) = Date
    Call InitGrid(0)
    Call InitGrid(1)
    

    
    Call InitGridSub
    pnlSub.Visible = False
    pnlPrn.Visible = False
    
    With CboOrderFlag
        .AddItem "9.전체"
        .AddItem "1.사용"
        .AddItem "0.비사용"
        .ListIndex = 0
    End With
    
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    txtSearch(1).Enabled = False
    cmdFind(1).Enabled = False
    
    txtSearch(2).Enabled = False
    cmdFind(2).Enabled = False
    pnlMemo.Visible = False
    Call SetStuffWidth(cboSubulWidth)
    cboSubulWidth.ListIndex = 0
    cboSubulWidth.Enabled = False
'    cmdprnint.Picture = LoadResPicture("CHECK", vbResIcon)
    
End Sub

Sub FillGrdPrintHeader(ByVal kCustom As String)
    Dim i%, nRows As Integer
    
    With grdData(1)
        .Rows = grdData(0).FixedRows
        .FixedRows = grdData(0).FixedRows
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridInset
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHidden(4) = False
        
        
        nRows = 0
        .RowHeight(nRows) = 500
        .FontSize = 10
        
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "가 공 품  수 불  내 역 서"
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRows, 0, nRows, .Cols - 1) = True
        
        .RowHeight(nRows) = 800
        
        
        nRows = 1
        .RowHeight(nRows) = 500
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "▶ 거 래 처 : " & kCustom
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 11
        
        
        nRows = 2
        .RowHeight(nRows) = 500
        
        .Cell(flexcpText, nRows, 3, nRows, .Cols - 1) = "▶ 정산일자 : " & MakeDate(DF_FULL, dtpDate(1)) & " ~ " & MakeDate(DF_FULL, dtpDate(2))
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 11
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter
        '.Cell(flexcpBackColor, 3, 2, 4, .Cols - 1) = &HE0E0E0
        .Cell(flexcpFontBold, 3, 2, 4, .Cols - 1) = True
        
        .ColWidth(2) = grdData(0).ColWidth(2) + 500
        .ColWidth(14) = 800
        .ColWidth(15) = 0
        
        .RowHeight(3) = 450
        .RowHeight(4) = 450
        .ExtendLastCol = True
        
        .MergeRow(.Rows - 2) = True
        .MergeCells = flexMergeFree
        
        '--- 실제 데이터 부분과 Merge 분리하기 위해 빈라인 하나 넣음
        .AddItem ""
        .RowHidden(.Rows - 1) = True
        .SheetBorder = vbBlack
        
        .GridLines = flexGridInset

        .ExtendLastCol = False
        .Redraw = flexRDDirect
    End With
    
End Sub


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index
        Case 1
            txtSearch(1).Enabled = chkSearch(Index).Value
            cmdFind(1).Enabled = chkSearch(Index).Value
            If chkSearch(Index).Value Then
                txtSearch(1).SetFocus
            End If
        Case 2
            txtSearch(2).Enabled = chkSearch(2).Value
            cboSubulWidth.Enabled = chkSearch(2).Value
            
            cmdFind(2).Enabled = chkSearch(2).Value
            If chkSearch(2).Value Then
                txtSearch(2).SetFocus
            End If
    End Select
    
    
'    If chkSearch(Index) Then
'        If Index = 1 Or Index = 2 Then
'            cmdFind(Index).Enabled = True
'        End If
'        txtSearch(Index).Enabled = True
'        txtSearch(Index).SetFocus
'    Else
'        If Index = 1 Or Index = 2 Then
'            cmdFind(Index).Enabled = False
'        End If
'        txtSearch(Index).Enabled = False
'        cmdSearch.SetFocus
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub




Sub FillGrdOrder()
    Dim Key_Var As Variant
    Dim IOClss As String
    Dim StuffDate As String, StuffClss As String, StuffSeq As Integer
    Dim OrderID As String, OutSeq As Integer
    
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
    Dim i%, sOrderID$, bFlag As Boolean, II%

    Screen.MousePointer = vbHourglass

   ' On Error GoTo ErrHandler

    pnlSub.Visible = True
    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    IOClss = ""
    StuffDate = "": StuffClss = "": StuffSeq = 0
    OrderID = "": OutSeq = 0
    With grdData(0)
        Key_Var = Split(.TextMatrix(.Row, 17), "-")
        IOClss = .TextMatrix(.Row, 16)
    End With
    If IOClss = "1" Then    ' 입고
        StuffDate = Key_Var(0)
        StuffClss = Key_Var(1)
        StuffSeq = Key_Var(2)
    Else
        OrderID = Key_Var(0)
        OutSeq = Key_Var(1)
    End If

    Set rs = oSubul.GetSubulOrderSub(IOClss, StuffDate, StuffClss, StuffSeq, OrderID, OutSeq)
    Set oSubul = Nothing
    
    With grdStuffIN
        .Rows = .FixedRows
    End With
    
    With grdOutWare
        .Rows = .FixedRows
    End With

    With grdOrder
        .Rows = .FixedRows
        Do Until rs.EOF
            .AddItem "" & vbTab & rs!OrderNo & vbTab & SetCurrency(rs!OrderQty, 0) & vbTab & SetCurrency(rs!OutQty, 0) & vbTab & SetCurrency(rs!OrderQty - rs!OutQty) & vbTab & _
                   SetCurrency(rs!StuffQty, 0) & vbTab & SetCurrency(rs!OutRealQty, 0) & vbTab & 0 & vbTab & rs!OrderID
            rs.MoveNext
        Loop
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            Call grdOrder_RowColChange
        End If
    End With
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub grdData_RowColChange(Index As Integer)
    If Index = 0 Then

        If optView(0).Value = True Then
            Call grdDataSelect(Index)
        Else
            pnlSub.Visible = False
    
        End If
    End If
End Sub

Private Sub grdDataSelect(ByVal Index As Integer)
    With grdData(Index)
        If .TextMatrix(.Row, 16) = "1" Or .TextMatrix(.Row, 16) = "2" Then
            If .TextMatrix(.Row, 0) < (.TopRow + 9) Then
                pnlSub.Top = 4700
            Else
                pnlSub.Top = 900
            End If
            
            Call FillGrdOrder
        Else
            pnlSub.Visible = False
            
        End If
    End With

End Sub

Private Sub grdOrder_RowColChange()

    With grdData(0)
        If .TextMatrix(.Row, 16) = "1" Or .TextMatrix(.Row, 16) = "2" Then
            Call FillGrdSubulSub
        End If
    End With

End Sub

Sub FillGrdSubulSub()
    Dim OrderID  As String
    
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset

    Screen.MousePointer = vbHourglass


    With grdOrder
        OrderID = .TextMatrix(.Row, 8)
    End With

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon

    With grdStuffIN
        .Rows = .FixedRows
    End With
    
    With grdOutWare
        .Rows = .FixedRows
    End With
    
    Set rs = oSubul.GetsubulsReportSub(OrderID, MakeDate(DF_SHORT, dtpDate(1)), MakeDate(DF_SHORT, dtpDate(2)))
    Set oSubul = Nothing
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
            If rs!IOCls = "1" Then
                With grdStuffIN
                    .AddItem "" & vbTab & MakeDate(DF_LONG, rs!IODate) & vbTab & rs!Roll & vbTab & SetCurrency(rs!Qty, 0) & vbTab & rs!Pkey
                End With
            Else
                With grdOutWare
                    .AddItem "" & vbTab & MakeDate(DF_LONG, rs!IODate) & vbTab & rs!Roll & vbTab & SetCurrency(rs!Qty, 0) & vbTab & rs!OutType & vbTab & rs!Pkey
                End With
                
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    
    With grdStuffIN
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways

            .Select .FixedRows, 0, .FixedRows, .Cols - 1
        End If
    End With
    
    
    With grdOutWare
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0, .FixedRows, .Cols - 1
        End If
    End With
    
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub optView_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
        pnlSub.Visible = True
    Else
        pnlSub.Visible = False
    End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    End If
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGridSub()
    Dim i%, nRows%

    'Order내역나타내는 Grid
    Call SetVSFlexGrid(grdOrder)
    With grdOrder
        .Rows = 1
        .Cols = 9
        
        .FixedRows = 1
        .FixedCols = 1
        
        .RowHeight(0) = 350

        nRows = 0
        .TextMatrix(nRows, 0) = " "
        .TextMatrix(nRows, 1) = "OrderNO":             .ColWidth(1) = 3000:     .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(nRows, 2) = "수주량":              .ColWidth(2) = 1500:     .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(nRows, 3) = "출고량":              .ColWidth(3) = 1500:     .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(nRows, 4) = "Over량":              .ColWidth(4) = 1500:     .ColAlignment(4) = flexAlignRightCenter
        
        .TextMatrix(nRows, 5) = "입고량":              .ColWidth(5) = 1000:     .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(nRows, 6) = "소요량":              .ColWidth(6) = 1000:     .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(nRows, 7) = "재고량":              .ColWidth(7) = 1000:     .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(nRows, 8) = "OrderID":             .ColWidth(8) = 1000:     .ColAlignment(8) = flexAlignLeftCenter
        
        .ExplorerBar = flexExNone
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
'
        .ColHidden(5) = True
        .ColHidden(6) = True
        .ColHidden(7) = True
        .ColHidden(8) = True
        
    End With

    '입고내역나타내는 Grid
    Call SetVSFlexGrid(grdStuffIN)
    With grdStuffIN
        .Rows = 1
        .Cols = 5
        
        .FixedRows = 1
        .FixedCols = 1
        
        .RowHeight(0) = 350

        nRows = 0
        .TextMatrix(nRows, 0) = " "
        .TextMatrix(nRows, 1) = "일자":              .ColWidth(1) = 1000:     .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "절수":              .ColWidth(2) = 800:     .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(nRows, 3) = "수량":              .ColWidth(3) = 1300:     .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(nRows, 4) = "pkey":              .ColWidth(4) = 0:     .ColAlignment(4) = flexAlignLeftCenter
        
        .ExplorerBar = flexExNone
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarVertical
        .Redraw = flexRDDirect
    End With

    '출고내역나타내는 Grid
    Call SetVSFlexGrid(grdOutWare)
    With grdOutWare
        .Rows = 1
        .Cols = 6
        
        .FixedRows = 1
        .FixedCols = 1
        
        .RowHeight(0) = 350

        nRows = 0
        .TextMatrix(nRows, 0) = " "
        .TextMatrix(nRows, 1) = "일자":              .ColWidth(1) = 1000:     .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "절수":              .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(nRows, 3) = "수량":              .ColWidth(3) = 1300:     .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(nRows, 4) = "OutType":           .ColWidth(4) = 0:        .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(nRows, 5) = "pkey":              .ColWidth(5) = 0:        .ColAlignment(5) = flexAlignLeftCenter
        
        .ExplorerBar = flexExNone
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarVertical
        .Redraw = flexRDDirect
    End With

End Sub
Private Sub InitGrid(ByVal Index As Integer)
    Dim i%, nRows%

    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Rows = 5
        .Cols = 20      'S_201211_태을염직_03 에 의한 수정(OLD:19)
        
        .FixedRows = 5
        .FixedCols = 1
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        nRows = 3
        .TextMatrix(nRows, 0) = " "
        .TextMatrix(nRows, 1) = "거래처"
        .TextMatrix(nRows, 2) = "일자"
        .TextMatrix(nRows, 3) = "품    명"
        .TextMatrix(nRows, 4) = "가공구분"
        .TextMatrix(nRows, 5) = "전월이월"
        
        .TextMatrix(nRows, 6) = "생지입고"
        .TextMatrix(nRows, 7) = "생지입고"
        .TextMatrix(nRows, 8) = ""
        .TextMatrix(nRows, 9) = "가공지 출고"
        .TextMatrix(nRows, 10) = "가공지 출고"
        .TextMatrix(nRows, 11) = "가공지 출고"
        .TextMatrix(nRows, 12) = ""
        .TextMatrix(nRows, 13) = "재고량"
        .TextMatrix(nRows, 14) = "OrderNO"
        
        .TextMatrix(nRows, 15) = "비고"
        .TextMatrix(nRows, 16) = "M"
        .TextMatrix(nRows, 17) = "Cls"
        .TextMatrix(nRows, 18) = "pkey"
        .TextMatrix(nRows, 19) = "거래처코드"          'S_201211_태을염직_03 에 의한 추가
        
        
        
        nRows = 4
        
        .TextMatrix(nRows, 0) = " ":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "거래처":            .ColWidth(1) = 1500:     .ColAlignment(1) = flexAlignLeftCenter:       .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "일자":              .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:     .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "품    명":          .ColWidth(3) = 2800:     .ColAlignment(3) = flexAlignLeftCenter:       .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "가공구분":          .ColWidth(4) = 1100:     .ColAlignment(4) = flexAlignLeftCenter:       .FixedAlignment(4) = flexAlignCenterCenter
        .TextMatrix(nRows, 5) = "전월이월":          .ColWidth(5) = 900:      .ColAlignment(5) = flexAlignRightCenter:      .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "절수":              .ColWidth(6) = 900:      .ColAlignment(6) = flexAlignRightCenter:      .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "수량":              .ColWidth(7) = 900:      .ColAlignment(7) = flexAlignRightCenter:      .FixedAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = "":                  .ColWidth(8) = 0:        .ColAlignment(8) = flexAlignCenterCenter:     .FixedAlignment(8) = flexAlignCenterCenter
        .TextMatrix(nRows, 9) = "절수":              .ColWidth(9) = 900:      .ColAlignment(9) = flexAlignRightCenter:      .FixedAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "출고량":           .ColWidth(10) = 900:     .ColAlignment(10) = flexAlignRightCenter:     .FixedAlignment(10) = flexAlignCenterCenter
        .TextMatrix(nRows, 11) = "소요량":           .ColWidth(11) = 1000:    .ColAlignment(11) = flexAlignRightCenter:     .FixedAlignment(11) = flexAlignCenterCenter
        .TextMatrix(nRows, 12) = "":                 .ColWidth(12) = 0:       .ColAlignment(12) = flexAlignRightCenter:     .FixedAlignment(12) = flexAlignCenterCenter
        .TextMatrix(nRows, 13) = "재고량":           .ColWidth(13) = 1000:    .ColAlignment(13) = flexAlignRightCenter:     .FixedAlignment(13) = flexAlignCenterCenter
        .TextMatrix(nRows, 14) = "OrderNO":          .ColWidth(14) = 1400:    .ColAlignment(14) = flexAlignLeftCenter:      .FixedAlignment(14) = flexAlignCenterCenter
        .TextMatrix(nRows, 15) = "비고":             .ColWidth(15) = 500:    .ColAlignment(15) = flexAlignLeftCenter:      .FixedAlignment(15) = flexAlignCenterCenter
        .TextMatrix(nRows, 16) = "M":                .ColWidth(16) = 300:     .ColAlignment(16) = flexAlignCenterCenter:    .FixedAlignment(16) = flexAlignCenterCenter
        .TextMatrix(nRows, 17) = "Cls":              .ColWidth(17) = 0
        .TextMatrix(nRows, 18) = "pkey":             .ColWidth(18) = 0
        .TextMatrix(nRows, 19) = "거래처코드":             .ColWidth(19) = 0
        .Cell(flexcpFontBold, 3, 0, 4, .Cols - 1) = True
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(3) = True
        
        
        For i = 0 To .FixedRows - 3
            .RowHidden(i) = True
        Next i
        
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i
        
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
        
    End With

End Sub

Private Sub FillGridData()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
    Dim i%, sOrderID$, bFlag As Boolean, II%, SubulWidthID As String, dWorkName As String

    Dim lsAdditemStr As String
        
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    
    SubulWidthID = Format(cboSubulWidth.ItemData(cboSubulWidth.ListIndex), "0#")


    Set rs = oSubul.GetSubulReport(MakeDate(DF_SHORT, dtpDate(1)), MakeDate(DF_SHORT, dtpDate(2)) _
                        , IIf(chkSearch(1), 1, 0), txtSearch(1).Tag _
                        , IIf(chkSearch(2), 1, 0), txtSearch(2).Tag, SubulWidthID _
                        , Left(CboOrderFlag, 1))
    Set oSubul = Nothing
    cboCustom.Clear
    cboCustom.AddItem AllStr
    With grdData(0)
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            
            If Trim(rs!kCustom) <> Trim(dCustom_str) Then
'                .AddItem ""
'                .RowHidden(.Rows - 1) = True
                cboCustom.AddItem Trim(rs!kCustom)
                
            End If
'
'            ElseIf dDate_str <> rs!IODate Then
'                .AddItem CStr(.Rows - 1) & vbTab & Trim(rs!kCustom)
'                .RowHidden(.Rows - 1) = True
'                ElseIf dWorkName = Trim(rs!WorkName) Then
'                    .AddItem CStr(.Rows - 1) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!WorkName)
'                    .RowHidden(.Rows - 1) = True
'
'            End If


''            'S_201211_태을염직_03 에 의한 의한 수정-OLD소스
''            .AddItem CStr(i) & vbTab & Trim(rs!kCustom) & vbTab & MakeDate(DF_MD, rs!IODate) & vbTab & Trim(rs!Article) & "“" & vbTab & _
''                     Trim(rs!WorkName) & vbTab & IIf(rs!BeforeQty = 0, "", SetCurrency(rs!BeforeQty)) & vbTab & IIf(rs!StuffRoll = 0, "", rs!StuffRoll) & vbTab & _
''                     IIf(rs!StuffQty = 0, "", SetCurrency(rs!StuffQty, 0)) & vbTab & "" & vbTab & _
''                     IIf(rs!OutRoll = 0, "", rs!OutRoll) & vbTab & _
''                     IIf(rs!OutQty = 0, " ", SetCurrency(rs!OutQty, 0)) & IIf(rs!UnitClss = "M", " M", Space(2)) & vbTab & _
''                     IIf(rs!OutRealQty = 0, "", SetCurrency(rs!OutRealQty, 0)) & vbTab & "" & vbTab & IIf(rs!StockQty = 0, "", SetCurrency(rs!StockQty)) & vbTab & _
''                     Trim(rs!OrderNo) & vbTab & Trim(rs!Remark) & vbTab & IIf(rs!Memo = 1, "▼", "") & vbTab & rs!Cls & vbTab & rs!Pkey
                     
            
            'S_201211_태을염직_03 에 의한 의한 수정-NEW소스
            lsAdditemStr = CStr(i)                                                                      '0)순서
            lsAdditemStr = lsAdditemStr & vbTab & Trim(rs!kCustom)                                      '1)거래처명
            lsAdditemStr = lsAdditemStr & vbTab & MakeDate(DF_MD, rs!IODate)                            '2)일자
            lsAdditemStr = lsAdditemStr & vbTab & Trim(rs!Article)                                      '3)품명
            lsAdditemStr = lsAdditemStr & vbTab & Trim(rs!WorkName)                                     '4)가공구분
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!BeforeQty = 0, "", SetCurrency(rs!BeforeQty))  '5)전월이월
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!StuffRoll = 0, "", rs!StuffRoll)               '6)생지입고-절수
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!StuffQty = 0, "", SetCurrency(rs!StuffQty, 0)) '7)생기입고-수량
            lsAdditemStr = lsAdditemStr & vbTab & ""                                                    '8)공백
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!OutRoll = 0, "", rs!OutRoll)                   '9)가공지출고-절수
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!OutQty = 0, " ", SetCurrency(rs!OutQty, 0)) & IIf(rs!UnitClss = "M", " M", Space(2))  '10)가공지출고-수량
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!OutRealQty = 0, "", SetCurrency(rs!OutRealQty, 0))  '11)가공지출고-소요량
            lsAdditemStr = lsAdditemStr & vbTab & ""                                                    '12)공백
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!StockQty = 0, "", SetCurrency(rs!StockQty))    '13)재고량
            lsAdditemStr = lsAdditemStr & vbTab & Trim(rs!OrderNo)                                      '14)OrderNO
            lsAdditemStr = lsAdditemStr & vbTab & Trim(rs!Remark)                                       '15)비고
            lsAdditemStr = lsAdditemStr & vbTab & IIf(rs!Memo = 1, "▼", "")                            '16)M
            lsAdditemStr = lsAdditemStr & vbTab & rs!Cls                                                '17)Cls
            lsAdditemStr = lsAdditemStr & vbTab & rs!Pkey                                               '18)pkey
            lsAdditemStr = lsAdditemStr & vbTab & Trim(rs!CustomID)                                                    '19)거래처코드
            
            .AddItem lsAdditemStr
            
            
            dCustom_str = Trim(rs!kCustom)
            dDate_str = Trim(rs!IODate)
            dWorkName = Trim(rs!WorkName)

            Select Case rs!Cls
                Case "0"
                    .TextMatrix(.Rows - 1, 2) = "이월"
'                    .TextMatrix(.Rows - 1, 13) = SetCurrency(rs!BeforeQty)
'                    If (Left(CboOrderFlag.Text, 1) = "2" Or Left(CboOrderFlag.Text, 1) = "3") Then
'                        .RowHidden(.Rows - 1) = True
'                    End If
                Case "3"
                    .TextMatrix(.Rows - 1, 2) = "소계"
                    If Trim(.TextMatrix(.Rows - 1, 13)) = "" Then
                        .TextMatrix(.Rows - 1, 13) = 0
                    End If
                    
'                    .TextMatrix(.Rows - 1, 3) = ""
                    .Cell(flexcpFontBold, .Rows - 1, 2, .Rows - 1, .Cols - 2) = True
                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 2) = PRNHeaderColor
'                    If rs!nAllCnt < 2 Then
'                        If rs!nBeforeCnt < 1 Then
'                            .RowHidden(.Rows - 1) = True
'                        End If
'                    End If
            End Select
            
            .AddItem "" & vbTab & Trim(rs!kCustom)
            .RowHidden(.Rows - 1) = True
            
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        .MergeCells = flexMergeRestrictColumns

        .MergeCol(1) = True

'
'
'        For II = 0 To 1
'            .MergeCol(II) = True
'        Next II
        
        .Redraw = flexRDDirect
        .SetFocus
        cboCustom.ListIndex = 0
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmSubulReport.FillGridData", Err.Description)
End Sub


