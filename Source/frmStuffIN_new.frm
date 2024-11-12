VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffIN_new 
   Caption         =   "생지 입고 관리"
   ClientHeight    =   9255
   ClientLeft      =   1860
   ClientTop       =   2625
   ClientWidth     =   15180
   Icon            =   "frmStuffIN_new.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15180
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   720
      Left            =   14310
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   50
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   810
   End
   Begin VB.ComboBox CboStuffClss2 
      Height          =   300
      Left            =   11940
      Style           =   2  '드롭다운 목록
      TabIndex        =   49
      Top             =   30
      Width           =   1965
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전일"
      Height          =   315
      Index           =   0
      Left            =   60
      MousePointer    =   99  '사용자 정의
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   60
      MousePointer    =   99  '사용자 정의
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   30
      Width           =   615
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   4890
      TabIndex        =   46
      Top             =   30
      Width           =   1935
   End
   Begin VB.TextBox txtArticle 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4890
      TabIndex        =   45
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   8760
      TabIndex        =   44
      Top             =   60
      Width           =   1695
   End
   Begin VB.ComboBox cboOrderID 
      Height          =   300
      Left            =   11940
      Style           =   2  '드롭다운 목록
      TabIndex        =   43
      Top             =   390
      Width           =   1965
   End
   Begin VB.CommandButton cmdShrink 
      Caption         =   "확장"
      Height          =   345
      Index           =   0
      Left            =   3180
      TabIndex        =   42
      Top             =   780
      Width           =   765
   End
   Begin VB.CommandButton cmdShink 
      Caption         =   "축소"
      Height          =   345
      Index           =   1
      Left            =   3960
      TabIndex        =   41
      Top             =   780
      Width           =   705
   End
   Begin VB.OptionButton optGroup 
      Caption         =   "오더별"
      Height          =   330
      Index           =   0
      Left            =   30
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   780
      Value           =   -1  'True
      Width           =   990
   End
   Begin VB.OptionButton optGroup 
      Caption         =   "거래처별"
      Height          =   330
      Index           =   1
      Left            =   1080
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   780
      Width           =   990
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13500
      TabIndex        =   13
      Top             =   8520
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdStuffIN 
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   8580
      Visible         =   0   'False
      Width           =   1800
      _cx             =   3175
      _cy             =   661
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
   Begin VSFlex7LCtl.VSFlexGrid grdGroup 
      Height          =   7020
      Left            =   0
      TabIndex        =   17
      Top             =   1110
      Width           =   9210
      _cx             =   16245
      _cy             =   12382
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
   Begin Threed.SSPanel pnlData 
      Height          =   7680
      Left            =   9240
      TabIndex        =   18
      Top             =   780
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   13547
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   795
         Index           =   3
         Left            =   1860
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   74
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   795
         Index           =   0
         Left            =   3450
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   73
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   795
         Index           =   2
         Left            =   5040
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   72
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   795
         Index           =   1
         Left            =   4245
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   71
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   795
         Index           =   4
         Left            =   2655
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   70
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         IMEMode         =   10  '한글 
         Index           =   0
         Left            =   1470
         TabIndex        =   5
         Top             =   3030
         Width           =   2340
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1470
         Style           =   2  '드롭다운 목록
         TabIndex        =   8
         Top             =   4095
         Width           =   1350
      End
      Begin VB.TextBox txtRemark 
         Height          =   1350
         Left            =   1470
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   12
         Top             =   6270
         Width           =   4335
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   10
         Top             =   5595
         Width           =   1170
      End
      Begin VB.ComboBox CboStuffClss 
         Height          =   300
         Left            =   1470
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   1260
         Width           =   2175
      End
      Begin VB.TextBox txtStuffSeq 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1590
         Width           =   1005
      End
      Begin VB.TextBox txtCustomID 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   1470
         TabIndex        =   4
         Top             =   2685
         Width           =   2355
      End
      Begin VB.TextBox txtThreadName 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   1470
         TabIndex        =   6
         Top             =   3405
         Width           =   2340
      End
      Begin VB.TextBox TxtArticleID2 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   1470
         TabIndex        =   7
         Top             =   3750
         Width           =   2340
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   14
         Left            =   1470
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   4440
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "다른색상입력"
         Height          =   345
         Left            =   3930
         TabIndex        =   19
         Top             =   1170
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtOrderID 
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   1950
         Width           =   2340
      End
      Begin VB.TextBox txtOrderNO 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   2310
         Width           =   2340
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   11
         Top             =   5940
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   2
         Left            =   1470
         TabIndex        =   0
         Top             =   915
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyy-MM-dd (ddd)"
         Format          =   23724035
         CurrentDate     =   37068
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   10
         Left            =   75
         TabIndex        =   21
         Top             =   1260
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "입고 구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   75
         TabIndex        =   22
         Top             =   6270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "비고 사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdStuffINSub 
         Height          =   585
         Left            =   3960
         TabIndex        =   23
         Top             =   1290
         Visible         =   0   'False
         Width           =   1425
         _cx             =   2514
         _cy             =   1032
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   60
         TabIndex        =   24
         Top             =   4095
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "입고 단위"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2685
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   26
         Top             =   5595
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "전체입고절수"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   75
         TabIndex        =   27
         Top             =   5940
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "전체입고수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   7
         Left            =   75
         TabIndex        =   28
         Top             =   1590
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "입고 순번"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   8
         Left            =   75
         TabIndex        =   29
         Top             =   915
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "입고 일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   13
         Left            =   60
         TabIndex        =   30
         Top             =   3405
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "사종"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   4
         Left            =   3840
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3780
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   14
         Left            =   60
         TabIndex        =   32
         Top             =   3750
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "품명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   12
         Left            =   60
         TabIndex        =   33
         Top             =   4440
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
         Caption         =   "가공 구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   780
         Left            =   60
         TabIndex        =   34
         Top             =   4770
         Width           =   5745
         _cx             =   10134
         _cy             =   1376
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
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   16777215
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
         Height          =   315
         Index           =   15
         Left            =   60
         TabIndex        =   35
         Top             =   2310
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "OrderNO"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   36
         Top             =   3030
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "입고처 명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   12
         Left            =   60
         TabIndex        =   37
         Top             =   2685
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "거래처"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   38
         Top             =   1950
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "관리번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   3
         Left            =   3840
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2340
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   450
         Left            =   60
         TabIndex        =   75
         Top             =   405
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   794
         _Version        =   196609
         BackColor       =   12648447
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   330
      Left            =   0
      TabIndex        =   40
      Top             =   8130
      Width           =   9210
      _cx             =   16245
      _cy             =   582
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   10650
      TabIndex        =   51
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입고구분"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   52
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   315
      Index           =   0
      Left            =   6840
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   30
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   9
      Left            =   3570
      TabIndex        =   54
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   55
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   11
      Left            =   3570
      TabIndex        =   56
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   57
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   315
      Index           =   2
      Left            =   6840
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   390
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1980
      TabIndex        =   59
      Top             =   30
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
      Left            =   1980
      TabIndex        =   60
      Top             =   390
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23724033
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   690
      TabIndex        =   61
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입고 일자"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   62
         Top             =   60
         Value           =   1  '확인
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   345
      Left            =   6420
      TabIndex        =   63
      Top             =   780
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   609
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   3
         Left            =   1380
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   90
         Width           =   1140
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   16
      Left            =   7470
      TabIndex        =   66
      Top             =   60
      Width           =   1260
      _ExtentX        =   2223
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
         Index           =   0
         Left            =   60
         TabIndex        =   67
         Top             =   60
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   315
      Left            =   10650
      TabIndex        =   68
      Top             =   390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "확정구분"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   69
         Top             =   60
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmStuffIN_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_iFlag As String * 1
Dim m_bGroupClss As Boolean     '거래처별, 오더별 Grid 구분

Private Const LIMIT_WIDTH1 = 1640
Private Const LIMIT_WIDTH2 = 2100
Private Const LIMIT_WIDTH3 = 560
Private Const LIMIT_WIDTH4 = 2000
Private Const LIMIT_ROW1 = 11
Private Const LIMIT_ROW2 = 28
Private Const LIMIT_ROW3 = 9
Private m_bSortForward As Boolean
Private m_StuffDate As String, m_StuffClss As String, m_StuffSeq As Integer


Private Sub cboName_KeyPress(Index As Integer, KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub



Private Sub CboStuffClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cmdShrink_Click(Index As Integer)
    Dim II As Integer
    Dim nRows As String, sRows_var As Variant
    
    nRows = ""
    With grdGroup
        Select Case Index
            Case 0
                For II = .FixedRows To .Rows - 1
                    If .IsCollapsed(II) = flexOutlineCollapsed Then
                        nRows = nRows & "," & II
                    End If
                Next II
            Case 1
                For II = .Rows - 1 To .FixedRows Step -1
                    If .IsCollapsed(II) = flexOutlineExpanded And .IsSubtotal(II) Then
                        nRows = nRows & "," & II
                    End If
                Next II
        End Select
    End With
    
    nRows = Mid(nRows, 2)

    sRows_var = Split(nRows, ",")

    For II = 0 To UBound(sRows_var)
        Call GridCollapse(grdGroup, val(sRows_var(II)))
    Next II

End Sub

Private Sub Command1_Click()
    With grdStuffINSub
        .Rows = .Rows + 1
        .Select .Rows - 1, 1
    End With
    grdStuffINSub.SetFocus
End Sub

Private Sub dtpDate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

Sub SetToggle()
    Dim Index As Integer
    If optOrder(2).Value Then
        Index = 2
    Else
        Index = 3
    End If
    
    Select Case Index
        Case 2
            With grdGroup
                If m_bGroupClss Then
                    .ColHidden(2) = True
                    .ColHidden(3) = False
                    
                Else
                    .ColHidden(4) = True
                    .ColHidden(5) = False
                End If
            End With
 
        Case 3
            With grdGroup
                If m_bGroupClss Then
                    .ColHidden(2) = False
                    .ColHidden(3) = True
                Else
                    .ColHidden(4) = False
                    .ColHidden(5) = True
                End If
            End With
    End Select

End Sub

Private Sub GridCollapse(oFlex As VSFlexGrid, Row As Integer)
    With oFlex
        If Row < .FixedRows Then Exit Sub

        If .IsCollapsed(Row) = flexOutlineCollapsed Then
            .IsCollapsed(Row) = flexOutlineExpanded
        Else
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim i%
    
    PlusMDI.pnlMenu.Visible = False
    
    Me.Move 0, 0, 15300, 9660

    Call InitGrid
    Call InitGroup
    
    Call SetOperate(Me)
    
    '----- 입고단위 설정
    With cboUnit
        .AddItem "YDS"
        .AddItem "MTS"
        .ListIndex = 0
    End With
    
    '----- 입고구분 설정
    With CboStuffClss
        .AddItem "1.생지"
        .ItemData(0) = 1
        .AddItem "3.반품(생지)"
        .ItemData(1) = 3
    End With
    
    '----- 검색용 입고구분 설정
    With CboStuffClss2
        .AddItem "1.생지"
        .ItemData(0) = 1
        .AddItem "3.반품 생지"
        .ItemData(1) = 3
        .ListIndex = 0
    End With
    
    '----- 확정구분
    With cboOrderID
        .AddItem "수주확정"
        .AddItem "수주미확정"
        .ListIndex = 0
    End With
    
    Call MakeCodeCombo(cboName(14), CD_WORK)        ' 가공 구분
    
    '---- 날짜 설정
    For i = 0 To 2
        dtpDate(i) = Now
    Next i
    
    '--- find 컨트롤 icon설정
    For i = 0 To cmdFind.Count - 1
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    Next i

    cmdFind(1).Visible = True
    cmdFind(3).Visible = True
    
    
    
    cmdFind(0).Enabled = False
    cmdFind(2).Enabled = False
    txtArticle.Enabled = False
  '  txtOrderID.Enabled = True

    m_iFlag = ID_ADDNEW
    
    Call ClearText(txtNum, "0")
    
    
    '--- 필수입력 항목에 표시하기  거래처명
    pnlCaption(1).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(2).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(7).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(8).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(10).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(12).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(14).Picture = LoadResPicture("BASIC", vbResIcon)
    
    CboStuffClss.ListIndex = 0
    
    '---- 오더별 데이터 나타내기
    m_bGroupClss = True
    Call FillGridGroup(m_bGroupClss)
    Call optOrder_Click(2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Call SaveSetting(LoadResString(100), Me.Name, "Order", IIf(chkSearch(0) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Custom", IIf(chkSearch(1) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Article", IIf(chkSearch(2) = vbChecked, "1", "0"))
End Sub

'************************************************************
' 입고시 절수별 수량 입력하는 Grid Clear시킴
'************************************************************
Private Sub ClearGridSub()
    Dim i%, j%
    With grdStuffINSub
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .Rows = .FixedRows + 1
    End With
End Sub

'************************************************************
' 등록시 필수 입력 항목 확인
'************************************************************

Private Function CheckData() As Boolean
    CheckData = True
    
    If Trim(TxtArticleID2.Text) = "" Or TxtArticleID2.Tag = "" Then
        MsgBox "품명을 반드시 입력 하십시오.", vbInformation
        CheckData = False
        Exit Function
    End If
    
    If txtCustomID.Tag = "" Or Trim(txtCustomID.Text) = "" Then
        MsgBox "거래처을 반드시 입력 하십시오.", vbInformation
        CheckData = False
        Exit Function
    End If
    
    If val(txtNum(0)) = 0 Or val(txtNum(1)) = 0 Then
        MsgBox "입고 절수(또는 수량)이 없습니다. (절수)수량을 입력하십시오.", vbInformation
        CheckData = False
        Exit Function
    End If
End Function


Private Sub SetNewData(SetStuffINData As PlusLib2.TStuffIN)
    Dim sJobFlag As String, StuffClss As String
    
    On Error GoTo ErrHandler
    
    Select Case m_iFlag
        Case ID_ADDNEW
            sJobFlag = "I"
        Case ID_UPDATE
            sJobFlag = "U"
    End Select
    
    StuffClss = CboStuffClss.ItemData(CboStuffClss.ListIndex)
    With SetStuffINData
        .sJobFlag = sJobFlag
        .sStuffDate = MakeDate(DF_SHORT, dtpDate(2))     '[2] 입고 일자
        .sStuffClss = StuffClss                          '입고구분
        .nStuffSeq = val(txtStuffSeq)                    '순번
        .sCustomID = txtCustomID.Tag                     '발주처코드
        .sCustom = Trim(txtCustom(0))                    '원단입고처명
        .nTotRoll = CheckNum(txtNum(0))                       '원단절수
        .nTotQty = CheckNum(txtNum(1))                        '원단수량
        .sRemark = Trim(txtRemark.Text)                  '비고
        .sThreadName = Trim(txtThreadName.Text)          '사종
        .sUnitClss = cboUnit.ListIndex                   '입고단위
        .sOrderID = Trim(txtOrderID.Text)                 '관리번호
        .sWorkID = Format(cboName(14).ItemData(cboName(14).ListIndex), "0000")       ' 가공 구분
        .sArticleID = TxtArticleID2.Tag                  'item
        .sOrderNo = txtOrderNO                           'OrderNo
    End With

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Description, "frmStuffIN.SetNewData", Err.Description)

End Sub


Private Sub SetNewDataSub(SetSubData() As PlusLib2.TStuffINSub, nSeq As Integer)
    Dim iLoop%, nCount%
    Dim i%, j%, k%
    Dim nRollQty%  '--- 1절당 수량
    Dim nQty%      '--- 전체수량 절수 * 수량
    
    Dim nStuffQty() As Integer      '수량을 넣는 배열
    Dim nStuColor() As String       'Color를 넣는 배열
    
    If nSeq >= 0 Then
        ReDim nStuffQty(nSeq)
        ReDim nStuColor(nSeq)
    End If
    
    
    With grdStuffINSub
        If .Rows = .FixedRows Then Exit Sub
        
        For i = 1 To .Rows - 1
            
            For j = 2 To .Cols - 1
                If Len(.TextMatrix(i, j)) <> 0 Then
                    
                    '--- 절수 및 수량 가져오기
                    Call GetRollQty(.TextMatrix(i, j), nCount, nRollQty, nQty)
                    
                    If nCount > 1 Then
                        For k = 0 To nCount - 1
                            nStuffQty(iLoop) = nRollQty
                            nStuColor(iLoop) = .TextMatrix(i, 1)
                            iLoop = iLoop + 1
                        Next k
                    Else
                        '--- 색상명 배열에 넣기
                        nStuColor(iLoop) = .TextMatrix(i, 1)
                        nStuffQty(iLoop) = nRollQty
                        iLoop = iLoop + 1
                    End If
                End If
            Next j
        Next i
    End With
    
    ' 입고세부 구조체 배열에 데이터 넣기
    For iLoop = 0 To nSeq
        With SetSubData(iLoop)
            .sColor = nStuColor(iLoop)         '칼라명
            .sColorID = IIf(Len(nStuColor(iLoop)) > 0, GetColorID(nStuColor(iLoop)), "")        '칼라ID
            .nQty = nStuffQty(iLoop)            '수량
        End With
    Next iLoop
    
    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.SetNewDataSub", Err.Description)

End Sub


Function GetColorID(ByVal ColorName As String) As String
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    dSql_str = "SELECT ColorID FROM mt_color " & vbCr & _
               " WHERE Color = '" & ColorName & "' "
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
        GetColorID = Trim$(dRS(0))
    End If
    dRS.Close
    Set dRS = Nothing
                       
End Function
Private Sub CalcQty()
    Dim i%, j%
    Dim nRoll%, nQty%, nRollQty%       '합계
    Dim nTotRoll%, nTotQty%
    
    ' Grid Text 값 계산
    With grdStuffINSub
        If .Rows = .FixedRows Then Exit Sub
        
        For i = 1 To .Rows - 1
            For j = 2 To .Cols - 1
                If Len(.TextMatrix(i, j)) <> 0 Then
                    Call GetRollQty(.TextMatrix(i, j), nRoll, nRollQty, nQty)
                    nTotRoll = nTotRoll + nRoll
                    nTotQty = nTotQty + nQty
                End If
            Next j
        Next i
    End With

    txtNum(0) = nTotRoll
    txtNum(1) = nTotQty
End Sub

Private Function DeleteData() As Boolean
    Dim oStuffIn As PlusLib2.CStuffIN

    On Error GoTo ErrHandler
    
'    Call FillGrid
   
    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    
    DeleteData = oStuffIn.DeleteStuffIN(m_StuffDate, m_StuffClss, m_StuffSeq)
    txtOrderID.Tag = ""
    
    Set oStuffIn = Nothing
    Exit Function
    
ErrHandler:
    DeleteData = False
    Set oStuffIn = Nothing
    Call ErrorBox(Err.Number, "frmStuffIn.DeleteData", Err.Description)
End Function

Private Function SaveData() As Boolean
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim NewStuffIN   As PlusLib2.TStuffIN
    Dim StuffINSub() As PlusLib2.TStuffINSub
    Dim nSeq  As Integer
    Dim nStuffSeq As Integer
    
    On Error GoTo ErrHandler
    
    SaveData = False
    
    Dim sWorkID$
    
    '절수 개수 만큼 배열의 크기를 만든다.
    ' nSeq = CheckSub
    
    nSeq = val(txtNum(0))
    
    ' cStuffIn 클래스의 구조체에 값 대입
    Call SetNewData(NewStuffIN)
    
    '--- Stuffinsub 구조체에 값 넣기
''    If nSeq > 0 Then
''        ReDim StuffINSub(nSeq - 1)
''        Call SetNewDataSub(StuffINSub, nSeq - 1)
''    End If
    
    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    SaveData = oStuffIn.AddNewStuffIN(NewStuffIN, StuffINSub, nStuffSeq)
    txtStuffSeq = nStuffSeq
    
    Set oStuffIn = Nothing
    
    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.SaveData", Err.Description)
    Set oStuffIn = Nothing

End Function
Sub SetKeyEdit(ByVal dEdit As Boolean)
    dtpDate(2).Enabled = dEdit
    CboStuffClss.Enabled = dEdit
    txtStuffSeq.Enabled = dEdit
End Sub
'-----------------------------------------------------------------------
' StuffIN , StuffINSub Record 나타내기 및 StuffINSub Grid에 나타내기
' 수정 모드에서 사용함
'
'-------------- 입고데이터 나타내기  --- 수정모드
'-----------------------------------------------------------------------
Private Sub FillGridStuffSub(ByVal StuffDate As String, ByVal StuffClss As String, ByVal StuffSeq As Integer)
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim iRow%, i%, j%, iCount%
    
    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    ''''' StuffIN 1개의 record을 읽어온다.
    Set rs = oStuffIn.GetStuffINOne(StuffDate, StuffClss, StuffSeq)
    
    Set oStuffIn = Nothing
    
    dtpDate(2) = MakeDate(DF_LONG, StuffDate)
    CboStuffClss.ListIndex = StuffClss - 1
    txtStuffSeq.Text = StuffSeq
    Call SetKeyEdit(False)
    
    With rsData
        txtCustomID.Tag = Trim$(rs!CustomID)
        txtCustomID.Text = Trim$(rs!kCustom)
        txtCustom(0).Text = Trim$(rs!Custom)
        txtThreadName.Text = Trim$(rs!ThreadName)
        TxtArticleID2.Text = Trim$(rs!Article)
        TxtArticleID2.Tag = Trim$(rs!ArticleID)
        txtNum(0).Text = Trim$(rs!TotRoll)
        txtNum(1).Text = Format$(rs!TotQty, "###,##0")
        cboUnit = Trim$(rs!UnitName)
        cboName(14).ListIndex = FindItem(cboName(14), rs!WorkName)
        txtOrderID.Text = Trim$(rs!OrderID)
        txtRemark.Text = Trim$(rs!Remark)
    End With
    
    rs.Close
    Set rs = Nothing
    
    '--- Order 에 대한 내용 grid에 나타내기
    Call FillStuffOrderData(txtOrderID.Text)

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.FillGridStuffSub", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Sub


Private Sub chkSearch_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
       Case 0    '관리번호
            If chkSearch(0) Then
                txtSearch(3).Enabled = True
                txtSearch(3).SetFocus
            Else
                txtSearch(3).Enabled = False
                txtSearch(3).Text = ""
            End If
        Case 1    '거래처
            If chkSearch(1) = vbChecked Then
                txtCustom(1).Enabled = True
                txtCustom(1).SetFocus
                cmdFind(0).Enabled = True
            Else
                txtCustom(1).Enabled = False
                cmdFind(0).Enabled = False
                txtCustom(1).Tag = ""
            End If
        Case 2    '품명
            If chkSearch(2) = vbChecked Then
                txtArticle.Enabled = True
                txtArticle.SetFocus
                cmdFind(2).Enabled = True
            Else
                txtArticle.Enabled = False
                txtArticle.Tag = ""
                cmdSearch.SetFocus
                cmdFind(2).Enabled = False
            End If
        Case 3     '입고일자 Term
            If chkSearch(3) = vbChecked Then
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            Else
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            End If
        Case 4     '입고구분
            If chkSearch(Index) = vbChecked Then
                CboStuffClss2.Enabled = True
            Else
                CboStuffClss2.Enabled = False
            End If
        Case 5     '확정구분
            If chkSearch(5) = vbChecked Then
                cboOrderID.Enabled = True
            Else
                cboOrderID.Enabled = False
            End If
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    ' Set Order Grid
'''    Call SetVSFlexGrid(grdStuffIN)
'''    With grdStuffIN
'''        .Redraw = False
'''        .Cols = 6
'''
'''        .TextArray(0) = "완료":         .ColWidth(0) = 450
'''        .TextArray(1) = "관리번호":     .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter
'''        .TextArray(2) = "Order No.":    .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignLeftCenter
'''        .TextArray(3) = "거래처명":     .ColWidth(3) = 1720:    .ColAlignment(3) = flexAlignLeftCenter
'''        .TextArray(4) = "입고" & vbCrLf & "절수":     .ColWidth(4) = 500:    .ColAlignment(4) = flexAlignRightCenter
'''        .TextArray(5) = "입고량":       .ColWidth(5) = 900:    .ColAlignment(5) = flexAlignRightCenter
'''
'''        .ColHidden(0) = True
'''        .ColHidden(2) = True
'''        .ColAlignment(1) = flexAlignCenterCenter
'''
'''        .WordWrap = False
'''        .ScrollBars = flexScrollBarBoth
'''        .Redraw = True
'''    End With
    
    
    With grdData
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 14
        Call SetVSFlexGrid(grdData)
        .FixedCols = 0
        
        .TextArray(0) = "관리번호":         .ColWidth(0) = 0:                   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "Order NO":         .ColWidth(1) = 1300:                .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "접수일자":         .ColWidth(2) = 1050:                .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "거래처":           .ColWidth(3) = LIMIT_WIDTH1:        .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "품명":             .ColWidth(4) = 1700:                .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "구분":             .ColWidth(5) = 1300:                 .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "원단폭":           .ColWidth(6) = 660:                 .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "축율" & vbCrLf & "LOSS":             .ColWidth(7) = 1000:                 .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상수":           .ColWidth(8) = 800:                 .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "주문량":          .ColWidth(9) = 1050:               .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "입고량":          .ColWidth(10) = 1050:               .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "주문" & vbCrLf & "단위": .ColWidth(11) = 600:         .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "입고절수":        .ColWidth(12) = 0:                  .ColAlignment(12) = flexAlignCenterCenter
        
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        .ColHidden(8) = True
        
        .Redraw = flexRDDirect
    End With
    
    ' 절수, 수량 입력 Grid
    Dim sComBoList$
    
    '--- 사용가능한 ColorID, Color명 가져와서 Grid의 combobox의 item으로 설정
    Dim rs As ADODB.Recordset
    
    Set rs = GetColor
    sComBoList$ = " " & "|"
    Do Until rs.EOF
        sComBoList$ = sComBoList$ & rs(0) & "|"
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
  
''    With grdStuffINSub
''        .Redraw = flexRDNone
''        .FixedRows = 1
''        .FixedCols = 1
''    Call SetVSFlexGrid(grdStuffINSub)
''        .Cols = 12
''        .Rows = .FixedRows + 1
''        .TextArray(0) = ""
''        .ColWidth(0) = LIMIT_WIDTH3
''
''        '----------------------------------------------------------------------------------------------
''        '------ 진호염직의 경우 Color관리를 하지 않기 때문에 2번째 Color Combo Col의 width를 0으로 했음.
''        '------ 만약 다른업체에서 Color관리를 한다면 width = 1300으로 설정 함.
''        '.ColComboList(1) = sComBoList
''        '.ColWidth(1) = 1300
''        '----------------------------------------------------------------------------------------------
''        .ColWidth(1) = 0
''
''        For i = 2 To .Cols - 1
''            .TextArray(i) = i - 1
''            .ColWidth(i) = Int((grdStuffINSub.Width - .ColWidth(0) - .ColWidth(1)) / 10)
''        Next i
''
''        .Editable = flexEDKbdMouse
''        .FocusRect = flexFocusHeavy
''        .Redraw = flexRDBuffered
''
''        .ScrollBars = flexScrollBarBoth
''        .Redraw = flexRDDirect
''
''    End With

End Sub

Private Sub ChangeScroll(Index As Integer)
    Select Case Index
    Case 0
        With grdData
            If .Rows > LIMIT_ROW1 Then
                .ColWidth(4) = LIMIT_WIDTH1 - 240
            Else
                .ColWidth(4) = LIMIT_WIDTH1
            End If
        End With
    Case 1
        With grdGroup
            If m_bGroupClss Then
                If .Rows > LIMIT_ROW2 Then
                    .ColWidth(5) = LIMIT_WIDTH2 - 240
                Else
                    .ColWidth(5) = LIMIT_WIDTH2
                End If
            Else
                If .Rows > LIMIT_ROW2 Then
                    .ColWidth(7) = LIMIT_WIDTH4 - 240
                Else
                    .ColWidth(7) = LIMIT_WIDTH4
                End If
            
            End If
        End With
    Case 3
        With grdStuffINSub
            If .Rows > LIMIT_ROW3 Then
                .ColWidth(0) = LIMIT_WIDTH3 - 240
            Else
                .ColWidth(0) = LIMIT_WIDTH3
            End If
        End With
    End Select
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0                '[1] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustom(1))
        Case 1                '[2] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustomID)
        Case 2                '[3] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Case 3                '[4] 오더 코드
            Call ReturnCode(LG_ORDER, , False, txtOrderNO)
            If Trim(txtOrderNO.Tag) = "" Then
                txtOrderNO.Text = ""
                txtOrderID.Text = ""
                txtCustomID.Text = ""
                TxtArticleID2.Text = ""
            Else
                txtOrderID.Text = txtOrderNO.Tag
                If FillStuffOrderData(txtOrderID) Then
                    txtCustomID.Enabled = False
                    TxtArticleID2.Enabled = False
                Else
                    txtCustomID.Enabled = True
                    TxtArticleID2.Enabled = True
                End If
            End If
        Case 4                '[4] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, TxtArticleID2)
    End Select
End Sub


Private Sub cmdOperate_Click(Index As Integer)
    Dim sStuffKey As String
    
    Select Case Index
        Case ID_ADDNEW
            m_iFlag = ID_ADDNEW
            
            pnlMsg.Visible = True
            pnlMsg.Caption = "자료입력(추가)중..."
            
            Call SetClearEdit
            Call SetKeyEdit(True)
            Call ChangeMode(Me, False)
            
            cmdFind(1).Enabled = True
            cmdFind(3).Enabled = True
            cmdFind(4).Enabled = True
            
            txtCustomID.Enabled = True
            TxtArticleID2.Enabled = True
            
            grdData.Rows = grdData.FixedRows
            txtOrderID.SetFocus
        Case ID_UPDATE

            If val(txtStuffSeq) > 0 Then
                Call ChangeMode(Me, False)
                m_iFlag = ID_UPDATE
                pnlMsg.Visible = True
                pnlMsg.Caption = "자료입력(수정)중..."
            End If
        '-------------------------------------------------------------------------------------'
        Case ID_SAVE
            If CheckData = False Then Exit Sub
        
            If SaveData Then
                MsgBox "입력한 내용이 저장 되었습니다.", vbInformation
                Call SetClearEdit
                Call cmdSearch_Click
            End If
            m_iFlag = -1
            Call ChangeMode(Me, True)
            Call cmdShrink_Click(0)
        '-------------------------------------------------------------------------------------'
        Case ID_DELETE
            If val(txtStuffSeq) > 0 Then
                m_iFlag = ID_DELETE
                If StuffInDelete Then
                    If optGroup(0) Then
                        Call FillGridGroup
                        Call optGroup_Click(0)
                    Else
                        Call FillGridGroup(False)
                        Call optGroup_Click(1)
                    End If
                End If
                
                m_iFlag = -1
                Call SetClearEdit
                Call ClearGridSub
                Call ChangeMode(Me, True)
                grdData.Rows = grdData.FixedRows
                grdData.HighLight = flexHighlightNever
                Call cmdShrink_Click(0)
            End If
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL
            m_iFlag = -1
            pnlMsg.Visible = False
            pnlMsg.Caption = ""
            Call SetClearEdit
            Call ChangeMode(Me, True)
            Call cmdShrink_Click(0)
    End Select

End Sub
'------- StuffIN Data삭제
Function StuffInDelete() As Boolean
''    Dim sStuffKey As String
''
''    On Error GoTo ErrHandler
''
''    StuffInDelete = True
''
''    With grdGroup
''        '일자+구분+일련번호( StuffIN Key가져오기 )
''        sStuffKey = .TextMatrix(.Row, .Cols - 1)
''        Call MakeStuffKey(sStuffKey, m_StuffDate, m_StuffClss, m_StuffSeq)
''
''
''        Call FillGridStuffSub(m_StuffDate, m_StuffClss, m_StuffSeq)
        
        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
            m_iFlag = ID_DELETE
            StuffInDelete = DeleteData
        End If
''    End With
    Exit Function
ErrHandler:
    StuffInDelete = False
End Function


Private Sub SetClearEdit()
    
    Call ClearScreen(Me, "pnlData")
    
    cmdFind(1).Enabled = False
    cmdFind(3).Enabled = False
    cmdFind(4).Enabled = False
    
    txtCustomID.Tag = ""
    TxtArticleID2.Tag = ""
    
    dtpDate(2) = Now
    grdData.Rows = grdData.FixedRows
    grdData.HighLight = flexHighlightNever
    
    Call SetKeyEdit(True)
End Sub

Private Sub cmdSearch_Click()
    Dim Index As Integer
    
'    Select Case tabForm.Tab
'    Case 0
        If optGroup(0) Then
            Call FillGridGroup(True)
            Call optOrder_Click(2)
        Else
            Call InitGroup(False)
            Call FillGridGroup(False)
        End If
'    Case 1
'        Call FillGrdStuffIN
'    End Select
'    Call optOrder_Click(2)
End Sub
Sub FillGrdStuffIN()
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim iTop(2) As Integer
    Dim i%

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    Set rs = oStuffIn.GetStuffINByCustom(IIf(chkSearch(3) = vbChecked, 1, 0) _
                                , MakeDate(DF_SHORT, dtpDate(0)) _
                                , MakeDate(DF_SHORT, dtpDate(1)) _
                                , IIf(chkSearch(1) = vbChecked, 1, 0) _
                                , txtCustom(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0) _
                                , txtArticle.Tag _
                                , IIf(chkSearch(4) = vbChecked, 1, 0) _
                                , Left(CboStuffClss2, 1))

    Set oStuffIn = Nothing
    
    
    With grdStuffIN
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        If rs.RecordCount = 0 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
            Do Until rs.EOF
                .AddItem "" & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & Trim(rs!OrderNO) & vbTab & _
                         Trim(rs!kCustom) & vbTab & SetCurrency(CheckNum(rs!StuffRoll)) & vbTab & SetCurrency(CheckNum(rs!StuffQty))
                i = i + 1
                
                If (i Mod 2) = 0 Then
                    .Row = .FixedRows + i - 1
                    .Col = .FixedCols
                    .ColSel = .Cols - 1
                    .CellBackColor = COLOR_GRIDROW
                End If
                rs.MoveNext
            Loop
            .Row = .FixedRows
        End If
        .Redraw = flexRDDirect
    End With
    
    Call ChangeScroll(0)
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.FillGrdStuffIN", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing

End Sub
Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[3] 금일
        dtpDate(0) = Date - 1
        dtpDate(1) = Date - 1
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub SetGridGroup(NewFlex As VSFlexGrid)
    With NewFlex
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .BackColorBkg = vbWhite
        .SheetBorder = vbWhite
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 450
        .RowHeightMin = 275
    End With
End Sub

Private Sub InitGroup(Optional NewValue As Boolean = True)
    Dim i%
    Call SetGridGroup(grdGroup)
    
    For i = 0 To grdGroup.Cols - 1
        grdGroup.ColHidden(i) = False
    Next i
    grdGroup.Redraw = flexRDNone
    
    '----- 오더별 집계 조회
    If NewValue Then
        With grdGroup
            .Redraw = flexRDNone
            
            .Rows = 2
            .FixedRows = 2
            .FixedCols = 0
            .Cols = 18
            .RowHeight(0) = 350
            .RowHeight(1) = 350

            .TextArray(0) = " ":                                .ColWidth(0) = 200
            .TextArray(1) = " ":                                .ColWidth(1) = 200
            .TextArray(2) = "관리번호":                         .ColAlignment(2) = flexAlignCenterCenter
            .TextArray(3) = "Order NO":                         .ColAlignment(3) = flexAlignLeftCenter
            .TextArray(4) = "접수일자":                         .ColAlignment(4) = flexAlignCenterCenter
            .TextArray(5) = "거  래  처":                       .ColAlignment(5) = flexAlignLeftCenter
            .TextArray(6) = "품      명":                       .ColAlignment(6) = flexAlignLeftCenter
            .TextArray(7) = "가공구분":                         .ColAlignment(7) = flexAlignCenterCenter
            .TextArray(8) = "원단폭":                           .ColAlignment(8) = flexAlignCenterCenter
            .TextArray(9) = "축율" & vbCrLf & "LOSS":           .ColAlignment(9) = flexAlignCenterCenter
            .TextArray(10) = "색상수":                          .ColAlignment(10) = flexAlignRightCenter
            .TextArray(11) = "주문량":                          .ColAlignment(11) = flexAlignRightCenter
            .TextArray(12) = "입고" & vbCrLf & "절수":          .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "입고량":                          .ColAlignment(13) = flexAlignRightCenter
            .TextArray(14) = "배색량":                          .ColAlignment(14) = flexAlignRightCenter
            .TextArray(15) = "Sort OrderID":                    .ColAlignment(15) = flexAlignCenterCenter
            .TextArray(16) = "OrderNo":                         .ColAlignment(17) = flexAlignCenterCenter
            .TextArray(17) = "StuffDate":                       .ColAlignment(16) = flexAlignCenterCenter

            .TextArray(.Cols + 0) = " "
            .TextArray(.Cols + 1) = "  "
            .TextArray(.Cols + 2) = "관리번호"
            .TextArray(.Cols + 3) = "Order NO"
            .TextArray(.Cols + 4) = "접수일자"
            .TextArray(.Cols + 5) = "입  고  처"
            .TextArray(.Cols + 6) = "입고일자(사종)"
            .TextArray(.Cols + 7) = "가공구분"
            .TextArray(.Cols + 8) = "원단폭"
            .TextArray(.Cols + 9) = "축율" & vbCrLf & "LOSS"
            .TextArray(.Cols + 10) = "색상수"
            .TextArray(.Cols + 11) = "주문량"
            .TextArray(.Cols + 12) = "입고" & vbCrLf & "절수"
            .TextArray(.Cols + 13) = "입고량"
            .TextArray(.Cols + 14) = "배색량"
            .TextArray(.Cols + 15) = "Sort OrderID"
            .TextArray(.Cols + 16) = "OrderNo"
            .TextArray(.Cols + 17) = "StuffDate"

            .ColWidth(0) = 200
            .ColWidth(1) = 200
            .ColWidth(2) = 1400
            .ColWidth(3) = 1400
            .ColWidth(4) = 800
            .ColWidth(5) = 1400
            .ColWidth(6) = 2400
            .ColWidth(7) = 1000
            .ColWidth(8) = 800
            .ColWidth(9) = 1300
            .ColWidth(10) = 600
            .ColWidth(11) = 1000
            .ColWidth(12) = 600
            .ColWidth(13) = 2000
            .ColWidth(14) = 1400
            .ColWidth(15) = 0
            .ColWidth(16) = 0
            .ColWidth(17) = 0
            
            .ColHidden(7) = True
            .ColHidden(8) = True
            .ColHidden(9) = True
            .ColHidden(10) = True
            .ColHidden(14) = True
            
            .ColHidden(2) = True
            .ColHidden(15) = True
            .ColHidden(16) = True
            .ColHidden(17) = True
            
            For i = 1 To .Cols - 1
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next i
            .ScrollBars = flexScrollBarVertical
            .Redraw = flexRDDirect
        End With

    Else
        With grdGroup
            .Rows = 2
            .FixedRows = 2
            .FixedCols = 0
            .Cols = 19
            .RowHeight(0) = 350
            .RowHeight(1) = 350
            
            .Redraw = flexRDNone
            
    
            .TextArray(0) = "":                                       .ColWidth(0) = 100
            .TextArray(1) = "":                                       .ColWidth(1) = 200
            .TextArray(2) = "거래처ID":                               .ColWidth(2) = 600:              .ColAlignment(2) = flexAlignCenterCenter
            .TextArray(3) = "거래처명":                               .ColWidth(3) = 1400:             .ColAlignment(3) = flexAlignCenterCenter
            .TextArray(4) = "관리번호":                               .ColWidth(4) = 1400:             .ColAlignment(4) = flexAlignCenterCenter
            .TextArray(5) = "Order NO":                               .ColWidth(5) = 1400:             .ColAlignment(5) = flexAlignLeftCenter
            .TextArray(6) = "접 수 일":                               .ColWidth(6) = 1200:              .ColAlignment(6) = flexAlignCenterCenter
            .TextArray(7) = "품    명":                               .ColWidth(7) = 2400:             .ColAlignment(7) = flexAlignLeftCenter
            .TextArray(8) = "가공구분":                               .ColWidth(8) = 1000:              .ColAlignment(8) = flexAlignRightCenter
            .TextArray(9) = "원단폭":                                 .ColWidth(9) = 1000:              .ColAlignment(9) = flexAlignCenterCenter
            .TextArray(10) = "축율" & vbCrLf & "LOSS":                .ColWidth(10) = 1000:            .ColAlignment(10) = flexAlignCenterCenter
            .TextArray(11) = "색상수":                                .ColWidth(11) = 800:             .ColAlignment(11) = flexAlignRightCenter
            .TextArray(12) = "주문량":                                .ColWidth(12) = 800:             .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "입고" & vbCrLf & "절수":                .ColWidth(13) = 600:             .ColAlignment(13) = flexAlignRightCenter
            .TextArray(14) = "입고량":                                .ColWidth(14) = 2000:             .ColAlignment(14) = flexAlignRightCenter
            .TextArray(15) = "배색량":                                .ColWidth(15) = 1400:             .ColAlignment(15) = flexAlignRightCenter:
            .TextArray(16) = "CustomID":                              .ColWidth(16) = 0:               .ColAlignment(16) = flexAlignCenterCenter
            .TextArray(17) = "OrderID":                               .ColWidth(17) = 0:               .ColAlignment(17) = flexAlignCenterCenter
            .TextArray(18) = "Stuff-Pkey":                            .ColWidth(18) = 0:               .ColAlignment(18) = flexAlignCenterCenter
    
    
            .TextArray(.Cols + 0) = ""
            .TextArray(.Cols + 1) = ""
            .TextArray(.Cols + 2) = "거래처ID"
            .TextArray(.Cols + 3) = "거래처명"
            .TextArray(.Cols + 4) = "관리번호"
            .TextArray(.Cols + 5) = "Order NO"
            .TextArray(.Cols + 6) = "입 고 일"
            .TextArray(.Cols + 7) = "입 고 처"
            .TextArray(.Cols + 8) = "가공구분"
            .TextArray(.Cols + 9) = "원단폭"
            .TextArray(.Cols + 10) = "축율" & vbCrLf & "LOSS"
            .TextArray(.Cols + 11) = "색상수"
            .TextArray(.Cols + 12) = "주문량"
            .TextArray(.Cols + 13) = "입고" & vbCrLf & "절수"
            .TextArray(.Cols + 14) = "입고량"
            .TextArray(.Cols + 15) = "배색량"
            .TextArray(.Cols + 16) = "CustomID"
            .TextArray(.Cols + 17) = "OrderID"
            .TextArray(.Cols + 18) = "Stuff-Pkey"
    
            For i = 1 To .Cols - 1
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next i
    
            .ColHidden(8) = True
            .ColHidden(9) = True
            .ColHidden(10) = True
            .ColHidden(11) = True
            .ColHidden(15) = True
            .ColHidden(2) = True
            .ColHidden(4) = True
            .ColHidden(16) = True
            .ColHidden(17) = True
            .ColHidden(18) = True
            .ScrollBars = flexScrollBarVertical
    
            .Redraw = flexRDDirect
        End With
    End If
    
    With grdGroup
        .MergeCells = flexMergeFixedOnly
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i
        .Redraw = flexRDDirect
    End With
    Call SetToggle
    
    
    '--- Total값 넣기
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 7
        .ExtendLastCol = True
        
        .RowHeight(0) = 275
        
        .TextArray(0) = "합계":            .ColWidth(0) = 2000: .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "주문량:":         .ColWidth(1) = 1000: .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "0 YDS":           .ColWidth(2) = 1250: .ColAlignment(2) = flexAlignRightCenter
        
        .TextArray(3) = "입고절수:":       .ColWidth(3) = 1000: .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "0 절":            .ColWidth(4) = 1250: .ColAlignment(4) = flexAlignRightCenter
        
        .TextArray(5) = "입고수량:":       .ColWidth(5) = 1000: .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "0 YDS":           .ColWidth(6) = 1250: .ColAlignment(6) = flexAlignRightCenter
        
        For i = 1 To 5 Step 2
            .Cell(flexcpForeColor, 0, i, 0, i) = &HFFFFFF
            .Cell(flexcpBackColor, 0, i, 0, i) = &H800000
        Next
        .ScrollBars = flexScrollBarNone
        .Redraw = flexRDDirect
    End With
End Sub


'--------------------------------------------------------------------
'  OrderID로 입고와 관련된 Order 데이터 1건 나타내기
'---------------------------------------------------------------------
Private Function FillStuffOrderData(ByVal OrderID As String) As Boolean
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim lNowRow%, sUnit$
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    Set rs = oStuffIn.GetStuffINByOrder(OrderID)
    
    Set oStuffIn = Nothing
    
    grdData.Rows = grdData.FixedRows
    
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        Screen.MousePointer = vbDefault
        FillStuffOrderData = False
        Exit Function
    End If
    
    With grdData
        .Redraw = flexRDNone

        lNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        sUnit = rs!UnitClss
        
        .AddItem MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNO & vbTab & _
                 MakeDate(DF_LONG, rs!AcptDate) & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                 rs!WorkName & vbTab & rs!Width & vbTab & MakeRating(rs!ChunkRate, rs!LossRate) & vbTab & _
                 CheckNum(rs!ColorCnt) & vbTab & SetCurrency(CheckNum(rs!OrderQty)) & vbTab & _
                 SetCurrency(CheckNum(rs!INQty)) & vbTab & rs!UnitClss & vbTab & CheckNum(rs!InRoll)
        
        txtCustomID.Tag = CheckNull(rs!CustomID)
        txtCustomID.Text = CheckNull(rs!kCustom)
        TxtArticleID2.Text = CheckNull(rs!Article)
        TxtArticleID2.Tag = CheckNull(rs!ArticleID)
        txtOrderNO.Text = CheckNull(rs!OrderNO)
        .Redraw = flexRDDirect
    End With
    FillStuffOrderData = True
    grdData.Editable = flexEDKbdMouse
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrHandler:
    FillStuffOrderData = False
    grdData.Redraw = flexRDDirect
    Call ErrorBox(Err.Number, "frmStuffIN.FillStuffOrderData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Function

Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(iRow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        End Select
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub grdGroup_DblClick()
    With grdGroup
        If .Row <= .FixedRows Then Exit Sub
        
        If .IsSubtotal(.Row) Then
            If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                .IsCollapsed(.Row) = flexOutlineExpanded
            Else
                .IsCollapsed(.Row) = flexOutlineCollapsed
            End If
        Else
            Call GetStuffData
        End If
    End With
End Sub

Sub GetStuffData()
    Dim sStuffKey As String
    
    '일자+구분+일련번호( StuffIN Key가져오기 )
    sStuffKey = grdGroup.TextMatrix(grdGroup.Row, grdGroup.Cols - 1)
    
    Call MakeStuffKey(sStuffKey, m_StuffDate, m_StuffClss, m_StuffSeq)
    Call FillGridStuffSub(m_StuffDate, m_StuffClss, m_StuffSeq)
    
    cmdFind(1).Enabled = True
    cmdFind(3).Enabled = True
    cmdFind(4).Enabled = True
End Sub

Private Sub grdGroup_RowColChange()
    With grdGroup
        If .Row <= .FixedRows Then Exit Sub
        
        If .IsSubtotal(.Row) Then
            Call SetClearEdit
        Else
            Call GetStuffData
        End If
    End With
End Sub

Private Sub grdStuffINSub_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With grdStuffINSub
        If Row < .FixedRows Then Exit Sub
        
        If Col = 11 Then
            If Row = .Rows - 1 Then
                .Rows = .Rows + 1
                .Select .Rows - 1, 1
            End If
        Else
            .Select Row, Col + 1
        End If
    End With
    
    Call CalcQty
End Sub
'*******************************************************************************************
'--- 생지 입고 오더별, 거래처별 조회
'*******************************************************************************************
Private Sub FillGridGroup(Optional NewValue As Boolean = True)
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim iTop(2) As Integer, nTop%
    Dim i%, xpName As String
    Dim nCheckNon As Integer
    Dim nTotOrderQty As Long, nTotRoll As Long, nTotQty As Long, nTotColorQty As Long
    Dim StuffClss As String

'    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    If NewValue Then
        xpName = "xp_StuffIN_sStuffIN"
    Else
        xpName = "xp_StuffIN_sStuffIN_Custom"
    End If
    
    ' 확정구분
    If chkSearch(5).Value Then
        nCheckNon = cboOrderID.ListIndex + 1
    Else
        nCheckNon = 0  '전체
    End If
    
    If chkSearch(4).Value Then
        StuffClss = CboStuffClss2.ItemData(CboStuffClss2.ListIndex)
    Else
        StuffClss = ""
    End If
    
    Set rs = oStuffIn.GetStuffIN(xpName, IIf(chkSearch(3) = vbChecked, 1, 0) _
                                , MakeDate(DF_SHORT, dtpDate(0)) _
                                , MakeDate(DF_SHORT, dtpDate(1)) _
                                , IIf(chkSearch(1) = vbChecked, 1, 0) _
                                , txtCustom(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0) _
                                , txtArticle.Tag _
                                , IIf(chkSearch(4) = vbChecked, 1, 0) _
                                , StuffClss _
                                , IIf(chkSearch(0) = vbChecked, 1, 0) _
                                , txtSearch(3).Text _
                                , nCheckNon)

    Set oStuffIn = Nothing
    
    If rs.RecordCount = 0 Then
        grdGroup.Rows = grdGroup.FixedRows
        Exit Sub
    End If
    
    Call InitGroup(NewValue)
    
    nTotOrderQty = 0: nTotRoll = 0: nTotQty = 0: nTotColorQty = 0
    
    '------- 오더별 집계 조회
    If NewValue Then
        
        With grdGroup
            .Redraw = flexRDNone
            .Rows = .FixedRows
            Do Until rs.EOF
                
                '---- 첫번째 그룹설정 (OrderID)
                If Trim(rs!OrderID) <> Trim(.TextMatrix(.Rows - 1, 15)) Then
                    .AddItem " "
                    .TextMatrix(.Rows - 1, 2) = IIf(Trim(rs!OrderID) = "*", "", MakeOrderID(rs!OrderID, OM_EXPAND))
                    .TextMatrix(.Rows - 1, 3) = IIf(Trim(rs!OrderNO) = "", "", rs!OrderNO)
                    .TextMatrix(.Rows - 1, 4) = MakeDate(DF_MD, rs!AcptDate)
                    .TextMatrix(.Rows - 1, 5) = Trim(rs!Custom1)
                    .TextMatrix(.Rows - 1, 6) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 7) = rs!WorkName
                    .TextMatrix(.Rows - 1, 8) = rs!Width
                    .TextMatrix(.Rows - 1, 9) = MakeRating(rs!ChunkRate, rs!LossRate)
                    .TextMatrix(.Rows - 1, 10) = CheckNum(rs!ColorQty)
                    .TextMatrix(.Rows - 1, 11) = SetCurrency(CheckNum(rs!OrderQty))
                    .TextMatrix(.Rows - 1, 12) = 0
                    .TextMatrix(.Rows - 1, 13) = 0
                    .TextMatrix(.Rows - 1, 14) = rs!배색Qty
                    .TextMatrix(.Rows - 1, 15) = rs!OrderID
                    .TextMatrix(.Rows - 1, 16) = rs!OrderNO
                    .TextMatrix(.Rows - 1, 17) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
                    Call GridCollapse(grdGroup, nTop)
                    nTop = .Rows - 1
                    
                    iTop(1) = .Rows - 1
                    
                    nTotOrderQty = nTotOrderQty + CheckNum(rs!OrderQty)
                    nTotColorQty = nTotColorQty + CheckNum(rs!배색Qty)
                End If
'
                .AddItem "" & vbTab & "" & vbTab & "" & vbTab & ""
                .TextMatrix(.Rows - 1, 5) = CheckNull(rs!Custom2)
                .TextMatrix(.Rows - 1, 6) = MakeDate(DF_MID, rs!StuffDate) & "(" + CheckNull(rs!ThreadName) + ")"
                .TextMatrix(.Rows - 1, 12) = rs!StuffRoll
                .TextMatrix(.Rows - 1, 13) = SetCurrency(rs!StuffQty)
                .TextMatrix(.Rows - 1, 15) = rs!OrderID
                .TextMatrix(.Rows - 1, 16) = rs!OrderNO
                .TextMatrix(.Rows - 1, 17) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                         
                         
                '-------입고절수 , 입고수량 Order별로 합계
                .TextMatrix(iTop(1), 12) = SetCurrency(.TextMatrix(iTop(1), 12) + rs!StuffRoll)
                .TextMatrix(iTop(1), 13) = SetCurrency(.TextMatrix(iTop(1), 13) + rs!StuffQty)
                nTotRoll = nTotRoll + CheckNum(rs!StuffRoll)
                nTotQty = nTotQty + CheckNum(rs!StuffQty)
    
                rs.MoveNext
            Loop
            
'            Call ChangeScroll(1)
            
            .Redraw = flexRDDirect
        End With
        
    Else
        '---- 거래처별
        With grdGroup
            .Redraw = flexRDNone
            .Rows = .FixedRows
            
            Do Until rs.EOF
            
                '---- 첫번째 그룹설절  CustomID1 확인
                If Trim(rs!customid1) <> Trim(.TextMatrix(.Rows - 1, 16)) Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 2) = rs!customid1
                    .TextMatrix(.Rows - 1, 3) = rs!Custom1
                    .TextMatrix(.Rows - 1, 12) = 0
                    .TextMatrix(.Rows - 1, 13) = 0
                    .TextMatrix(.Rows - 1, 14) = 0
                    .TextMatrix(.Rows - 1, 15) = 0
                    .TextMatrix(.Rows - 1, 16) = rs!customid1
                    .TextMatrix(.Rows - 1, 18) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
                    Call GridCollapse(grdGroup, nTop)
                    nTop = .Rows - 1
                    
                    iTop(1) = .Rows - 1
                End If
                
                '--- 두번째 그룹설절 OrderID 확인
                If Trim(rs!OrderID) <> Trim(.TextMatrix(.Rows - 1, 17)) Then
                    .AddItem ""
                    
                    .TextMatrix(.Rows - 1, 4) = MakeOrderID(rs!OrderID, OM_EXPAND)
                    .TextMatrix(.Rows - 1, 5) = rs!OrderNO
                    .TextMatrix(.Rows - 1, 6) = MakeDate(DF_MD, rs!AcptDate)
                    .TextMatrix(.Rows - 1, 7) = rs!Article
                    .TextMatrix(.Rows - 1, 8) = rs!WorkName
                    .TextMatrix(.Rows - 1, 9) = rs!Width
                    .TextMatrix(.Rows - 1, 10) = rs!ChunkRate & "+" & rs!LossRate
                    .TextMatrix(.Rows - 1, 11) = rs!ColorQty
                    .TextMatrix(.Rows - 1, 12) = SetCurrency(rs!OrderQty, 0)
                    .TextMatrix(.Rows - 1, 13) = 0
                    .TextMatrix(.Rows - 1, 14) = 0
                    .TextMatrix(.Rows - 1, 15) = rs!배색Qty
                    
                    .TextMatrix(.Rows - 1, 16) = rs!customid1
                    .TextMatrix(.Rows - 1, 17) = rs!OrderID
                    .TextMatrix(.Rows - 1, 18) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 2)
                 '   Call GridCollapse(grdGroup, nTop)
                 '   nTop = .Rows - 1
                    
                    iTop(2) = .Rows - 1
                    .TextMatrix(iTop(1), 12) = SetCurrency(.TextMatrix(iTop(1), 12) + rs!OrderQty)
                    
                    nTotOrderQty = nTotOrderQty + rs!OrderQty
                End If
                
                .AddItem ""
                .TextMatrix(.Rows - 1, 6) = MakeDate(DF_MD, rs!StuffDate)
                .TextMatrix(.Rows - 1, 7) = CheckNull(rs!Custom2) & "(" & rs!ThreadName & ")"
                .TextMatrix(.Rows - 1, 12) = 0
                .TextMatrix(.Rows - 1, 13) = rs!StuffRoll
                .TextMatrix(.Rows - 1, 14) = SetCurrency(rs!StuffQty)
                .TextMatrix(.Rows - 1, 15) = 0
                
                .TextMatrix(.Rows - 1, 16) = rs!customid1
                .TextMatrix(.Rows - 1, 17) = rs!OrderID
                .TextMatrix(.Rows - 1, 18) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                nTotRoll = nTotRoll + rs!StuffRoll
                nTotQty = nTotQty + rs!StuffQty
                
                For i = 1 To 2
                    .TextMatrix(iTop(i), 13) = SetCurrency(.TextMatrix(iTop(i), 13) + rs!StuffRoll)
                    .TextMatrix(iTop(i), 14) = SetCurrency(.TextMatrix(iTop(i), 14) + rs!StuffQty)
                Next i
                .Redraw = flexRDDirect

                rs.MoveNext
            Loop
            .Redraw = flexRDDirect
        End With
    End If
    
    If grdGroup.Rows > grdGroup.FixedRows Then
        grdGroup.Row = grdGroup.FixedRows
    Else
        MsgBox LoadResString(203), vbInformation
    End If
    
    rs.Close
    Set rs = Nothing
    
    Call SetToggle
    
    Call GridCollapse(grdGroup, nTop)
    
    
    With grdTotal
        .TextMatrix(0, 2) = Format(nTotOrderQty, "#,##0 YDS")
        .TextMatrix(0, 4) = Format(nTotRoll, "#,##0 절")
        .TextMatrix(0, 6) = Format(nTotQty, "#,##0 YDS")
        .Redraw = flexRDDirect
    End With
    
    Exit Sub

ErrHandler:
    grdGroup.Redraw = flexRDDirect
    Call ErrorBox(Err.Number, "frmStuffINView.FillGridGroup", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Sub

Private Sub grdStuffINSub_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    '---- 1번 컬럼은 Validateedit에서 확인 하지 않음.
    
    With grdStuffINSub
        If Col <> 1 Then
            If Not IsNumeric(.EditText) Then
                Cancel = True
            End If

            If .EditText = "0" Then
                Cancel = True
            End If

            If InStr(.EditText, "*") > 0 Then
                Cancel = False
            End If

            If Len(.EditText) = 0 Then
                Cancel = False
            End If

        End If
    End With
    
End Sub

Private Sub optGroup_Click(Index As Integer)
    
'    If tabForm.Tab = 0 Then
        If Index = 0 Then
            m_bGroupClss = True
            Call InitGroup
        Else
            m_bGroupClss = False
            Call InitGroup(m_bGroupClss)
        End If
        Call FillGridGroup(m_bGroupClss)
 '       Call optOrder_Click(2)
 '   End If
End Sub

Private Sub optOrder_Click(Index As Integer)
    Call SetToggle

''    Select Case Index
''        Case 2
''            Select Case tabForm.Tab
''                Case 0
''                    With grdGroup
''                        If m_bGroupClss Then
''                            .ColHidden(2) = True
''                            .ColHidden(3) = False
''
''                        Else
''                            .ColHidden(4) = True
''                            .ColHidden(5) = False
''                        End If
''                    End With
''                Case 1
''                    With grdStuffIN
''                        .ColHidden(1) = True
''                        .ColHidden(2) = False
''                    End With
''            End Select
''        Case 3
''            Select Case tabForm.Tab
''                Case 0
''                    With grdGroup
''                        If m_bGroupClss Then
''                            .ColHidden(2) = False
''                            .ColHidden(3) = True
''                        Else
''                            .ColHidden(4) = False
''                            .ColHidden(5) = True
''                        End If
''                    End With
''                Case 1
''                    With grdStuffIN
''                        .ColHidden(1) = False
''                        .ColHidden(2) = True
''                    End With
''            End Select
''    End Select
End Sub

'''Private Sub tabform_Click(PreviousTab As Integer)
'''    If PreviousTab = 1 Then
'''        Call ChangeMode(Me, True)
'''        pnlData.Enabled = False
'''    End If
'''
'''    Select Case tabForm.Tab
'''        Case 0
'''            Call SetClearEdit
'''''          Call cmdSearch_Click
'''''        Case 1
'''''            If Trim(pnlMsg) = "" Or val(txtStuffSeq.Text) = 0 Then
'''''                Call cmdOperate_Click(0)
'''''            End If
'''    End Select
'''
'''End Sub

Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Call MoveFocus(KeyAscii)
    End If

End Sub

Private Sub TxtArticleID2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(4)
'        Call ReturnCode(LG_ARTICLE, , False, TxtArticleID2)
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Call MoveFocus(KeyAscii)
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call ReturnCode(LG_CUSTOM, , False, txtCustom(Index))
                Call MoveFocus(KeyAscii)
            End If
    End Select
End Sub

Private Sub txtCustomID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_CUSTOM, , False, txtCustomID)
        Call MoveFocus(KeyAscii)
    End If
End Sub


Private Sub txtNum_GotFocus(Index As Integer)
    Call GotFocusText(txtNum(Index))
End Sub

Private Sub txtNum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub

Private Sub txtNum_LostFocus(Index As Integer)
    If IsNumeric(txtNum(Index)) Then
        Select Case Index
            Case 1
                txtNum(Index) = Format(txtNum(Index), "###,##0")
    
            Case Else
                txtNum(Index) = SetCurrency(txtNum(Index))
        End Select
    Else
        txtNum(Index).Text = 0
    End If
End Sub

Private Sub txtOrderNO_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

Private Sub txtOrderNO_LostFocus()
    
    If Len(txtOrderNO) > 0 Then
        Call cmdFind_Click(3)
    End If
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MoveFocus (KeyAscii)
    End If
End Sub

Private Sub txtOrderID_LostFocus()
    Dim dCheck_bol As Boolean
    
    If Len(txtOrderID) = 10 Then
        If FillStuffOrderData(txtOrderID) Then
            txtCustomID.Enabled = False
            TxtArticleID2.Enabled = False
        Else
            txtCustomID.Enabled = True
            TxtArticleID2.Enabled = True
        End If
    Else
        txtOrderID = ""
        grdData.Rows = grdData.FixedRows
    End If
End Sub

Private Sub txtThreadName_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

