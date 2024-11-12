VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCardPattern 
   ClientHeight    =   9255
   ClientLeft      =   735
   ClientTop       =   210
   ClientWidth     =   11850
   Icon            =   "frmCardPattern.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Crystal.CrystalReport cryReport 
      Left            =   7230
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlMessage 
      Height          =   795
      Left            =   30
      TabIndex        =   45
      Top             =   8430
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   1402
      _Version        =   196609
      Enabled         =   0   'False
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtMessage 
         Appearance      =   0  '평면
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  '없음
         Height          =   555
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   46
         Text            =   "frmCardPattern.frx":000C
         Top             =   120
         Width           =   7845
      End
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "취소(&C)"
      Height          =   780
      Index           =   4
      Left            =   10155
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   30
      ToolTipText     =   "자료 취소"
      Top             =   7500
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "수정(&U)"
      Height          =   780
      Index           =   1
      Left            =   10935
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   29
      ToolTipText     =   "자료 수정"
      Top             =   7500
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "저장(&S)"
      Height          =   780
      Index           =   3
      Left            =   9360
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   28
      ToolTipText     =   "자료 저장"
      Top             =   7500
      Visible         =   0   'False
      Width           =   780
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   390
      TabIndex        =   25
      Top             =   2940
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
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSFrame frmPattern 
      Height          =   4395
      Left            =   0
      TabIndex        =   20
      Top             =   3990
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   7752
      _Version        =   196609
      Begin VB.ComboBox cboUseClss 
         Height          =   300
         Left            =   4800
         TabIndex        =   39
         Text            =   "cboUseClss"
         Top             =   90
         Width           =   1545
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   90
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "현재 공정 패턴"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdProcess 
         Height          =   2985
         Left            =   9360
         TabIndex        =   23
         Top             =   450
         Width           =   2325
         _cx             =   4101
         _cy             =   5265
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
      Begin VSFlex7LCtl.VSFlexGrid grdNewPattern 
         Height          =   3855
         Left            =   1650
         TabIndex        =   22
         Top             =   450
         Width           =   6435
         _cx             =   11351
         _cy             =   6800
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
      Begin VSFlex7LCtl.VSFlexGrid grdCardPattern 
         Height          =   3855
         Left            =   90
         TabIndex        =   21
         Top             =   450
         Width           =   1485
         _cx             =   2619
         _cy             =   6800
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
         Height          =   315
         Index           =   1
         Left            =   1650
         TabIndex        =   32
         Top             =   90
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "새 공정 패턴"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   795
         Left            =   8250
         TabIndex        =   33
         Top             =   3510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "삭제"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   795
         Left            =   8250
         TabIndex        =   34
         Top             =   2580
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "추가"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdUP 
         Height          =   795
         Left            =   8250
         TabIndex        =   35
         Top             =   420
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "위"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   795
         Left            =   8250
         TabIndex        =   36
         Top             =   1320
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "아래"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   2
         Left            =   9360
         TabIndex        =   37
         Top             =   90
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "공 정 명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   3
         Left            =   3240
         TabIndex        =   38
         Top             =   90
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "카드상태"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   3105
      Left            =   0
      TabIndex        =   19
      Top             =   870
      Width           =   11835
      _cx             =   20876
      _cy             =   5477
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
   Begin Threed.SSFrame frmSearch 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1508
      _Version        =   196609
      Begin VB.TextBox txtSearch 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   10140
         MaxLength       =   4
         TabIndex        =   44
         Top             =   495
         Width           =   525
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   40
         Top             =   495
         Width           =   885
      End
      Begin VB.ComboBox cboProcess 
         Height          =   300
         Left            =   6120
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   495
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   720
         Left            =   10980
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   4
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   6120
         TabIndex        =   3
         Top             =   120
         Width           =   1665
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2640
         TabIndex        =   2
         Top             =   495
         Width           =   1785
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   1
         Top             =   120
         Width           =   1785
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1191
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   90
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   390
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   9
         Top             =   120
         Width           =   1230
         _ExtentX        =   2170
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
         Left            =   4440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
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
         Left            =   1380
         TabIndex        =   12
         Top             =   495
         Width           =   1230
         _ExtentX        =   2170
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
         Left            =   4440
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
         Left            =   4830
         TabIndex        =   15
         Top             =   120
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
         Left            =   7890
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
         Left            =   4830
         TabIndex        =   41
         Top             =   495
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
            Caption         =   "대기공정"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   7890
         TabIndex        =   47
         Top             =   120
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
            Caption         =   "완료건포함"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   1215
         End
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10110
      TabIndex        =   24
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8340
      TabIndex        =   43
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      발행(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmCardPattern"
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
    Select Case Index
    Case 0
        chkSearch(4).Value = chkSearch(0).Value
        txtSearch(4).Enabled = chkSearch(0).Value
        txtSearch(5).Enabled = chkSearch(0).Value
        If chkSearch(0).Value = vbChecked Then
            txtSearch(4).SetFocus
        End If
        
    Case 1, 2, 3
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
    Case 4
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(4).Enabled = True
            txtSearch(5).Enabled = True
            txtSearch(4).SetFocus
        Else
            txtSearch(4).Enabled = False
            txtSearch(5).Enabled = False
        End If
    Case Else
        If chkSearch(Index).Value = vbChecked Then
            cboProcess.Enabled = True
            cboProcess.SetFocus
        Else
            cboProcess.Enabled = False
        End If
    End Select
End Sub

Private Sub cmdAdd_Click()
    Dim i%
    
    With grdNewPattern
        If .Row < .Rows - 1 Then
            If .TextMatrix(.Row + 1, 3) = "*" Then
                MsgBox "작업이 완료된 공정안에는 공정을 추가할 수 없습니다.", vbInformation + vbOKOnly
                Exit Sub
            End If
        End If
        
        .Redraw = flexRDNone
        
        .AddItem vbTab & grdProcess.TextMatrix(grdProcess.Row, 1) & vbTab & grdProcess.TextMatrix(grdProcess.Row, 2) & vbTab & "" & vbTab & 0, .Row + 1
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim sCardID$, sSplitID$, sPatternID$
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    With grdData
        sCardID = MakeCardID(.TextMatrix(.Row, 6), OM_REDUCE)
        sSplitID = .TextMatrix(.Row, 7)
        sPatternID = .TextMatrix(.Row, 16)
    End With
    
    Call PrintWorkCard(cryReport, sCardID, sSplitID, sPatternID, PlusMDI.PrintPreview)
End Sub


Private Sub grdNewPattern_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdNewPattern
        If Col = 4 Then
            If IsNumeric(.TextMatrix(Row, Col)) Then
                .Select Row, Col + 1
            Else
                .TextMatrix(Row, Col) = "0"
            End If
        ElseIf Col = 5 Then
            .Select Row, Col + 1
        ElseIf Col = 6 Then
            If Row < .Rows - 1 Then
                .Select Row + 1, 4
            End If
        End If
    End With
End Sub

Private Sub grdNewPattern_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 4 Or grdNewPattern.TextMatrix(grdNewPattern.Row, 3) = "*" Then Cancel = True
End Sub

Private Sub grdNewPattern_Click()
    
    With grdNewPattern
'        If Trim(.TextMatrix(.Row, 3)) = "" And .Col = 7 Then
'            If .Cell(flexcpChecked, .Row, 7) = flexChecked Then
'                .Cell(flexcpChecked, .Row, 7) = flexUnchecked
'            Else
'                .Cell(flexcpChecked, .Row, 7) = flexChecked
'            End If
'        End If
    End With
End Sub

Private Sub grdProcess_DblClick()
    Dim i%
    With grdNewPattern
        If .Row < .Rows - 1 Then
            If .TextMatrix(.Row + 1, 3) = "*" Then
                MsgBox "작업이 완료된 공정안에는 공정을 추가할 수 없습니다.", vbInformation + vbOKOnly
                Exit Sub
            End If
        End If
        
        .Redraw = flexRDNone
        
        .AddItem vbTab & grdProcess.TextMatrix(grdProcess.Row, 1) & vbTab & grdProcess.TextMatrix(grdProcess.Row, 2) & vbTab & "" & vbTab & 0, .Row + 1
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdDel_Click()
    Dim i%
    
    With grdNewPattern
        If .TextMatrix(.Row, 3) = "*" Then
            MsgBox "작업이 완료된 공정은 공정을 삭제할 수 없습니다.", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        .RemoveItem .Row
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdUP_Click()
    Dim i%
    Dim vTemp(6) As String
    
    With grdNewPattern
        If .Row <= .FixedRows Then Exit Sub
        
        If .TextMatrix(.Row, 3) = "*" Or .TextMatrix(.Row - 1, 3) = "*" Then
            MsgBox "작업이 완료된 공정에 대해서는 순서를 변경할 수 없습니다.", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        For i = 1 To 6
            vTemp(i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = .TextMatrix(.Row - 1, i)
            .TextMatrix(.Row - 1, i) = vTemp(i)
        Next i
        .Select .Row - 1, 1
    End With
End Sub

Private Sub cmdDown_Click()
    Dim i%
    Dim vTemp(6) As String
    
    With grdNewPattern
        If .Row = .Rows - 1 Then Exit Sub
        
        If .TextMatrix(.Row, 3) = "*" Then
            MsgBox "작업이 완료된 공정에 대해서는 순서를 변경할 수 없습니다.", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        For i = 1 To 6
            vTemp(i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = .TextMatrix(.Row + 1, i)
            .TextMatrix(.Row + 1, i) = vTemp(i)
        Next i
        .Select .Row + 1, 1
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
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
            If grdData.TextMatrix(grdData.Row, 13) = "작업" Then
                MsgBox "작업중인 카드는 카드변경을 할 수 없습니다.", vbInformation + vbOKOnly
                Exit Sub
            End If
            
            If grdCardPattern.Rows = grdCardPattern.FixedRows Then Exit Sub
            If grdProcess.Rows = grdProcess.FixedRows Then Exit Sub
            grdProcess.Row = grdProcess.FixedRows
            m_iFlag = ID_UPDATE
            
            Call ChangeMode(Me, False)
            Call ModeChange(False)
            
            grdNewPattern.SetFocus
            grdNewPattern.Select 1, 1
        '-------------------------------------------------------------------------------------'
        Case ID_SAVE
            If SaveData() Then
                m_iFlag = -1
                Call ChangeMode(Me, True)
                Call ModeChange(True)
                Call FillGridData
              
                
            End If
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL
            m_iFlag = -1
            Call ChangeMode(Me, True)
            Call ModeChange(True)
            Call FillGridData
    End Select

    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmCardPattern.cmdOperate_Click", Err.Description)
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
    Call FillGridProcess
    Call ModeChange(True)
   
   With cboUseClss
        .AddItem "대기"
        .AddItem "보류"
        
        .ListIndex = -1
   End With
   
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    txtSearch(4).Enabled = False
    txtSearch(5).Enabled = False
    cboProcess.Enabled = False
    
    cmdAdd.Picture = LoadResPicture("BACK", vbResIcon)
    cmdDel.Picture = LoadResPicture("FRONT", vbResIcon)
    cmdUP.Picture = LoadResPicture("UP", vbResIcon)
    cmdDown.Picture = LoadResPicture("DOWN", vbResIcon)
    
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
        .Cols = 17
        
        Call SetVSFlexGrid(grdData)
        .Rows = 1
        .RowHeightMin = 390
        
        .TextArray(0) = " ":
        .TextArray(1) = " ":                          .ColWidth(1) = 250:             .ColHidden(1) = True
        .TextArray(2) = "거래처":                     .ColWidth(2) = 1000:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":                       .ColWidth(3) = 1700:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":                   .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNo":                    .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "카드번호":                   .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "분할" & vbCrLf & "번호":     .ColWidth(7) = 500:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":                     .ColWidth(8) = 1000:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "절수":                       .ColWidth(9) = 500:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "수량":                      .ColWidth(10) = 600:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "완료공정":                  .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "대기공정":                  .ColWidth(12) = 900:            .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "카드상태":                  .ColWidth(13) = 900:            .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "계획공정":                  .ColWidth(14) = 7000:           .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "패턴":                      .ColHidden(15) = True
        .TextArray(16) = "패턴코드":                  .ColHidden(16) = True
        
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
    
    With grdCardPattern
        .Redraw = flexRDNone
        .Cols = 7
        
        Call SetVSFlexGrid(grdCardPattern)
        .Rows = 1
        
        .TextArray(0) = "순서":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColHidden(1) = True
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "완료여부":     .ColHidden(3) = True
        .TextArray(4) = "요구폭":       .ColHidden(4) = True
        .TextArray(5) = "지시사항":     .ColHidden(5) = True
        .TextArray(6) = "비고":         .ColHidden(6) = True
        
        .HighLight = flexHighlightNever
        .Redraw = flexRDDirect
    End With
    
    With grdNewPattern
        .Redraw = flexRDNone
        .Cols = 8
        
        Call SetVSFlexGrid(grdNewPattern)
        .Rows = 1
        
        .TextArray(0) = "순서":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColWidth(1) = 0
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "완료여부":     .ColWidth(3) = 0
        .TextArray(4) = "요구폭":       .ColWidth(4) = "700":       .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "지시사항":     .ColWidth(5) = "2300":      .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "비고":         .ColWidth(6) = "1000":      .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "재작업":       .ColWidth(7) = "500":       .ColAlignment(7) = flexAlignLeftCenter:     .ColDataType(7) = flexDTBoolean

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarBoth
        .ExtendLastCol = True
        
        .Redraw = flexRDDirect
    End With
    
    With grdProcess
        .Redraw = flexRDNone
        .Cols = 3
        
        Call SetVSFlexGrid(grdProcess)
        .Rows = 1
        
        .TextArray(0) = "":             .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColWidth(1) = 0
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        
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

Private Sub FillGridProcess()
    Dim oCard As PlusLib2.CCard
    Dim rs As Recordset
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon

    Set rs = oCard.GetProcess()
    Set oCard = Nothing

    With grdProcess
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!ProcessID & vbTab & rs!Process
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
    End With

    m_bloading = False
    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    m_bloading = False
    Call ErrorBox(Err.Number, "frmCardPattern.FillGridProcess", Err.Description)
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
                                 IIf(chkSearch(5) = vbChecked, 1, 0), Format(Left(cboProcess.ItemData(cboProcess.ListIndex), 2), "00"), _
                                 IIf(chkSearch(0) = vbChecked, 1, 0))
    Set oCard = Nothing
       
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & _
                    rs!Color & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                    rs!UseClss & vbTab & CheckNull(rs!AfterProc) & vbTab & rs!Pattern & vbTab & rs!PatternID
            
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
            grdCardPattern.Rows = grdCardPattern.FixedRows
            grdNewPattern.Rows = grdNewPattern.FixedRows
            
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
    Call ErrorBox(Err.Number, "frmCardPattern.FillGridData", Err.Description)
End Sub

Private Sub FillGridPattern()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetCardPattern(MakeCardID(grdData.TextMatrix(grdData.Row, 6), OM_REDUCE), grdData.TextMatrix(grdData.Row, 7))
    Set oCard = Nothing
    
    If rs.EOF Then
        grdCardPattern.Rows = grdCardPattern.FixedRows
        grdNewPattern.Rows = grdNewPattern.FixedRows
        Set rs = Nothing
        Exit Sub
    End If
    
    With grdCardPattern
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        For i = 0 To rs.RecordCount - 1
            .AddItem rs!PlanSeq & vbTab & rs!ProcessID & vbTab & rs!Process & vbTab & rs!CompleteClss & vbTab & _
                rs!NeedWidth & vbTab & rs!InstRemark & vbTab & rs!Remark
            
            If rs!CompleteClss = "*" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
            End If
            
            rs.MoveNext
        Next i
        
        .Redraw = flexRDDirect
    End With
    
    rs.MoveFirst
    With grdNewPattern
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        For i = 0 To rs.RecordCount - 1
            .AddItem rs!PlanSeq & vbTab & rs!ProcessID & vbTab & rs!Process & vbTab & rs!CompleteClss & vbTab & _
                rs!NeedWidth & vbTab & rs!InstRemark & vbTab & rs!Remark & vbTab & ""
                
            If rs!ReWorkClss = "*" Then
                .Cell(flexcpChecked, .Rows - 1, 7) = flexChecked
            End If
            
            If rs!CompleteClss = "*" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
            End If
            
            rs.MoveNext
        Next i
        
        .Redraw = flexRDDirect
    End With
    rs.Close
    Set rs = Nothing
        
    Exit Sub
    
ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmCardPattern.FillGridPattern", Err.Description)
End Sub

Private Sub ModeChange(bValue As Boolean)
    frmSearch.Enabled = bValue
    frmPattern.Enabled = Not bValue
    grdData.Enabled = bValue
    pnlMessage.Visible = Not bValue
End Sub

Private Sub ShowData()
    With grdData
        
        If .TextMatrix(.Row, 13) = "완료" Then
            cboUseClss = "대기"
            cboUseClss.Tag = "대기"

        ElseIf .TextMatrix(.Row, 13) <> "작업" Then
            cboUseClss = .TextMatrix(.Row, 13)
            cboUseClss.Tag = .TextMatrix(.Row, 13)
        
        End If
        
    End With
    
    Call FillGridPattern
End Sub

Private Function SaveData() As Boolean
    Dim tItem As PlusLib2.TCard
    Dim tItemSub() As PlusLib2.TCardPattern
    Dim oCard As PlusLib2.CCard
    Dim i%
    
    On Error GoTo ErrHandler
    
    SaveData = False
    
    With grdData
        tItem.sCardID = MakeCardID(.TextMatrix(.Row, 6), OM_REDUCE)
        tItem.sSplitID = .TextMatrix(.Row, 7)
        tItem.sUseClss = cboUseClss
        tItem.sPersonID = g_sUserName
        tItem.nChkUseClss = 0
        If cboUseClss <> cboUseClss.Tag And cboUseClss.Tag = "보류" Then
            tItem.nChkUseClss = 1   '보류에서 대기로 변경될때 Hold Table 보류 취소 업데이트
        End If
    End With
    
    With grdCardPattern
'        tItem.sPrePlanProc = grdData.TextMatrix(grdData.Row, 15) & "::"
        For i = .FixedRows To .Rows - .FixedRows
            tItem.sPrePlanProc = tItem.sPrePlanProc & .TextMatrix(i, 2) & "→"
        Next i
        tItem.sPrePlanProc = Left(tItem.sPrePlanProc, Len(tItem.sPrePlanProc) - 1)
    End With
    
    With grdNewPattern
'        tItem.sPostPlanProc = grdData.TextMatrix(grdData.Row, 15) & "::"
        For i = .FixedRows To .Rows - .FixedRows
            tItem.sPostPlanProc = tItem.sPostPlanProc & .TextMatrix(i, 2) & "→"
        Next i
        tItem.sPostPlanProc = Left(tItem.sPostPlanProc, Len(tItem.sPostPlanProc) - 1)
    End With
    
    With grdNewPattern
'        tItem.sAfterProc = grdData.TextMatrix(grdData.Row, 15) & "::"
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 3) <> "*" Then
                tItem.sAfterProc = tItem.sAfterProc & .TextMatrix(i, 2) & "→"
            End If
        Next i
        tItem.sAfterProc = Left(tItem.sAfterProc, Len(tItem.sAfterProc) - 1)
    End With
    
    With grdNewPattern
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 3) <> "*" Then
                tItem.sWaitProcID = .TextMatrix(i, 1)
                tItem.nWaitProcSeq = .TextMatrix(i, 0)
                Exit For
            End If
        Next i
    End With
    
    With grdNewPattern
        ReDim tItemSub(.Rows - .FixedRows - 1)
        
        For i = .FixedRows To .Rows - 1
            tItemSub(i - 1).sCardID = tItem.sCardID
            tItemSub(i - 1).sSplitID = tItem.sSplitID
            tItemSub(i - 1).nPlanSeq = .TextMatrix(i, 0)
            tItemSub(i - 1).sProcessID = .TextMatrix(i, 1)
            tItemSub(i - 1).sCompleteClss = .TextMatrix(i, 3)
            tItemSub(i - 1).nNeedWidth = .TextMatrix(i, 4)
            tItemSub(i - 1).sInstRemark = .TextMatrix(i, 5)
            tItemSub(i - 1).sRemark = .TextMatrix(i, 6)
            tItemSub(i - 1).sReWorkClss = IIf(.Cell(flexcpChecked, i, 7) = flexChecked, "*", "")
            
        Next i
    End With
   
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    oCard.UserName = g_sUserName
    
    If oCard.UpdateCardPattern(tItem, tItemSub()) Then
        SaveData = True
    End If
    Set oCard = Nothing
    Exit Function
ErrHandler:
    Set oCard = Nothing
    SaveData = False
    Call ErrorBox(Err.Number, "frmCardPattern.SaveData", Err.Description)
End Function



