VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrder 
   Caption         =   "수주 관리(2010)"
   ClientHeight    =   9360
   ClientLeft      =   2145
   ClientTop       =   2220
   ClientWidth     =   16920
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   16920
   Begin Threed.SSPanel SSPanel1 
      Height          =   3345
      Left            =   12855
      TabIndex        =   153
      Top             =   1650
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   5900
      _Version        =   196610
      Caption         =   "SSPanel1"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   23
         Left            =   60
         Style           =   2  '드롭다운 목록
         TabIndex        =   154
         Top             =   60
         Width           =   2445
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPattern 
         Height          =   2865
         Left            =   75
         TabIndex        =   155
         Top             =   420
         Width           =   2400
         _cx             =   4233
         _cy             =   5054
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
      End
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   2970
      Top             =   6615
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   330
      Left            =   15
      TabIndex        =   99
      Top             =   8100
      Width           =   3900
      _cx             =   6879
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
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   4605
      Left            =   30
      TabIndex        =   98
      Top             =   3480
      Width           =   3900
      _cx             =   6879
      _cy             =   8123
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
   Begin VB.Frame fraSearch 
      Height          =   3495
      Left            =   30
      TabIndex        =   89
      Top             =   -30
      Width           =   3900
      Begin VB.ComboBox cboSearch 
         Height          =   300
         Index           =   2
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   152
         Top             =   3120
         Width           =   1875
      End
      Begin VB.ComboBox cboSearch 
         Height          =   300
         Index           =   1
         Left            =   1455
         Style           =   2  '드롭다운 목록
         TabIndex        =   145
         Top             =   2745
         Width           =   1875
      End
      Begin VB.ComboBox cboSearch 
         Height          =   300
         Index           =   0
         Left            =   1455
         Style           =   2  '드롭다운 목록
         TabIndex        =   144
         Top             =   2385
         Width           =   1890
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   134
         Top             =   1650
         Width           =   1905
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3390
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1305
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196610
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
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   63
         Top             =   1305
         Width           =   1905
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   2985
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   66
         ToolTipText     =   "자료 저장"
         Top             =   210
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   1440
         TabIndex        =   65
         Top             =   2010
         Width           =   1905
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   345
         MousePointer    =   99  '사용자 정의
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   900
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   345
         MousePointer    =   99  '사용자 정의
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   540
         Width           =   600
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   1005
         TabIndex        =   60
         Top             =   540
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   138477569
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1005
         TabIndex        =   61
         Top             =   900
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   138477569
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   90
         Top             =   1305
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
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
            Caption         =   "거 래 처"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Top             =   2010
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
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
            TabIndex        =   64
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   92
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
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
            Caption         =   "수주 일자"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   57
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   135
         Top             =   1650
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
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
            TabIndex        =   136
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   3390
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   1650
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196610
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
         Index           =   4
         Left            =   120
         TabIndex        =   140
         Top             =   2385
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
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
            Caption         =   "원단구분"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   141
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   142
         Top             =   2745
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
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
            Caption         =   "완료구분"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   143
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   120
         TabIndex        =   151
         Top             =   3120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "패턴지정"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   2295
         TabIndex        =   94
         Top             =   990
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   2295
         TabIndex        =   93
         Top             =   615
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   12015
      TabIndex        =   73
      Top             =   8475
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13740
      TabIndex        =   74
      Top             =   8475
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196610
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
   Begin TabDlg.SSTab tabOrder 
      Height          =   4800
      Left            =   3915
      TabIndex        =   87
      Top             =   4005
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   8467
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   741
      TabCaption(0)   =   "  수주 색상  "
      TabPicture(0)   =   "frmOrder.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtBox(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pnlName(53)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pnlName(52)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlName(49)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "pnlName(18)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtBox(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtBox(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "pnlName(13)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pnlName(17)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBox(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pnlName(38)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdPlus"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdErase"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "grdColor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboName(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboName(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboName(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtUnitPrice"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboSubulWidth"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboUnitPriceClss"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "  공정 기재사항  "
      TabPicture(1)   =   "frmOrder.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1(2)"
      Tab(1).Control(1)=   "Line1(3)"
      Tab(1).Control(2)=   "txtName(10)"
      Tab(1).Control(3)=   "txtName(9)"
      Tab(1).Control(4)=   "pnlName(23)"
      Tab(1).Control(5)=   "pnlName(1)"
      Tab(1).Control(6)=   "pnlName(47)"
      Tab(1).Control(7)=   "txtBox(2)"
      Tab(1).Control(8)=   "txtBox(0)"
      Tab(1).Control(9)=   "txtBox(5)"
      Tab(1).Control(10)=   "txtBox(1)"
      Tab(1).Control(11)=   "pnlName(35)"
      Tab(1).Control(12)=   "pnlName(34)"
      Tab(1).Control(13)=   "pnlName(33)"
      Tab(1).Control(14)=   "pnlName(2)"
      Tab(1).Control(15)=   "pnlName(16)"
      Tab(1).Control(16)=   "cboName(3)"
      Tab(1).Control(17)=   "cboName(9)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "  검사 기재사항  "
      TabPicture(2)   =   "frmOrder.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTagDestition"
      Tab(2).Control(1)=   "txtTagQuality"
      Tab(2).Control(2)=   "cboName(7)"
      Tab(2).Control(3)=   "cboName(8)"
      Tab(2).Control(4)=   "cboName(11)"
      Tab(2).Control(5)=   "cboName(10)"
      Tab(2).Control(6)=   "cboName(12)"
      Tab(2).Control(7)=   "cboName(13)"
      Tab(2).Control(8)=   "txtRemark"
      Tab(2).Control(9)=   "txtEndMark"
      Tab(2).Control(10)=   "txtTag"
      Tab(2).Control(11)=   "cboName(15)"
      Tab(2).Control(12)=   "cboName(16)"
      Tab(2).Control(13)=   "pnlName(0)"
      Tab(2).Control(14)=   "txtName(8)"
      Tab(2).Control(15)=   "pnlName(45)"
      Tab(2).Control(16)=   "pnlName(44)"
      Tab(2).Control(17)=   "pnlName(24)"
      Tab(2).Control(18)=   "pnlName(25)"
      Tab(2).Control(19)=   "pnlName(15)"
      Tab(2).Control(20)=   "pnlName(26)"
      Tab(2).Control(21)=   "pnlName(28)"
      Tab(2).Control(22)=   "pnlName(31)"
      Tab(2).Control(23)=   "pnlName(32)"
      Tab(2).Control(24)=   "pnlName(22)"
      Tab(2).Control(25)=   "pnlName(40)"
      Tab(2).Control(26)=   "pnlName(43)"
      Tab(2).Control(27)=   "pnlName(30)"
      Tab(2).Control(28)=   "txtName(4)"
      Tab(2).Control(29)=   "txtName(3)"
      Tab(2).Control(30)=   "txtName(2)"
      Tab(2).Control(31)=   "txtName(6)"
      Tab(2).Control(32)=   "pnlName(29)"
      Tab(2).Control(33)=   "pnlName(36)"
      Tab(2).Control(34)=   "pnlName(27)"
      Tab(2).Control(35)=   "txtName(11)"
      Tab(2).Control(36)=   "pnlName(6)"
      Tab(2).Control(37)=   "txtName(5)"
      Tab(2).Control(38)=   "pnlName(56)"
      Tab(2).Control(39)=   "pnlName(57)"
      Tab(2).Control(40)=   "pnlName(21)"
      Tab(2).Control(41)=   "txtName(7)"
      Tab(2).Control(42)=   "pnlName(20)"
      Tab(2).Control(43)=   "txtName(13)"
      Tab(2).Control(44)=   "pnlName(51)"
      Tab(2).Control(45)=   "txtName(14)"
      Tab(2).Control(46)=   "Line1(0)"
      Tab(2).ControlCount=   47
      Begin VB.ComboBox cboUnitPriceClss 
         BackColor       =   &H00D1F3FE&
         Height          =   300
         Left            =   3840
         Style           =   2  '드롭다운 목록
         TabIndex        =   165
         Top             =   495
         Width           =   765
      End
      Begin VB.TextBox txtTagDestition 
         Height          =   285
         Left            =   -73290
         TabIndex        =   48
         Top             =   4005
         Width           =   3165
      End
      Begin VB.TextBox txtTagQuality 
         Height          =   330
         Left            =   -68430
         TabIndex        =   162
         Top             =   3960
         Width           =   2310
      End
      Begin VB.ComboBox cboSubulWidth 
         Height          =   300
         Left            =   5130
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   90
         Width           =   975
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   7170
         TabIndex        =   157
         Top             =   1020
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   0
         ItemData        =   "frmOrder.frx":0060
         Left            =   6810
         List            =   "frmOrder.frx":0062
         TabIndex        =   24
         Text            =   "cboName"
         Top             =   90
         Width           =   1005
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   4
         Left            =   870
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   90
         Width           =   1155
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   9
         Left            =   -69840
         Style           =   2  '드롭다운 목록
         TabIndex        =   31
         Top             =   60
         Width           =   2205
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   3
         Left            =   -74895
         Style           =   2  '드롭다운 목록
         TabIndex        =   26
         Top             =   435
         Width           =   1155
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   7
         Left            =   -74910
         Style           =   2  '드롭다운 목록
         TabIndex        =   34
         Top             =   390
         Width           =   1590
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   8
         Left            =   -73290
         Style           =   2  '드롭다운 목록
         TabIndex        =   35
         Top             =   390
         Width           =   1590
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   11
         Left            =   -73290
         Style           =   2  '드롭다운 목록
         TabIndex        =   38
         Top             =   1125
         Width           =   1590
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   10
         Left            =   -74910
         Style           =   2  '드롭다운 목록
         TabIndex        =   37
         Top             =   1125
         Width           =   1590
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   12
         Left            =   -71655
         Style           =   2  '드롭다운 목록
         TabIndex        =   39
         Top             =   1125
         Width           =   1590
      End
      Begin VB.ComboBox cboName 
         BackColor       =   &H00FFC0C0&
         Height          =   300
         Index           =   13
         ItemData        =   "frmOrder.frx":0064
         Left            =   -68430
         List            =   "frmOrder.frx":0066
         Style           =   2  '드롭다운 목록
         TabIndex        =   51
         Top             =   795
         Width           =   2310
      End
      Begin VB.TextBox txtRemark 
         Height          =   780
         Left            =   -68430
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   55
         Top             =   3120
         Width           =   2310
      End
      Begin VB.TextBox txtEndMark 
         Height          =   780
         Left            =   -68430
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   54
         Top             =   2310
         Width           =   2310
      End
      Begin VB.TextBox txtTag 
         Height          =   780
         Left            =   -68430
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   53
         Top             =   1485
         Width           =   2310
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   15
         Left            =   -71640
         Style           =   2  '드롭다운 목록
         TabIndex        =   36
         Top             =   390
         Width           =   1590
      End
      Begin VB.ComboBox cboName 
         BackColor       =   &H00FFC0C0&
         Height          =   300
         Index           =   16
         ItemData        =   "frmOrder.frx":0068
         Left            =   -68430
         List            =   "frmOrder.frx":006A
         Style           =   2  '드롭다운 목록
         TabIndex        =   52
         Top             =   1125
         Width           =   2310
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   6
         Left            =   1965
         Style           =   2  '드롭다운 목록
         TabIndex        =   69
         Top             =   495
         Width           =   885
      End
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   3285
         Left            =   60
         TabIndex        =   71
         Top             =   1020
         Width           =   8850
         _cx             =   15610
         _cy             =   5794
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
      Begin Threed.SSCommand cmdErase 
         Height          =   495
         Left            =   7755
         TabIndex        =   70
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "색상 삭제"
      End
      Begin Threed.SSCommand cmdPlus 
         Height          =   495
         Left            =   6555
         TabIndex        =   25
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   873
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "색상 추가"
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   38
         Left            =   90
         TabIndex        =   88
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "총 주문량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   7
         Left            =   975
         TabIndex        =   103
         Top             =   495
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   -74910
         TabIndex        =   104
         Top             =   3315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG Remark1"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   8
         Left            =   -73290
         TabIndex        =   46
         Top             =   3300
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
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
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   45
         Left            =   -69915
         TabIndex        =   105
         Top             =   1500
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG 내용"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   44
         Left            =   -69915
         TabIndex        =   106
         Top             =   2310
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "End Mark"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   24
         Left            =   -74910
         TabIndex        =   107
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LABEL"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   25
         Left            =   -73290
         TabIndex        =   108
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "BAND"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   15
         Left            =   -73275
         TabIndex        =   109
         Top             =   795
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "MadeKorea"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   26
         Left            =   -74910
         TabIndex        =   110
         Top             =   795
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "EndMark"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   28
         Left            =   -71640
         TabIndex        =   111
         Top             =   795
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사용 면"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   31
         Left            =   -73290
         TabIndex        =   112
         Top             =   1515
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Advn 샘플"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   32
         Left            =   -71670
         TabIndex        =   113
         Top             =   1515
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LOT 샘플"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   22
         Left            =   -74910
         TabIndex        =   114
         Top             =   2595
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG 품명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   40
         Left            =   -69930
         TabIndex        =   115
         Top             =   795
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "검사 기준"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   43
         Left            =   -69915
         TabIndex        =   116
         Top             =   3120
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "비고 사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   30
         Left            =   -74910
         TabIndex        =   117
         Top             =   1515
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ship 샘플"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   4
         Left            =   -71685
         TabIndex        =   42
         Top             =   1860
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
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
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   3
         Left            =   -73290
         TabIndex        =   41
         Top             =   1860
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
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
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   2
         Left            =   -74910
         TabIndex        =   40
         Top             =   1860
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
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
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   6
         Left            =   -73290
         TabIndex        =   44
         Top             =   2595
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   29
         Left            =   -71640
         TabIndex        =   118
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   36
         Left            =   -69915
         TabIndex        =   119
         Top             =   1125
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "검사 단위"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   16
         Left            =   -74895
         TabIndex        =   120
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "생지 폭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   2
         Left            =   -73635
         TabIndex        =   121
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "생지 중량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   33
         Left            =   -73635
         TabIndex        =   122
         Top             =   810
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가공 밀도"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   34
         Left            =   -74895
         TabIndex        =   123
         Top             =   810
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가공 중량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   35
         Left            =   -72345
         TabIndex        =   124
         Top             =   810
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "감량율 (%)"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   1
         Left            =   -74895
         TabIndex        =   28
         Top             =   1155
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   5
         Left            =   -72345
         TabIndex        =   30
         Top             =   1155
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   0
         Left            =   -73635
         TabIndex        =   27
         Top             =   435
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   2
         Left            =   -73635
         TabIndex        =   29
         Top             =   1155
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   47
         Left            =   -71010
         TabIndex        =   133
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "염색기 구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   1
         Left            =   -71010
         TabIndex        =   138
         Top             =   420
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "BT 접수번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   23
         Left            =   -71010
         TabIndex        =   139
         Top             =   780
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "BT 접수순번"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   9
         Left            =   -69840
         TabIndex        =   32
         Top             =   420
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         MaxLength       =   8
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   10
         Left            =   -69840
         TabIndex        =   33
         Top             =   780
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   27
         Left            =   -74910
         TabIndex        =   146
         Top             =   3675
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG Remark2"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   11
         Left            =   -73290
         TabIndex        =   47
         Top             =   3660
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
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
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   6
         Left            =   -74910
         TabIndex        =   147
         Top             =   2220
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "P/O NO."
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   5
         Left            =   -73290
         TabIndex        =   43
         Top             =   2220
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
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
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   17
         Left            =   90
         TabIndex        =   148
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가공 폭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   13
         Left            =   2070
         TabIndex        =   149
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "축율/Loss"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   4
         Left            =   3810
         TabIndex        =   159
         Top             =   90
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Text            =   "12.12"
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   3
         Left            =   3150
         TabIndex        =   22
         Top             =   90
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Text            =   "12.12"
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   18
         Left            =   6150
         TabIndex        =   150
         Top             =   90
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "필 장"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   49
         Left            =   4410
         TabIndex        =   158
         Top             =   90
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "재고폭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   56
         Left            =   -69915
         TabIndex        =   163
         Top             =   3960
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG출력 Quality"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   57
         Left            =   -74910
         TabIndex        =   164
         Top             =   4005
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG출력 Destition"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   52
         Left            =   2880
         TabIndex        =   166
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "단가기준"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   53
         Left            =   4680
         TabIndex        =   167
         Top             =   495
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "야드당g"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtBox 
         Height          =   300
         Index           =   8
         Left            =   5595
         TabIndex        =   168
         Top             =   495
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   21
         Left            =   -69930
         TabIndex        =   169
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG 주문"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   7
         Left            =   -68430
         TabIndex        =   49
         Top             =   60
         Width           =   2310
         _ExtentX        =   5609
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   20
         Left            =   -74910
         TabIndex        =   170
         Top             =   2970
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG 품명2"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   13
         Left            =   -73290
         TabIndex        =   45
         Top             =   2970
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   51
         Left            =   -69930
         TabIndex        =   171
         Top             =   390
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "TAG 주문2"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   14
         Left            =   -68430
         TabIndex        =   50
         Top             =   390
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
      Begin VB.Line Line2 
         X1              =   -30
         X2              =   7830
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         Index           =   3
         X1              =   -71100
         X2              =   -71100
         Y1              =   30
         Y2              =   4050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -71085
         X2              =   -71085
         Y1              =   30
         Y2              =   4005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         Index           =   0
         X1              =   -70020
         X2              =   -70020
         Y1              =   45
         Y2              =   4275
      End
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   3930
      Left            =   3960
      TabIndex        =   75
      Top             =   45
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   6932
      _Version        =   196610
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
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   795
         Index           =   4
         Left            =   8250
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   72
         ToolTipText     =   "자료 취소"
         Top             =   75
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   795
         Index           =   1
         Left            =   9840
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   67
         ToolTipText     =   "자료 수정"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   795
         Index           =   2
         Left            =   10635
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   68
         ToolTipText     =   "자료 삭제"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   795
         Index           =   0
         Left            =   9045
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         ToolTipText     =   "자료 추가"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   795
         Index           =   3
         Left            =   7455
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   56
         ToolTipText     =   "자료 저장"
         Top             =   75
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   60
         TabIndex        =   76
         Top             =   360
         Visible         =   0   'False
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   900
         _Version        =   196610
         BackColor       =   65535
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
      End
      Begin Threed.SSPanel pnlEdit 
         Height          =   2955
         Left            =   75
         TabIndex        =   77
         Top             =   930
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   5212
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   22
            Left            =   3300
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   420
            Width           =   1005
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   21
            Left            =   2760
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   90
            Width           =   1245
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   17
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   1155
            Width           =   2205
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   5
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   19
            Top             =   2595
            Width           =   810
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   20
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   2235
            Width           =   2205
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   19
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   1875
            Width           =   2205
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   18
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   1530
            Width           =   2205
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   14
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   75
            Width           =   2205
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   1
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   435
            Width           =   2205
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   2
            Left            =   5460
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   795
            Width           =   2205
         End
         Begin MSMask.MaskEdBox mskOrderID 
            Height          =   300
            Left            =   1335
            TabIndex        =   1
            Top             =   90
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   12648384
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "####-##-####"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   3
            Left            =   135
            TabIndex        =   78
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "관리 번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   4
            Left            =   135
            TabIndex        =   79
            Top             =   435
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Order No"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   7
            Left            =   4350
            TabIndex        =   80
            Top             =   435
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "주문 형태"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   8
            Left            =   4350
            TabIndex        =   81
            Top             =   795
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "주문 구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   9
            Left            =   135
            TabIndex        =   82
            Top             =   2235
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "접수 일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   10
            Left            =   135
            TabIndex        =   83
            Top             =   2595
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkDvlyDate 
               Caption         =   "납기 일자"
               Height          =   180
               Left            =   60
               TabIndex        =   10
               Top             =   60
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   3
            Left            =   1335
            TabIndex        =   11
            Top             =   2595
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyy년 MM월 dd일 (ddd)"
            Format          =   138477571
            CurrentDate     =   36871
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   12
            Left            =   4350
            TabIndex        =   86
            Top             =   75
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
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
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   14
            Left            =   135
            TabIndex        =   85
            Top             =   1515
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "품     명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   11
            Left            =   135
            TabIndex        =   84
            Top             =   1155
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "납품 장소"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   5
            Left            =   3960
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   1440
            Visible         =   0   'False
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196610
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
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   4
            Left            =   3300
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1550
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196610
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
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   2
            Left            =   1335
            TabIndex        =   9
            Top             =   2235
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyy년 MM월 dd일 (ddd)"
            Format          =   138477571
            CurrentDate     =   36871
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   0
            Left            =   1335
            TabIndex        =   3
            Top             =   420
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
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
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   1
            Left            =   1335
            TabIndex        =   6
            Top             =   1155
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IMEMode         =   10
         End
         Begin MRPPlus2.WizText txtCode 
            Height          =   300
            Index           =   1
            Left            =   1335
            TabIndex        =   7
            Top             =   1515
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IMEMode         =   8
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   41
            Left            =   4350
            TabIndex        =   125
            Top             =   1515
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가공료 정산"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   42
            Left            =   4350
            TabIndex        =   126
            Top             =   1875
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "세금 계산서"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   46
            Left            =   4350
            TabIndex        =   127
            Top             =   2235
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "확정 구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   19
            Left            =   4350
            TabIndex        =   128
            Top             =   2595
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "화폐 단위"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   37
            Left            =   6300
            TabIndex        =   129
            Top             =   2595
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "환율"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   5
            Left            =   135
            TabIndex        =   130
            Top             =   795
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "거 래 처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   3
            Left            =   3300
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   795
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196610
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
         Begin MRPPlus2.WizText txtCode 
            Height          =   300
            Index           =   0
            Left            =   1335
            TabIndex        =   5
            Top             =   795
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IMEMode         =   10
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   39
            Left            =   4350
            TabIndex        =   132
            Top             =   1155
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "소요량 정산"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   6
            Left            =   6915
            TabIndex        =   20
            Top             =   2595
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Text            =   "1,234.12"
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   48
            Left            =   135
            TabIndex        =   156
            Top             =   1860
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "ITEM"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtCode 
            Height          =   300
            Index           =   2
            Left            =   1335
            TabIndex        =   8
            Top             =   1860
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IMEMode         =   8
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   50
         Left            =   8910
         TabIndex        =   160
         Top             =   930
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "출고처 전화번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   12
         Left            =   8910
         TabIndex        =   161
         Top             =   1260
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IMEMode         =   8
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   795
      Left            =   45
      TabIndex        =   100
      Top             =   8415
      Width           =   1500
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   510
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   210
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
' 변경이력
' 요청ID : S_201112_태을염직_01
' 요청자 : 김대진
' 요청일자: 2011.12.09
' 요청내용 : 수주등록시 속성배열 인덱스 오류 발생
' 변경일자 : 2011.12.09
' 변경내용 : 재고폭이 선택이 안됨-틀수항목으로 오류 체크 하게함
'
' 요청 ID : S_201211_태을염직_01
' 요청내용 : 태을염직 신규 태그 추가
' 변경일자 : 2012.11.12
' 변경및처리내용 : TagArticle 16에서 22로 변경
'
' 요청 ID : S_201303_태을염직_02
' 요청내용 : 수주등록시 가공료 정산부분이 거래처 변경시 이상하게 변경되어 오류남
' 변경일자 : 2013.03.31
' 변경및처리내용 :소요량 정산방법 과 가공료 정산방법 콤보를 거래처 등록의 값과 일치 시킴

' 2019.01.08, 도지웅, S_201901_태을염직_01, 자라TAG 내용 추가
' 2019.05.16, 도지웅, S_201905_태을염직_01, 자라TAG 바코드 내용 수정
'*******************************************************************************

'*************************************************************************
Option Explicit

'----------------------------------------------------------------'
Private Const REPORTFILE = "\Report\Order.rpt"

'----------------------------------------------------------------'
Private m_nBaseX As Long
Private m_nBaseY As Long
Private m_nBaseBlank As Long

'----------------------------------------------------------------'
Private m_iFlag    As Integer   ' 현재 상태 (추가/수정/삭제/검색)
Private m_bloading As Boolean

Public Sub LoadOrder(ByVal OrderID As String)
    Me.Show
    chkSearch(0).Value = False
    chkSearch(3).Value = 1
    txtSearch(3).Text = OrderID
    Call FillGridOrder
    
End Sub

Private Sub cboName_Click(Index As Integer)
    ' 가공전폭 >>> 가공후폭 설정
    Select Case Index
''        Case 3
''            If cboName(3).ListCount = cboName(4).ListCount Then
''                cboName(4).ListIndex = cboName(3).ListIndex
''            End If
        Case 4, 14
            Call GetUnitPrice
        
        Case 5
            If cboName(5).ListIndex = 0 Then
                grdColor.ColFormat(5) = "#,###"
            ElseIf cboName(5).ListIndex = 1 Then
                grdColor.ColFormat(5) = "#,###.00"
            End If
        Case 6
            cboName(16).ListIndex = cboName(6).ListIndex
        Case 11
            If cboName(Index).ListIndex = 1 Then
                txtName(11).Text = "MADE IN KOREA"
            Else
                txtName(11).Text = ""
            End If
        Case 23
            If cboName(23).ListIndex = 0 Then
                grdPattern.Rows = grdPattern.FixedRows
            Else
                Call SetPatternSub
                
            End If
    End Select
End Sub

Sub SetComboPattern()
    Dim oPattern As PlusLib2.CPattern
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oPattern = New PlusLib2.CPattern
    oPattern.Connection = g_adoCon
    
    Set rs = oPattern.GetPattern
    Set oPattern = Nothing
    
    With cboName(23)
        .Clear
        .AddItem "00. 패턴미지정"
        
        Do Until rs.EOF
            .AddItem rs!PatternID & "." & CheckNull(rs!Pattern)
            rs.MoveNext
        Loop
            
    End With
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmOrder.SetComboPattern", Err.Description)
    Err.Clear
    Set rs = Nothing
    Set oPattern = Nothing

End Sub
Sub SetPatternSub()
    Dim oPattern As PlusLib2.CPattern
    Dim rs As ADODB.Recordset
    Dim iLoop%, i%
    Dim sProcess$
    
    On Error GoTo ErrHandler
    
    Set oPattern = New PlusLib2.CPattern
    oPattern.Connection = g_adoCon

    Set rs = oPattern.GetPatternSub(Left(cboName(23), 2))
    Set oPattern = Nothing

    With grdPattern
        .Redraw = flexRDNone
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem (i)
        Next i
        .Redraw = flexRDDirect
    End With
    
    With grdPattern
        Do Until rs.EOF
            
            .AddItem CStr(grdPattern.Rows) & vbTab & CheckNull(rs!Process) & vbTab & CheckNull(rs!ProcessID)
            
            rs.MoveNext
        Loop
        .Redraw = flexRDDirect
    End With
    
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrHandler:
    'MsgBox "[" & Err.Number & "]" & ":" & Err.Description, vbCritical
    Call ErrorBox(Err.Number, "frmPatternCode.ShowData", Err.Description)
    Set rs = Nothing
    Set oPattern = Nothing

End Sub

'S_201901_태을염직_01 에 의한 추가 : 자라Tag 내용 추가(단가기준 단위)
Private Sub cboUnitPriceClss_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call NextFocus
    End If
End Sub

Private Sub chkDvlyDate_Click()
    If chkDvlyDate.Value Then
        dtpDate(3).Enabled = True
    Else
        dtpDate(3).Enabled = False
    End If
End Sub

Private Sub chkDvlyDate_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
    
    If cmdPrint.Enabled = True Then
        grdColor.ColHidden(5) = False
    End If

End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15700, 9800   'S_201901_태을염직_01 에 의한 수정 : 14610 -> 16500
    PlusMDI.pnlMenu.Visible = False
    
    m_iFlag = -1
    
    Call SetComboBox
    Call InitGrid
    Call SetOperate(Me)
    Call NonEditMode(True)

    dtpDate(0) = Date
    dtpDate(1) = Date
    dtpDate(2) = Date
    dtpDate(3) = Date
    chkDvlyDate.Value = vbUnchecked
    
    txtBox(4).Locked = True
    For i = 0 To txtBox.Count - 1
        txtBox(i).Alignment = vbRightJustify
    Next i
    
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False

    For i = 1 To cmdFind.Count
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
    cmdFind(1).Enabled = False
    cmdFind(2).Enabled = False

    cboSearch(0).Enabled = False
    cboSearch(1).Enabled = False
    
    pnlName(3).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(4).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(5).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(14).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(17).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(13).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(12).Picture = LoadResPicture("BASIC", vbResIcon)
    
        
    'S_201112_태을염직_01 에 의한 추가
    pnlName(18).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlName(49).Picture = LoadResPicture("BASIC", vbResIcon)
    
    With cboSearch(2)
        .AddItem "0. 전체"
        .AddItem "1. 지정"
        .AddItem "2. 미지정"
        .ListIndex = 0
    End With
    
    
    Call SetComboPattern
    cboName(23).ListIndex = 0
    
End Sub

Private Sub cboName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call NextFocus
    End If
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then '[0] 수주일자 선택
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    ElseIf Index >= 1 And Index <= 3 Then '[1, 2] 거래처, 관리번호 선택
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
'            txtSearch(Index).SetFocus
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    ElseIf Index >= 4 And Index <= 5 Then
        If chkSearch(Index).Value = vbChecked Then
            cboSearch(Index - 4).Enabled = True
        Else
            cboSearch(Index - 4).Enabled = False
        End If
    End If
End Sub

Private Sub cmdErase_Click()

    With grdColor
        If .Rows = 1 Or .Row < 1 Then
            MsgBox LoadResString(200), vbInformation
        Else
            If Len(.TextMatrix(.Row, 4)) > 0 Then
                If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                    If .TextMatrix(.Row, 6) = "A" Then
                        .RemoveItem .Row
                    Else
                        .TextMatrix(.Row, 6) = "D"
                        .RowHidden(.Row) = True
                        
                        If .Row = .Rows - 1 Then
                            .Row = .Row - 1
                        Else
                            .Row = .Row + 1
                        End If
                    End If
                End If
            Else
                If .TextMatrix(.Row, 6) = "A" Then
                    .RemoveItem .Row
                Else
                    .TextMatrix(.Row, 6) = "D"
                    .RowHidden(.Row) = True
                    
                    If .Row = .Rows - 1 Then
                        .Row = .Rows - 1
                    Else
                        .Row = .Rows + 1
                    End If
                End If
            End If
            
            Call CalcOrderQty
        End If
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub NonEditMode(NewValue As Boolean)
    Dim i%

    If NewValue Then '[1] 조회모드 = Truw
        grdColor.Editable = flexEDNone
    Else '[2] 편집모드 = False
        grdColor.Editable = flexEDKbdMouse
    End If
    
    grdOrder.Enabled = NewValue
    cmdPlus.Enabled = Not NewValue
    cmdErase.Enabled = Not NewValue
    For i = 1 To cboName.Count - 1
        cboName(i).Locked = NewValue
    Next i
    
    For i = 0 To txtName.Count - 1
        txtName(i).Locked = NewValue
    Next i
    
    For i = 0 To txtBox.Count - 1
        txtBox(i).Locked = NewValue
    Next i
    txtRemark.Locked = NewValue
    txtEndMark.Locked = NewValue
    txtTag.Locked = NewValue
    
    '---------------------------------------------
    'S_201901_태을염직_01 에 의한 추가 : 자라Tag 추가
    cboUnitPriceClss.Locked = NewValue
    txtTagDestition.Locked = NewValue
    txtTagQuality.Locked = NewValue
    '---------------------------------------------
    
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then           '[3] 거래처 코드
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then           '[4] 출고처 코드
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    ElseIf Index = 3 Then               '[1] 거래처 코드
        Call ReturnRef(LG_CUSTOM, , False, txtCode(0))
    ElseIf Index = 4 Then           '[2] 품명1 코드
        Call ReturnRef(LG_ARTICLE, , False, txtCode(1))
    End If
End Sub

'********************************************************
'* Date : 2001-06-21 (THU)
'*
'* Description: Operate Button의 Index 상수
'*
'********************************************************
Private Sub cmdOperate_Click(Index As Integer)
    Dim nFlag%
    
    On Error GoTo ErrHandler
    
    Select Case Index
        '-------------------------------------------------------------------------------------'
        Case ID_ADDNEW
            m_iFlag = ID_ADDNEW

            Call ClearData(1)
            Call ChangeMode(Me, False)
            Call NonEditMode(False)
            
            fraSearch.Enabled = False
            pnlMsg.Caption = LoadResString(302)
            mskOrderID.Enabled = True
            grdColor.SelectionMode = flexSelectionFree
            txtName(0).SetFocus
            cboName(13).ListIndex = 2
            cboName(14).ListIndex = 1
            tabOrder.Tab = 0
            
            mskOrderID = Left(MakeDate(DF_SHORT, Now), 6)
            mskOrderID.SetFocus
            mskOrderID.SelStart = 8
        '-------------------------------------------------------------------------------------'
        Case ID_UPDATE
            If grdOrder.Rows = grdOrder.FixedRows Then
                MsgBox LoadResString(111), vbInformation
                cmdSearch.SetFocus
                Exit Sub
            End If
            m_iFlag = ID_UPDATE
            
            Call ChangeMode(Me, False)
            Call NonEditMode(False)
            
            fraSearch.Enabled = False
            pnlMsg.Caption = LoadResString(303)
            mskOrderID.Enabled = False
            grdColor.SelectionMode = flexSelectionFree
            txtName(1).SetFocus
        '-------------------------------------------------------------------------------------'
        Case ID_DELETE
            If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub
            
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo) = vbYes Then
                If DeleteData() Then
                    Call FillGridOrder
                End If
            End If
        '-------------------------------------------------------------------------------------'
        Case ID_SAVE
            If Not CheckData() Then Exit Sub
            If SaveData() Then
                Call ChangeMode(Me, True)
                Call NonEditMode(True)
                Call FillGridOrder
                
                nFlag = m_iFlag
                m_iFlag = -1
                grdColor.SelectionMode = flexSelectionByRow
                fraSearch.Enabled = True
                
                If nFlag = ID_ADDNEW Then
                    If CheckStuffINOrder Then
                        frmStuffINOrder.Mode = True
                        frmStuffINOrder.Custom = txtCode(0)
                        frmStuffINOrder.CustomID = txtCode(0).Tag
                        frmStuffINOrder.Article = txtCode(1)
                        frmStuffINOrder.ArticleID = txtCode(1).Tag
                        
                        Call PlusMDI.RunForm(2130)
                    End If
                End If
            End If
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL
            m_iFlag = -1
            Call ChangeMode(Me, True)
            Call NonEditMode(True)
            Call ShowData
            grdColor.SelectionMode = flexSelectionByRow
            fraSearch.Enabled = True
    End Select

    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "Order.cmdOperate_Click", Err.Description)

End Sub

Private Sub cmdPlus_Click()
    Dim i%

    With grdColor
        .Rows = .Rows + 1
    
'        For i = .FixedRows To .Rows - 1
'            .TextMatrix(i, 0) = CStr(i)
'        Next i

        .TextMatrix(.Rows - 1, 0) = 0
        
        If IsNumeric(.TextMatrix(.Rows - 2, 1)) Then
            .TextMatrix(.Rows - 1, 1) = Format(.TextMatrix(.Rows - 2, 1) + 1, "00000")
        Else
            .TextMatrix(.Rows - 1, 1) = Format(.Rows - 1, "00000")
        End If
        
        If .Rows - 1 > .FixedRows Then
            .TextMatrix(.Rows - 1, 5) = .TextMatrix(.Rows - 2, 5)
            .TextMatrix(.Rows - 1, 7) = .TextMatrix(.Rows - 2, 7)
        End If
        
        .TextMatrix(.Rows - 1, 6) = "A"
        
        .SetFocus
        .Select .Rows - 1, 2
'        .EditCell
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim sParam() As String
   
    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
   
    Set rs = oOrder.PrintOrderDetail(IIf(chkSearch(0).Value = vbChecked, 1, 0), _
                MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
                IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
                IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0), txtSearch(3), _
                IIf(chkSearch(4).Value = vbChecked, 1, 0), cboSearch(0).ItemData(cboSearch(0).ListIndex), _
                IIf(chkSearch(5).Value = vbChecked, cboSearch(1).ListIndex + 1, 0))
                
    Set oOrder = Nothing
    
    ReDim sParam(4)
    sParam(0) = "수주 상세 현황"
    sParam(1) = CompanyName
    If dtpDate(0) = dtpDate(1) Then
            sParam(2) = "수주일자  : " & IIf(chkSearch(0), MakeDate(DF_LONG, dtpDate(0)), "")
        Else
            sParam(2) = "수주일자  : " & MakeDate(DF_LONG, dtpDate(0)) & " ~ " & MakeDate(DF_LONG, dtpDate(1))
        End If
    sParam(3) = "거 래 처   : " & IIf(chkSearch(1), txtSearch(1), "(전체)")
    sParam(4) = "품    명   : " & IIf(chkSearch(2), txtSearch(2), "(전체)")
    
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
   
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmOrder.cmdPrint_Click", Err.Description)
End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub


Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[1] 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = Date
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub InitGrid()
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 3
        .ExtendLastCol = True
        
        .RowHeight(0) = 275
        
        .TextArray(0) = "합계":         .ColWidth(0) = 710
        .TextArray(1) = "0 건":         .ColWidth(1) = 1250
        .TextArray(2) = "0 YDS"
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        
        .Redraw = flexRDDirect
    End With

    ' Set Order Grid
    Call SetVSFlexGrid(grdOrder)
    With grdOrder
        .Redraw = False
        .Cols = 6
            
        .TextArray(1) = "완료":         .ColWidth(1) = 450
        .TextArray(2) = "관리번호":     .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "Order No.":    .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처명":     .ColWidth(4) = 1720:    .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "품명":         .ColWidth(5) = 1200:    .ColAlignment(5) = flexAlignLeftCenter
        
        .ColHidden(3) = True
        .ColAlignment(1) = flexAlignCenterCenter
        
        .WordWrap = False
        .ScrollBars = flexScrollBarBoth
        .Redraw = True
    End With
    
    ' Set 색상 관리 Grid
    Call SetVSFlexGrid(grdColor)
    With grdColor
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 8

        .TextArray(0) = "색상" & vbCrLf & "순위":     .ColWidth(0) = 600:             .ColAlignment(0) = flexAlignCenterCenter '[1] 색상 순위
        .TextArray(1) = "색상":         .ColWidth(1) = 0:               .ColAlignment(1) = flexAlignCenterCenter '[1] 색상 코드
        .TextArray(2) = "색상명":       .ColWidth(2) = 2200:            .ColAlignment(2) = flexAlignLeftCenter  '[2] 색상명
        .TextArray(3) = "Design No.":   .ColWidth(3) = 1400:            .ColAlignment(3) = flexAlignLeftCenter '[3] 디자인명
        .TextArray(4) = "수주수량":     .ColWidth(4) = 1200:             .ColAlignment(4) = flexAlignRightCenter '[4] 수주 수량
        .TextArray(5) = "단 가":        .ColWidth(5) = 1200:             .ColAlignment(5) = flexAlignRightCenter '[5] 단가
        .TextArray(6) = "Flag":         .ColWidth(6) = 0
        .TextArray(7) = "P/O NO":       .ColWidth(7) = 1400:            .ColAlignment(7) = flexAlignLeftCenter '[5] pono
        
        .TextArray(1) = .TextArray(1) & vbCrLf & "코드"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0.00"
        
        
        grdColor.ColHidden(5) = True
        grdColor.ColHidden(6) = True
        

        
        .ExplorerBar = flexExNone
        .FocusRect = flexFocusSolid
        .FloodColor = RGB(255, 0, 0)
        .Redraw = flexRDDirect
    End With
    
    Call SetVSFlexGrid(grdPattern)
    With grdPattern
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 3
        
        .TextArray(1) = "공정명":   .ColWidth(1) = 1500:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "공정ID":   .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = flexRDDirect
        '.HighLight = flexHighlightAlways
    End With
    
    
End Sub
''Private Sub SetStuffWidth()
''    Dim oCode As Pluslib2.CCode
''    Dim rs    As ADODB.Recordset
''    Dim II%
''
''    On Error GoTo ErrHandler
''
''    Set oCode = New Pluslib2.CCode
''    oCode.Connection = g_adoCon
''
''    Set rs = oCode.GetStuffWidth
''    Set oCode = Nothing
''    II = 0
''    cboName(4).Clear
''    If Not rs Is Nothing Then
''        If Not rs.BOF Then
''           rs.MoveFirst
''           Do Until rs.EOF
''            cboName(4).AddItem Trim$(rs(0))
''            cboName(4).ItemData(II) = val(rs(1))
''            II = II + 1
''            rs.MoveNext
''           Loop
''        End If
''    End If
''
''    rs.Close
''    Set rs = Nothing
''
''    Exit Sub
''
''ErrHandler:
''    Set rs = Nothing
''    Set oCode = Nothing
''
''    Err.Raise Err.Number, "Start.MakeCodeCombo", Err.Description, Err.HelpFile, Err.HelpContext
''
''End Sub
Private Sub SetComboBox()
    
    Call MakeCodeCombo(cboName(3), CD_WIDTH, , False)        ' 생지폭
'    Call MakeCodeCombo(cboName(4), CD_WIDTH, , False)        ' 가공폭
    Call MakeCodeCombo(cboName(7), CD_LABEL)        ' Label 구분
    Call MakeCodeCombo(cboName(8), CD_BAND)         ' Band 구분
    Call MakeCodeCombo(cboName(13), CD_BASIS)       ' 검사 기준
    Call MakeCodeCombo(cboName(14), CD_WORK)        ' 가공 구분
    Call MakeCodeCombo(cboName(0), CD_LENGTH, , False)       ' 필장
    Call SetStuffWidth(cboName(4))
    Call SetStuffWidth(cboSubulWidth)
    
    '=============================================================================='
    
    
    ' 주문형태
    With cboName(1)
        .AddItem "1. 내수"
        .ItemData(0) = 1
        .AddItem "2. Local"
        .ItemData(1) = 3
        .AddItem "3. Driect"
        .ItemData(2) = 5
    End With
    
    ' 주문구분
    With cboName(2)
        .AddItem "1. 임가공"
        .ItemData(0) = 1
        .AddItem "2. 제직불량"
        .ItemData(1) = 3
        .AddItem "3. 가공불량"
        .ItemData(2) = 5
        .AddItem "4. 재고정리"
        .ItemData(3) = 7
        .AddItem "5. 자체판매분"
        .ItemData(4) = 8
        .AddItem "6. 시가공, Sample"
        .ItemData(5) = 9
    End With
    
    ' 화폐구분
    With cboName(5)
        .AddItem "\"
        .ItemData(0) = 0
        .AddItem "$"
        .ItemData(1) = 1
    End With
    
    ' 수량단위
    With cboName(6)
        .AddItem "YDS"
        .ItemData(0) = 0
        .AddItem "MTS"
        .ItemData(1) = 1
    End With
    
    '--------------------------------------------------------------------
    'S_201901_태을염직_01 에 의한 추가 : 자라Tag 내용 추가(단가기준 단위)
    With cboUnitPriceClss
        .AddItem "YDS"
        .ItemData(0) = 0
        .AddItem "KG"
        .ItemData(1) = 1
    End With
    '--------------------------------------------------------------------
    
    ' EndMark 구분
    With cboName(10)
        .AddItem "1.없음"
        .ItemData(0) = 0
        .AddItem "2.있음"
        .ItemData(1) = 1
    End With
    
    ' MadeKorea 구분
    With cboName(11)
        .AddItem "1.없음"
        .ItemData(0) = 0
        .AddItem "2.있음"
        .ItemData(1) = 1
    End With
    
    ' 사용면 구분
    With cboName(12)
        .AddItem "1.Surface"
        .ItemData(0) = 0
        .AddItem "2.BackSide"
        .ItemData(1) = 1
    End With
    
    ' Tag 구분
    With cboName(15)
        .AddItem "1. White"
        .ItemData(0) = 0
        .AddItem "2. Buyer"
        .ItemData(1) = 1
    End With
    
    ' 검사단위
    With cboName(16)
        .AddItem "1. YDS"
        .ItemData(0) = 0
        .AddItem "2. MTS"
        .ItemData(1) = 1
        .AddItem "3. SQY"
        .ItemData(2) = 2
        .AddItem "4. SQM"
        .ItemData(3) = 3
    End With
    
    ' 소요량 정산 구분
    'S_201303_태을염직_02 에 의한 수정 (OLD:0,1)
    With cboName(17)
        .AddItem "1. 출고량"
        .ItemData(0) = 1
        .AddItem "2. 오더량"
        .ItemData(1) = 2
    End With
    
    ' 가공료 정산 구분
    'S_201303_태을염직_02 에 의한 수정 (OLD:0,1)
    With cboName(18)
        .AddItem "1. 출고량"
        .ItemData(0) = 1
        .AddItem "2. 오더량"
        .ItemData(1) = 2
    End With
    
    ' 세금계산서 발행
    With cboName(19)
        .AddItem "1. 발행함"
        .ItemData(0) = 0
        .AddItem "2. 발행안함"
        .ItemData(1) = 1
    End With
    
    ' 수주확정 구분
    With cboName(20)
        .AddItem "1. 확정"
        .ItemData(0) = 0
        .AddItem "2. 미확정"
        .ItemData(1) = 1
    End With
    
    ' 염색기 구분
    With cboName(9)
        .AddItem "1. JIGGER"
        .ItemData(0) = 1
        .AddItem "2. RAPID"
        .ItemData(1) = 2
    End With
    
    With cboName(21)
        .AddItem "1. 면"
        .ItemData(0) = 1
        .AddItem "2. 화섬"
        .ItemData(1) = 3
    End With
    
    With cboName(22)
        .AddItem "0.비사용":         .ItemData(0) = 0
        .AddItem "1.사용":           .ItemData(1) = 1
        
        
'        .AddItem "수출"
'        .ItemData(0) = 0
'        .AddItem "내수"
'        .ItemData(1) = 1
'        .AddItem "시가공"
'        .ItemData(2) = 2
'        .AddItem "샘플"
'        .ItemData(3) = 3
    End With
    
    With cboSearch(0)
        .AddItem "1. 면":          .ItemData(0) = 1
        .AddItem "2. 화섬":        .ItemData(1) = 3
        
        .ListIndex = 0
    End With
    
    With cboSearch(1)
        .AddItem "1. 진행"
        .AddItem "2. 완료"
        
        .ListIndex = 0
    End With
End Sub

Private Sub ClearData(Index As Integer)
    Call ClearText(txtName)
    Call ClearText(txtCode)
    
    txtCode(0).Tag = ""
    txtCode(1).Tag = ""
    grdColor.Rows = grdColor.FixedRows
    
    If Index = 0 Then '[1] 완전히 Clear
        Call ClearText(txtBox)
        Call ClearCombo(cboName)
    Else '[2] 추가 버튼을 눌렀을때
        Call ClearText(txtBox, "0")
        Call ClearCombo(cboName, 0)
        cboName(1).ListIndex = 1
        cboName(12).ListIndex = 1
        txtBox(3) = ""
    End If
    txtTag = ""
    txtRemark = ""
    txtEndMark = ""
    mskOrderID = ""
    dtpDate(2) = Now
    dtpDate(3) = Now
    chkDvlyDate.Value = vbUnchecked
    txtUnitPrice.Text = ""
    
    '----------------------------------------------------
    'S_201901_태을염직_01 에 의한 추가 : 자라Tag 내용 추가
    txtTagQuality.Text = ""         'Quality
    txtTagDestition.Text = ""       'Destition
    cboUnitPriceClss.ListIndex = 0  '단가단위
    txtBox(8) = ""                  '야드당 KG 수
    '----------------------------------------------------
    
End Sub

Private Sub ShowData()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset

    Dim sOrderID As String

    On Error GoTo ErrHandler

    With grdOrder
        sOrderID = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)
    End With

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    Set rs = oOrder.GetOrderOne(sOrderID)
    Set oOrder = Nothing
    
    If rs.EOF Then
        Call ClearData(0)
    Else
        mskOrderID = sOrderID
        grdPattern.Rows = grdPattern.FixedRows
        With rs
            txtName(0) = !OrderNo       '오더번호
            txtCode(0) = !kCustom       '거래처명
            txtCode(0).Tag = !CustomID  '거래처코드
            txtName(1) = CheckNull(!DvlyPlace)  '납품 장소
            txtCode(1) = !Article           ' 품명
            txtCode(1).Tag = !ArticleID     ' 품명코드
            txtCode(2).Text = !Item
            
            dtpDate(2) = MakeDate(DF_LONG, !AcptDate)   '접수일자
            If Len(Trim(!DvlyDate)) > 0 Then
                chkDvlyDate.Value = vbChecked
                dtpDate(3) = MakeDate(DF_LONG, !DvlyDate)   '납기일자
            Else
                chkDvlyDate.Value = vbUnchecked
            End If
            
            cboName(14).ListIndex = FindComboBox(cboName(14), CLng(!WorkID))     ' 가공구분
            cboName(0) = CheckNull(!CutQty) '필장
            cboName(1).ListIndex = FindComboBox(cboName(1), CLng(!OrderForm))   '주문형태
            cboName(2).ListIndex = FindComboBox(cboName(2), CLng(!OrderClss))   '주문구분
            cboName(3).ListIndex = FindComboBox(cboName(3), CLng(!StuffWidth))  '생지폭
            cboName(4).ListIndex = FindComboBox(cboName(4), CLng(!WorkWidth))    '가공폭
            
            txtBox(0) = CStr(!StuffWeight)      '생지중량
            txtBox(1) = CheckNull(!WorkWeight) '가공중량
            txtBox(2) = CheckNull(!WorkDensity) '가공밀도
            txtBox(3) = CheckNull(!ChunkRate)    '축율
            txtBox(4) = CheckNull(!LossRate)    '로스율
            txtBox(5) = CheckNull(!ReduceRate)  '감량율
            txtBox(6) = CStr(!ExchRate)        '환율
'            txtBox(6) = CStr(!UnitCost)        '단가
            txtBox(7) = Format(!OrderQty, "#,##0")  '오더수량
            
            cboName(5).ListIndex = CInt(!Priceclss) '화폐구분
            cboName(6).ListIndex = CInt(!UnitClss) '수량단위
            cboName(7).ListIndex = FindComboBox(cboName(7), CLng(!LabelID)) '라벨구분
            cboName(8).ListIndex = FindComboBox(cboName(8), CLng(!BandID))  'Band 구분
            cboName(10).ListIndex = FindComboBox(cboName(10), CLng(!EndClss))      'End Mark 구분
            cboName(11).ListIndex = FindComboBox(cboName(11), CLng(!MadeClss))     'Made In Korea 표기 구분
            cboName(12).ListIndex = FindComboBox(cboName(12), CLng(!SurfaceClss))     '사용면 구분
            cboName(13).ListIndex = FindComboBox(cboName(13), CLng(!BasisID))   '검사기준
            cboName(15).ListIndex = FindComboBox(cboName(15), CLng(!TagClss))   'Tag 구분
            cboName(16).ListIndex = FindComboBox(cboName(16), CLng(!BasisUnit)) '검사기준단위
            cboName(17).ListIndex = FindComboBox(cboName(17), CLng(!SpendingClss)) '소요량 정산구분
            cboName(18).ListIndex = FindComboBox(cboName(18), CLng(!workingClss)) '가공료 정산구분
            cboName(19).ListIndex = FindComboBox(cboName(19), CLng(!AccountClss)) '세금계산서 발행구분
            cboName(20).ListIndex = FindComboBox(cboName(20), CLng(!ActiveClss)) '수주확정구분
            cboName(21).ListIndex = FindComboBox(cboName(21), CLng(!ChemClss)) '원단 구분
            cboName(9).ListIndex = FindComboBox(cboName(9), CLng(!DyeingID)) '염색기구분
            cboName(22).ListIndex = FindComboBox(cboName(22), CLng(!OrderFlag)) '오더구분
            cboName(23).ListIndex = FindItem(cboName(23), Format(val(!PatternID), "0#"), 2)  '오더구분
            txtName(2) = CheckNull(!ShipClss)   'Ship Sample
            txtName(3) = CheckNull(!AdvnClss)   'Advanced Sample
            txtName(4) = CheckNull(!LotClss)    'Lot Sample
            txtName(5) = CheckNull(!PoNO)       'Po No
            txtName(6) = CheckNull(!TagArticle) ' Tag Article
            txtName(13) = CheckNull(!TagArticle2) ' Tag Article2    'S_201905_태을염직_01 에 의한 추가
            txtName(7) = CheckNull(!TagOrderNo) 'Tag OrderNo
            txtName(14) = CheckNull(!TagOrderNo2) 'Tag OrderNo2 'S_201905_태을염직_01 에 의한 추가
            txtName(8) = CheckNull(!TagRemark)  'Tag Remark
            txtName(11) = CheckNull(!TagRemark2)  'Tag Remark
            txtName(9) = CheckNull(!BTID)       ' BT ID
            txtName(10) = IIf(!BTIDSeq = 0, "", !BTIDSeq) ' BTID Seq
            txtName(12) = Trim(rs!OutTelNO)               '출고처 전화번호
                    
            txtTag = CheckNull(!Tag)            'Tag 내용
            txtEndMark = CheckNull(!EndMark)    'End mark 내용
            txtRemark = CheckNull(!Remark)      '비고사항
            If Trim(!SubulWidthID) = "" Then
                cboSubulWidth.ListIndex = 0
            Else
                cboSubulWidth.ListIndex = FindComboBox(cboSubulWidth, CLng(!SubulWidthID))  ' 재고폭
            End If
            
            If cboName(4).ItemData(cboName(4).ListIndex) <> rs!WorkWidth Then
                MsgBox ("가공 폭이 잘못 설정 되었습니다." & vbCrLf & ". 담당자에게 연락 하십시오.")
                
            End If
            
            '------------------------------------------------------------------------
            'S_201901_태을염직_01 에 의한 추가 : 자라Tag 내용 추가
            txtTagQuality.Text = CheckNull(!TagQuality)             'Quality
            txtTagDestition.Text = CheckNull(!TagDestition)         'Destition
            cboUnitPriceClss.ListIndex = CInt(CheckNum(!UnitPriceClss))   '단가단위
            txtBox(8) = Format(CheckNum(!weightperyard), "#,##0")   '야드당 KG 수
            '------------------------------------------------------------------------
            
        End With
                
        Call FillGridColor(sOrderID)
        Call SetPatternSub

    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmOrder.ShowData", Err.Description)
    
    Resume Next
End Sub

Private Sub FillGridOrder()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim lNowRow&, lNowSum&, i%
    
    On Error GoTo ErrHandler
    
    m_bloading = True
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    Set rs = oOrder.GetDraftOrder(IIf(chkSearch(0).Value = vbChecked, 1, 0), _
                MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
                IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
                IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0), txtSearch(3), _
                IIf(chkSearch(4).Value = vbChecked, 1, 0), cboSearch(0).ItemData(cboSearch(0).ListIndex), _
                IIf(chkSearch(5).Value = vbChecked, cboSearch(1).ListIndex + 1, 0), Left(cboSearch(2), 1))
    Set oOrder = Nothing
        
    With grdOrder
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            lNowRow = .Row
            .Rows = 1
        Else
            lNowRow = 1
        End If
            
        Do Until rs.EOF
            If rs!UnitClss = 0 Then
                lNowSum = lNowSum + rs!OrderQty
            Else
                lNowSum = lNowSum + Int(rs!OrderQty / 0.9144)
            End If
            
            .AddItem CStr(.Rows) & vbTab & IIf(rs!CloseClss = " ", "", "■") & vbTab & _
                    MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & rs!kCustom & vbTab & rs!Article
            i = i + 1
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    
        If .Rows > .FixedRows Then
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            .HighLight = flexHighlightAlways
            .Col = .FixedCols
            .ColSel = .Cols - 1
            If Not .RowIsVisible(.Row) Then
                .TopRow = .Row
            End If

            Call ShowData
        Else
            .HighLight = flexHighlightNever
                    
            MsgBox LoadResString(203), vbInformation + vbOKOnly
                    
            Call ClearData(0)
        End If
        .Redraw = flexRDDirect
    End With
    
    m_bloading = False
    
    grdTotal.TextArray(1) = Format(grdOrder.Rows - 1, "#,##0 건")
    grdTotal.TextArray(2) = Format(lNowSum, "#,##0 YDS")
    

        
    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing
    
    grdOrder.Redraw = flexRDDirect
    
    Call ErrorBox(Err.Number, "Order.FillGridOrder", Err.Description)
End Sub

Private Function CheckData() As Boolean
    Dim i%
    
    CheckData = True
'    If Len(txtName(0)) = 0 Then
'        MessageBox (LoadResString(232))
'
'        txtName(0).SetFocus
'        CheckData = False
'        Exit Function
'    End If

    If Len(txtCode(0).Tag) = 0 Then
        MessageBox (LoadResString(233))
        txtCode(0).SetFocus
        CheckData = False
        Exit Function
    End If
    
    If Len(txtCode(1).Tag) = 0 Then
        MessageBox (LoadResString(234))
        txtCode(1).SetFocus
        CheckData = False
        Exit Function
    End If
    
        'S_201112_태을염직_01 에 의한 추가
    '재고폭 선택 체크
    If cboSubulWidth.ListIndex < 0 Then
        MsgBox "재고폭이 선택되지 않았습니다."
        cboSubulWidth.SetFocus
        CheckData = False
        Exit Function
    End If
    
'    If Len(txtCode(2).Tag) = 0 Then
'        MessageBox (LoadResString(234))
'        txtCode(2).SetFocus
'        CheckData = False
'        Exit Function
'    End If
    
    For i = 1 To grdColor.Rows - 1
        If Len(grdColor.TextMatrix(i, 4)) = 0 Then
            MessageBox (LoadResString(235))
            CheckData = False
            Exit Function
        End If
    Next i
End Function

Private Function SaveData() As Boolean
    Dim nColorRow%, i%, iCnt%
    Dim TOrder As PlusLib2.TOrder
    Dim TOrderSub() As PlusLib2.TOrderSub
    Dim TOrderInst As PlusLib2.TOrderInst
    Dim TOrderInstDet() As PlusLib2.TOrderInstDet

    Dim oOrder As PlusLib2.COrder

    If grdColor.FixedRows = grdColor.Rows Then
        Call MessageBox(LoadResString(231))
        Exit Function
    End If

    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHandler

    With TOrder
        .sOrderID = Trim(mskOrderID)                    ' 관리번호
        .sCustomID = txtCode(0).Tag                     ' 거래처코드
        .sOrderNO = txtName(0)                          ' 오더번호
        .sPoNO = txtName(5)                             ' PONO
        .sChemClss = cboName(21).ItemData(cboName(21).ListIndex)    '원단구분
        .sOrderForm = Format(cboName(1).ItemData(cboName(1).ListIndex), "0")   ' 주문 형태
        .sOrderClss = Format(cboName(2).ItemData(cboName(2).ListIndex), "0")  ' 주문 구분
        .sAcptDate = MakeDate(DF_SHORT, dtpDate(2))    ' 접수 일자
        If chkDvlyDate.Value Then
            .sDvlyDate = MakeDate(DF_SHORT, dtpDate(3))     ' 납기 일자
        Else
            .sDvlyDate = ""
        End If
        .sArticleID = txtCode(1).Tag                    ' 품명 코드
        .sDvlyPlace = txtName(1)                        ' 출고처
        .sWorkID = Format(cboName(14).ItemData(cboName(14).ListIndex), "0000")   ' 가공 구분
        .sPriceClss = Format(cboName(5).ItemData(cboName(5).ListIndex), "0")     ' 화폐 단위
        .nExchRate = CheckNum(txtBox(6))            ' 환율
        .nOrderQty = CheckNum(txtBox(7))                '수주량
        .sUnitClss = Format(cboName(6).ItemData(cboName(6).ListIndex), "0")        ' 수량단위
        .sStuffWidth = Format(cboName(3).ItemData(cboName(3).ListIndex), "00") '생지 폭
        .nStuffWeight = CheckNum(txtBox(0))                                      '생지 중량
'        .nCutQty = IIf(Len(cboName(0).Text) > 3, Mid(cboName(0).Text, InStr(cboName(0).Text, " ") + 1), cboName(0).Text) ' 필장
        .nCutQty = cboName(0).Text ' 필장
        .sWorkWidth = Format(cboName(4).ItemData(cboName(4).ListIndex), "0#") '가공 폭
        .nWorkWeight = CheckNum(txtBox(1))                 ' 가공 중량
        .nWorkDensity = CheckNum(txtBox(2))                ' 가공 밀도
        .nChunkRate = CheckNum(txtBox(3))                ' 축율
        .nLossRate = CheckNum(txtBox(4))                'LOSS(%)
        .nReduceRate = CheckNum(txtBox(5))              '감량율
        .sTagClss = Format(cboName(15).ItemData(cboName(15).ListIndex), "0")  'Tag 종류
        .sLabelID = Format(cboName(7).ItemData(cboName(7).ListIndex), "00")  'LABEL 구분
        .sBandID = Format(cboName(8).ItemData(cboName(8).ListIndex), "00")   'BAND 구분
        .sEndClss = Format(cboName(10).ItemData(cboName(10).ListIndex), "0")          ' EndMark 구분
        .sMadeClss = Format(cboName(11).ItemData(cboName(11).ListIndex), "0")         ' MadeKorea 구분
        .sSurfaceClss = Format(cboName(12).ItemData(cboName(12).ListIndex), "0")         ' 사용면 구분
        .sShipClss = txtName(2)                          'Ship Sample
        .sAdvnClss = txtName(3)                          'Advanced Sample
        .sLotClss = txtName(4)                           'LOT Sample
        .sEndMark = txtEndMark                           'End Mark 내용
        
        'S_201211_태을염직_01 에 의한 수정 (OLD:16)
        .sTagArticle = txtName(6) 'Tag 품명
        .sTagArticle2 = txtName(13) 'Tag 품명2   'S_201905_태을염직_01 에 의한 추가
        
        'S_201211_태을염직_01 에 의한 수정(OLD:16)
        .sTagOrderNo = txtName(7) 'Tag 주문번호
        .sTagOrderNo2 = txtName(14)    'Tag 주문번호2 'S_201905_태을염직_01 에 의한 추가
        
        'S_201211_태을염직_01 에 의한 수정 (OLD:16)
        .sTagRemark = txtName(8)                        'Tag 기재사항
        .sTagRemark2 = txtName(11)                      'Tag 기재사항2
        .sBTID = txtName(9)                              'BTID
        .nBTIDSeq = CheckNum(txtName(10))                'BTID Seq
        .sTag = txtTag                                   'Tag 내용
        .sBasisID = Format(cboName(13).ItemData(cboName(13).ListIndex), "00")          '검사기준
        .sBasisUnit = cboName(16).ItemData(cboName(16).ListIndex)          '검사기준 단위
        .sSpendingClss = Format(cboName(17).ItemData(cboName(17).ListIndex), "0") '소요량 정산구분
        .sDyeingID = Format(cboName(9).ItemData(cboName(9).ListIndex), "0") '소요량 정산구분
        .sWorkingClss = Format(cboName(18).ItemData(cboName(18).ListIndex), "0") '가공료 정산구분
        .sAccountClss = Format(cboName(19).ItemData(cboName(19).ListIndex), "0") '세금계산서 발행구분
'        .sModifyClss = ""      '정정구분 정정, 취소
'        .sModifyRemark = ""    '정정사유
'        .sCancelRemark = ""    '취소사유
        .sRemark = txtRemark    ' 비고사항
        .sActiveClss = Format(cboName(20).ItemData(cboName(20).ListIndex), "0") '소요량 정산구분
        .sOrderFlag = Format(cboName(22).ItemData(cboName(22).ListIndex), "0") '소요량 정산구분
        .sPatternID = IIf(Left(cboName(23), 2) = "00", "", Left(cboName(23), 2))
        .sItem = txtCode(2).Text
        .sSubulWidthID = Format(cboSubulWidth.ItemData(cboSubulWidth.ListIndex), "0#")
        .OutTelNO = Trim(txtName(12))
'        .sCloseClss = ""    '종결구분
'        .sModifyDate = ""   '정정일자
        
        '-------------------------------------------------------------------------------------------------------
        'S_201901_태을염직_01 에 의한 추가 : 자라Tag 내용 추가
        .sTagQuality = Trim(txtTagQuality.Text) 'Quality
        .sTagDestition = Trim(txtTagDestition.Text) 'Destition
        
        Dim sUnitPriceClss          As String
        
        sUnitPriceClss = Format(cboUnitPriceClss.ItemData(cboUnitPriceClss.ListIndex), "0")
        
        If sUnitPriceClss = 1 And CheckNum(txtBox(8)) = 0 Then
            MsgBox "단위가 KG 일 경우 야드당 g이 반드시 입력되어야 합니다.", vbInformation, "[야드당g 확인]"
            txtBox(8).SetFocus
            GoTo ErrHandler
        End If
        
        .sUnitPriceClss = Format(cboUnitPriceClss.ItemData(cboUnitPriceClss.ListIndex), "0")  '단가단위
        .nWeightPerYard = CheckNum(txtBox(8))   '야드당 g 수
        '-------------------------------------------------------------------------------------------------------
        
    End With
    
    iCnt = -1
    '색상별 주문 DATA Set
    With grdColor
        nColorRow = .Rows - .FixedRows - 1
        ReDim TOrderSub(nColorRow)
        
        For i = 0 To nColorRow
            TOrderSub(i).sOrderID = mskOrderID                            ' 관리 번호
            TOrderSub(i).nOrderSeq = .ValueMatrix(.FixedRows + i, 0)                            ' 색상 순위
            TOrderSub(i).sColorID = .TextMatrix(.FixedRows + i, 1)        ' 색상 코드
            TOrderSub(i).sColor = Trim(.TextMatrix(.FixedRows + i, 2))          ' 색상명
            TOrderSub(i).sDesignNO = .TextMatrix(.FixedRows + i, 3)       ' Design No
            TOrderSub(i).nColorQty = CLng(.TextMatrix(.FixedRows + i, 4)) ' 수주 수량
            If cboName(5).ListIndex = 0 Then
                TOrderSub(i).nUnitPrice = .ValueMatrix(.FixedRows + i, 5)  ' 단가
            Else
                TOrderSub(i).nUnitPrice = .ValueMatrix(.FixedRows + i, 5)  '단가
            End If
            TOrderSub(i).sFlag = .TextMatrix(.FixedRows + i, 6)           ' 추가,수정,삭제 플래그
            TOrderSub(i).sPoNO = .TextMatrix(.FixedRows + i, 7)           ' PoNO
            
            iCnt = iCnt + 1
        Next i
        TOrder.nColorCnt = nColorRow + 1
    End With
    
    ' 생지 투입계획
    With TOrderInst
        .sInstDate = MakeDate(DF_SHORT, dtpDate(2))
        .nInstSeq = 1
        .sOrderID = Trim(mskOrderID)
    End With
    
    ' 생지 투입계획
    With grdPattern
        nColorRow = .Rows - .FixedRows
        ReDim TOrderInstDet(nColorRow)
        
        For i = 0 To nColorRow - 1
            TOrderInstDet(i).sInstDate = MakeDate(DF_SHORT, dtpDate(2))
            TOrderInstDet(i).nInstSeq = 1
            TOrderInstDet(i).nProcSeq = .TextMatrix(.FixedRows + i, 0)
            TOrderInstDet(i).sProcessID = .TextMatrix(.FixedRows + i, 2)
        Next i
    End With
    
    '-----------------------------------------------------------------------------------------
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    oOrder.UserName = g_sUserName
    
    If m_iFlag = ID_ADDNEW Then
        SaveData = oOrder.AddNewOrder(TOrder, iCnt, TOrderSub, TOrderInst, TOrderInstDet)
    ElseIf m_iFlag = ID_UPDATE Then
        TOrder.sOrderID = mskOrderID
        SaveData = oOrder.UpdateOrder(TOrder, iCnt, TOrderSub, TOrderInst, TOrderInstDet)
    End If
    
    Set oOrder = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
    '-----------------------------------------------------------------------------------------
ErrHandler:
    Screen.MousePointer = vbDefault
    Set oOrder = Nothing
    
    Call ErrorBox(Err.Number, "Order.SaveData", Err.Description)
End Function

Private Function DeleteData() As Boolean
    Dim oOrder As PlusLib2.COrder
    
    On Error GoTo ErrHandler

    DeleteData = False
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    oOrder.UserName = g_sUserName
    
    DeleteData = oOrder.DeleteOrder(mskOrderID)
    
    Set oOrder = Nothing
    Exit Function
ErrHandler:
    Set oOrder = Nothing

    Call ErrorBox(Err.Number, "Order.DeleteData", Err.Description)
    
End Function

Private Sub FillGridColor(sOrderID As String)
    Dim lNowRow&

    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    Set rs = oOrder.GetOrderSub(sOrderID)
    Set oOrder = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        With grdColor
            .Rows = .FixedRows
            .HighLight = flexHighlightNever
        End With
        Exit Sub
    End If
    
    With grdColor
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            lNowRow = .Row
            .Rows = .FixedRows
        Else
            lNowRow = 1
        End If
        
        Do Until rs.EOF
            .AddItem CStr(rs!OrderSeq) & vbTab & rs!ColorID & vbTab & _
                    rs!Color & vbTab & CheckNull(rs!DesignNO) & vbTab & _
                    rs!ColorQty & vbTab & rs!UnitPrice & vbTab & "" & vbTab & rs!PoNO
            
            rs.MoveNext
        Loop
        rs.Close
        
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


Private Sub CalcOrderQty()
    Dim i%, nSum&
    
    nSum = 0
    With grdColor
        For i = .FixedRows To .Rows - 1
            If Not IsNumeric(.TextMatrix(i, 4)) Then
                .TextMatrix(i, 4) = "0"
            Else
                If Not .RowHidden(i) Then
                    nSum = nSum + CLng(.TextMatrix(i, 4))
                End If
            End If
        Next i
    End With
    
    txtBox(7) = SetCurrency(nSum)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub grdColor_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 0 And Not m_bloading Then Exit Sub
    
    If grdColor.TextMatrix(Row, 6) <> "A" Then
        grdColor.TextMatrix(Row, 6) = "U"
    End If

    With grdColor
        Select Case Col
            Case 2, 3
                .Select Row, Col + 1
            
            Case 4, 5
                If IsNumeric(.TextMatrix(Row, Col)) Then
                    .TextMatrix(Row, Col) = .TextMatrix(Row, Col)
                    .Cell(flexcpAlignment, Row, 5) = flexAlignRightCenter
                    Call CalcOrderQty
                Else
                    .TextMatrix(Row, Col) = "0"
                End If
                
                If Col = 4 Then
                    If Row = .FixedRows Then
                        .TextMatrix(Row, 5) = txtUnitPrice        ' 거래처별 단가 적용
                    End If
                    .Select Row, Col + 1
                Else
                    If Row = .Rows - 1 Then
                        If (MsgBox(LoadResString(236), vbQuestion + vbYesNo) = vbYes) Then
                            Call cmdPlus_Click
                        Else
                            tabOrder.Tab = 1
                            txtName(5).SetFocus 'PoNO
                        End If
                    End If
                End If
            
            
        End Select
''        If Col = 4 Or Col = 5 Then
''            If IsNumeric(.TextMatrix(Row, Col)) Then
''                .TextMatrix(Row, Col) = .TextMatrix(Row, Col)
''                .Cell(flexcpAlignment, Row, 5) = flexAlignRightCenter
''                Call CalcOrderQty
''                .TextMatrix(Row, 5) = txtUnitPrice
''            Else
''                .TextMatrix(Row, Col) = "0"
''            End If
''
''            If Col = 4 Then
'''                Call GetUnitPrice
''                .Select Row, Col + 1
''            Else
''                If Row = .Rows - 1 Then
''                    If (MsgBox(LoadResString(236), vbQuestion + vbYesNo) = vbYes) Then
''                        Call cmdPlus_Click
''                    Else
''                        tabOrder.Tab = 1
''                        txtName(5).SetFocus 'PoNO
''                    End If
''                End If
''            End If
''        ElseIf Col = 2 Or Col = 3 Then
''            .Select Row, Col + 1
''        End If
    End With
End Sub


Private Sub grdOrder_AfterSort(ByVal Col As Long, Order As Integer)
    Call ShowData
End Sub

Private Sub grdOrder_DblClick()
    With grdOrder
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
        
        'Call cmdOperate_Click(ID_UPDATE)
    End With
End Sub

Private Sub grdOrder_RowColChange()
    If m_bloading Then Exit Sub
    
    Call ShowData
End Sub

Private Sub mskOrderID_KeyPress(KeyAscii As Integer)
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Set oOrder = New PlusLib2.COrder
        oOrder.Connection = g_adoCon
        
        Set rs = oOrder.GetOrderOne(mskOrderID)
        Set oOrder = Nothing
        If rs.RecordCount = 1 Then
            MsgBox "이미 같은 번호로 입력한 관리번호가 있습니다. 확인하여주십시오", vbInformation
            mskOrderID.SetFocus
            Set rs = Nothing
            Exit Sub
        End If
        Set oOrder = Nothing
        Set rs = Nothing
        Call NextFocus
    End If
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdOrder
        If optOrder(0).Value Then '[0] 관리번호
            .ColHidden(2) = True
            .ColHidden(3) = False
            chkSearch(3).Caption = "Order No."
        Else '[1] Order No.
            .ColHidden(2) = False
            .ColHidden(3) = True
            chkSearch(3).Caption = "관리번호"
        End If
    End With
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    Else
        KeyAscii = KeyPress(txtBox(Index), KeyAscii, True)
    End If
End Sub

Private Sub txtBox_LostFocus(Index As Integer)
    If Not IsNumeric(txtBox(Index)) Then txtBox(Index) = "0"

End Sub

Private Sub txtCode_Change(Index As Integer)
    If Index = 1 And m_iFlag >= 0 Then
        txtName(6) = txtCode(1)         ' 품명 >>>> Tag 품명
        Call GetUnitPrice
    End If
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then               '[1] 거래처 코드
            Call ReturnRef(LG_CUSTOM, , False, txtCode(0))
            
            If Len(txtCode(0).Tag) > 0 Then
                Call GetCustomData(txtCode(0).Tag)
            Else
                txtCode(0).SetFocus
            End If
        ElseIf Index = 1 Then           '[2] 품명 코드
            Call ReturnRef(LG_ARTICLE, , False, txtCode(1))
            
            If Len(txtCode(1).Tag) > 0 Then
                Call GetArticleData(txtCode(1).Tag)
            Else
                txtCode(1).SetFocus
            End If
        End If
    End If
End Sub
Private Sub GetUnitPrice()
    Dim oOrder As PlusLib2.COrder
    Dim tCustmUnit As PlusLib2.TCustomUnit
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    
    If Trim(txtCode(1).Tag) = "" Or Trim(txtCode(0).Tag) = "" Or cboName(4).ListIndex = -1 Or cboName(14).ListIndex = -1 Then
        txtUnitPrice.Text = 0
    
        Exit Sub
    End If
    
    tCustmUnit.sCustomID = txtCode(0).Tag
    tCustmUnit.sArticleID = txtCode(1).Tag
    tCustmUnit.sStuffWidthID = Format(cboName(4).ItemData(cboName(4).ListIndex), "00")
    tCustmUnit.sWorkID = Format(cboName(14).ItemData(cboName(14).ListIndex), "0000")
    
    Set rs = oOrder.GetCustomPrice(tCustmUnit)
    
    Set oOrder = Nothing
    
    If rs.EOF Then
        txtUnitPrice.Text = 0
        Exit Sub
    Else
        txtBox(3).Text = rs!ChunkRate
        
        txtUnitPrice.Text = rs!UnitPrice
    
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    Exit Sub
    
ErrHandler:
    Set oOrder = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmOrder.GetCustomData", Err.Description)
    
End Sub
Private Sub txtName_Change(Index As Integer)
    If Index = 0 And m_iFlag >= 0 Then
        txtName(7) = txtName(0)    ' Order NO. >>>> Tag 주문번호
    End If
End Sub


Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    If KeyAscii = vbKeyReturn And Index = 0 Then
        Set oOrder = New PlusLib2.COrder
        oOrder.Connection = g_adoCon
        Set rs = oOrder.GetExistOrder(txtName(0))
        
        If Not rs.EOF Then
            MsgBox "이미 같은 오더번호가 있습니다." & vbCrLf & "관리번호 : " & rs!OrderID & "로 동일한 오더번호를 접수하셨습니다.", vbCritical
            Call MoveFocus(KeyAscii)
    '       txtName(0).SetFocus
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
ErrHandler:
    Set oOrder = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOrder.txtName_KeyPress", Err.Description)

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
    ElseIf KeyAscii = vbKeyReturn And Index = 3 Then
        Call NextFocus
    End If
End Sub

Private Sub GetCustomData(sCustomID As String)
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    Set rs = oOrder.GetCustomData(sCustomID)
    Set oOrder = Nothing
    
    If rs.EOF Then
        Exit Sub
    End If
        
    With rs
        If Not IsNull(!SpendingClss) Then
            cboName(17).ListIndex = FindComboBox(cboName(17), CLng(!SpendingClss)) '소요량 정산구분
        Else
            cboName(17).ListIndex = 0
        End If
        If Not IsNull(!workingClss) Then
            cboName(18).ListIndex = FindComboBox(cboName(18), CLng(!workingClss)) '가공료 정산구분
        Else
            cboName(18).ListIndex = 0
        End If
    End With
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
ErrHandler:
    Set oOrder = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmOrder.GetCustomData", Err.Description)
End Sub

Private Sub GetArticleData(sArticleID As String)
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    Set rs = oOrder.GetArticleData(sArticleID)
    Set oOrder = Nothing
    
    If rs.EOF Then
        Exit Sub
    End If
        
    With rs
        cboName(9).ListIndex = FindComboBox(cboName(9), CLng(!DyeingID)) '염색기구분
    End With
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
ErrHandler:
    Set oOrder = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmOrder.GetUnitPrice", Err.Description)
End Sub



Private Function CheckStuffINOrder() As Boolean
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    
    Set rs = oStuffIn.GetStuffInNotOrder(0, "", "", 1, txtCode(0).Tag, 1, txtCode(1).Tag, 0, "", "1")
    
    CheckStuffINOrder = False
    If rs.RecordCount > 0 Then
        If MsgBox("수주가 확정안된 입고가 있습니다" & vbCrLf & "지금 확정하시겠습니까?", vbInformation + vbYesNo) = vbYes Then
            CheckStuffINOrder = True
        End If
    End If
    Exit Function
    
ErrHandler:
    Set oStuffIn = Nothing
    Set rs = Nothing
    CheckStuffINOrder = False
    Call ErrorBox(Err.Number, "frmOrder.CheckStuffINOrder", Err.Description)
End Function

