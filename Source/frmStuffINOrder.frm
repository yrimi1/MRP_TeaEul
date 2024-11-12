VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffINOrder 
   ClientHeight    =   9255
   ClientLeft      =   3825
   ClientTop       =   2700
   ClientWidth     =   11850
   Icon            =   "frmStuffINOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin TabDlg.SSTab sTAB 
      Height          =   3555
      Left            =   0
      TabIndex        =   29
      Top             =   4770
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   6271
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   600
      TabCaption(0)   =   "수주내역"
      TabPicture(0)   =   "frmStuffINOrder.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdOrder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNew"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "수주할당내역"
      TabPicture(1)   =   "frmStuffINOrder.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdAssign"
      Tab(1).Control(1)=   "cmdUpdate"
      Tab(1).Control(2)=   "cmdDel"
      Tab(1).ControlCount=   3
      Begin Threed.SSCommand cmdNew 
         Height          =   435
         Left            =   60
         TabIndex        =   43
         Top             =   3030
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   196609
         Caption         =   "수주확정"
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   495
         Left            =   -64440
         TabIndex        =   41
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "삭제"
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   495
         Left            =   -65730
         TabIndex        =   40
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "수정"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOrder 
         Height          =   2955
         Left            =   60
         TabIndex        =   30
         Top             =   60
         Width           =   11715
         _cx             =   20664
         _cy             =   5212
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
      Begin VSFlex7LCtl.VSFlexGrid grdAssign 
         Height          =   2565
         Left            =   -74940
         TabIndex        =   31
         Top             =   60
         Width           =   11715
         _cx             =   20664
         _cy             =   4524
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10110
      TabIndex        =   27
      Top             =   8520
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
      Left            =   540
      TabIndex        =   24
      Top             =   3750
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
         TabIndex        =   25
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
         TabIndex        =   26
         Top             =   120
         Width           =   270
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   3795
      Left            =   0
      TabIndex        =   23
      Top             =   960
      Width           =   11835
      _cx             =   20876
      _cy             =   6694
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
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1667
      _Version        =   196609
      Begin VB.ComboBox cboAssignClss 
         Height          =   300
         Left            =   8850
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   510
         Width           =   1485
      End
      Begin VB.ComboBox cboStuffClss 
         Height          =   300
         Left            =   8850
         Style           =   2  '드롭다운 목록
         TabIndex        =   28
         Top             =   120
         Width           =   1485
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   120
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   510
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   120
         MousePointer    =   99  '사용자 정의
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   5385
         TabIndex        =   3
         Top             =   120
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   5385
         TabIndex        =   2
         Top             =   510
         Width           =   1485
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10950
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   2100
         TabIndex        =   6
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70975489
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   2100
         TabIndex        =   7
         Top             =   510
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70975489
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   795
         TabIndex        =   8
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
            Caption         =   "입고 일자"
            Height          =   240
            Index           =   0
            Left            =   30
            TabIndex        =   9
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   3885
         TabIndex        =   10
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   11
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   6915
         TabIndex        =   12
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
         Left            =   3885
         TabIndex        =   13
         Top             =   510
         Width           =   1485
         _ExtentX        =   2619
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
            TabIndex        =   14
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   6915
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   510
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
         Left            =   7320
         TabIndex        =   16
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
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
            Caption         =   "입고구분"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   735
         Left            =   -615
         TabIndex        =   18
         Top             =   135
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1296
         _Version        =   196609
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   7320
         TabIndex        =   32
         Top             =   510
         Width           =   1485
         _ExtentX        =   2619
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
         Caption         =   "확정구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   3420
         TabIndex        =   22
         Top             =   570
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   3420
         TabIndex        =   21
         Top             =   180
         Width           =   360
      End
   End
   Begin Threed.SSFrame sfAssign 
      Height          =   825
      Left            =   30
      TabIndex        =   34
      Top             =   8400
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   1455
      _Version        =   196609
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "수주할당"
      Begin VB.TextBox txtStuffRoll 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1140
         TabIndex        =   36
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtStuffQty 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3300
         TabIndex        =   37
         Top             =   270
         Width           =   975
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Left            =   150
         TabIndex        =   35
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "절수"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   465
         Left            =   4500
         TabIndex        =   38
         Tag             =   "PERM_UPDATE"
         Top             =   270
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "저장"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   2310
         TabIndex        =   39
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "수량"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   465
         Left            =   5880
         TabIndex        =   42
         Tag             =   "PERM_UPDATE"
         Top             =   270
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         _Version        =   196609
         Caption         =   "취소"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
   End
End
Attribute VB_Name = "frmStuffINOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bLoading As Boolean

Private m_bMode As Boolean
Private m_sCustomID As String
Private m_sCustom As String
Private m_sArticleID  As String
Private m_sArticle  As String

Public Property Let Mode(bMode As Boolean)
    m_bMode = bMode
End Property

Public Property Let CustomID(sCustomID As String)
    m_sCustomID = sCustomID
End Property

Public Property Let Custom(sCustom As String)
    m_sCustom = sCustom
End Property

Public Property Let ArticleID(sArticleID As String)
    m_sArticleID = sArticleID
End Property

Public Property Let Article(sArticle As String)
    m_sArticle = sArticle
End Property


Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else
        If chkSearch(Index).Value = vbChecked Then
            If Index = 1 Or Index = 2 Then
                If Not m_bMode Then
                    txtSearch(Index).Enabled = True
                    txtSearch(Index).SetFocus
                End If
                cmdFind(Index).Enabled = True
            ElseIf Index = 3 Then
                CboStuffClss.Enabled = True
            End If
        Else
            If Index = 1 Or Index = 2 Then
                txtSearch(Index).Enabled = False
                cmdFind(Index).Enabled = False
            ElseIf Index = 3 Then
                CboStuffClss.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    sfAssign.Enabled = False
End Sub

Private Sub cmdDel_Click()
    With grdAssign
        If .Rows = .FixedRows Then
            Exit Sub
        End If
        
        If Trim(grdData.TextMatrix(grdData.Row, 15)) <> "" Then
            MsgBox ("입고시 관리번호가 설정되었습니다. " & vbCrLf & _
                    " 이 화면에서 삭제 할 수 없습니다. ")
            Exit Sub
        End If
        If MsgBox(LoadResString(201), vbQuestion + vbYesNo) = vbYes Then
            If DeleteData() Then
                Call FillGridOrder
                Call FillGridData
                Call FillGridAssign
            End If
        End If
        
    End With
        
End Sub
Private Function DeleteData() As Boolean
    Dim cStuffIN As PlusLib2.cStuffIN
    Dim tItem As PlusLib2.TAssign
    Dim sOrderID As String, nSeq As Integer, sDate As String
    
    On Error GoTo ErrHandler

    DeleteData = False
    
    Set cStuffIN = New PlusLib2.cStuffIN
    cStuffIN.Connection = g_adoCon
    cStuffIN.UserName = g_sUserName
    
    With grdAssign
        sOrderID = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)
        nSeq = .ValueMatrix(.Row, 10)
        sDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 1))
    End With
    
    With grdData
        tItem.StuffDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 3))
        tItem.StuffClss = .TextMatrix(.Row, 14)
        tItem.StuffSeq = .TextMatrix(.Row, 5)
        tItem.OrderID = sOrderID
        tItem.AssignSeq = nSeq
    End With
    
    DeleteData = cStuffIN.DeleteAssign(tItem)
    
    Set cStuffIN = Nothing
    Exit Function
ErrHandler:
    Set cStuffIN = Nothing

    Call ErrorBox(Err.Number, "frmStuffINOrder.DeleteData", Err.Description)
    
End Function

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

Private Sub cmdNew_Click()
    sfAssign.Enabled = True
    txtStuffRoll.SetFocus
End Sub

Private Sub cmdSave_Click()
    If sTAB.Tab = 0 Then
        If SaveData Then
            txtStuffRoll.Text = ""
            txtStuffQty.Text = ""
            Call FillGridData
            Call FillGridAssign
        End If
    Else
        If UpdateData Then
            txtStuffRoll.Text = ""
            txtStuffQty.Text = ""
            Call FillGridData
            Call FillGridAssign
        End If
    End If
    sfAssign.Enabled = False
    
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[1] 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = Date
    End If
    chkSearch(0).Value = vbChecked
End Sub

Private Sub cmdUpdate_Click()
    
    With grdAssign
        If .Rows = .FixedRows Then
            Exit Sub
        End If
        
        If Trim(grdData.TextMatrix(grdData.Row, 15)) <> "" Then
            MsgBox ("입고시 관리번호가 설정되었습니다. " & vbCrLf & _
                    " 할당수량을 수정 할 수 없습니다.")
            Exit Sub
        End If
        
        sfAssign.Enabled = True
        txtStuffRoll.Text = .ValueMatrix(.Row, 8)
        txtStuffQty.Text = .ValueMatrix(.Row, 9)
        
    End With
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    Call InitGrid

    For i = 0 To 1
        dtpDate(i) = Now
    Next i
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
        
    With CboStuffClss
        .AddItem "생지":        .ItemData(0) = "1"
'        .AddItem "Shortage변상분"
'        .ItemData(1) = "2"
        .AddItem "반품(생지)":   .ItemData(1) = "3"
'        .AddItem "반품 완제품(가공불량)"
'        .ItemData(3) = "4"
        
        .ListIndex = 0
        .Enabled = False
    End With
    
    With cboAssignClss
        .AddItem "1.전체"
        .AddItem "2.미할당"
        .AddItem "3.완료"
        .ListIndex = 0
    End With
        
    pnlProgress.Visible = False
    chkSearch(0).Value = vbUnchecked
    sfAssign.Enabled = False
    
    If m_bMode Then
        chkSearch(1).Value = vbChecked
        chkSearch(2).Value = vbChecked
        
        txtSearch(1) = m_sCustom
        txtSearch(1).Tag = m_sCustomID
        txtSearch(2) = m_sArticle
        txtSearch(2).Tag = m_sArticleID
        
        Call FillGridData
    End If
    
    sfAssign.Enabled = False
End Sub



Private Sub grdAssign_RowColChange()
    txtStuffRoll.Text = ""
    txtStuffQty.Text = ""
End Sub

Private Sub grdData_RowColChange()
    If m_bLoading Then Exit Sub
    Call FillGridOrder
    Call FillGridAssign
    
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
'        If optOrder(0).Value Then
'            .ColWidth(11) = 1300
'            .ColWidth(10) = 0
'        Else
'            .ColWidth(11) = 0
'            .ColWidth(10) = 300
'        End If
    End With
End Sub



Private Sub sTAB_Click(PreviousTab As Integer)
    If PreviousTab <> sTAB.Tab Then
        txtStuffRoll.Text = ""
        txtStuffQty.Text = ""
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
    End If
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 17
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = " "
        .TextArray(1) = "거래처":         .ColWidth(1) = 1750:             .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "품명":           .ColWidth(2) = 2300:             .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "입고일자":       .ColWidth(3) = 1000:             .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "입고구분":       .ColWidth(4) = 1000:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "입고순번":       .ColWidth(5) = 0:                .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "사종":           .ColWidth(6) = 0:                .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "입고" & vbCrLf & "절수":       .ColWidth(7) = 600:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "입고수량":       .ColWidth(8) = 1000:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "할당수량":       .ColWidth(9) = 1000:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "미할당" & vbCrLf & "수량":    .ColWidth(10) = 1000:            .ColAlignment(10) = flexAlignRightCenter
        
        .TextArray(11) = "입고처":        .ColWidth(11) = 1000:            .ColAlignment(11) = flexAlignLeftCenter
        .TextArray(12) = "CustomID":      .ColWidth(12) = 0
        .TextArray(13) = "ArticleID":     .ColWidth(13) = 0
        .TextArray(14) = "StuffClss":     .ColWidth(14) = 0
        .TextArray(15) = "OrderID":       .ColWidth(15) = 0
        .TextArray(16) = "OrderCnt":      .ColWidth(16) = 0
        
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
    
    With grdOrder
        .Cols = 11
        Call SetVSFlexGrid(grdOrder)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = " "
        .TextArray(1) = "거래처":       .ColWidth(1) = 0:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "품명":         .ColWidth(2) = 0:            .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "관리번호":     .ColWidth(3) = 1350:            .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "OrderNo":      .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "생지폭":       .ColWidth(5) = 1350:            .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "수주량":       .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "소요량":       .ColWidth(7) = 1000:            .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "입고절수":     .ColWidth(8) = 0:               .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "입고수량":     .ColWidth(9) = 1000:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "미입고량":    .ColWidth(10) = 1000:           .ColAlignment(10) = flexAlignRightCenter
        
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .Redraw = flexRDDirect
    End With
    
    
    With grdAssign
        .Cols = 11
        Call SetVSFlexGrid(grdAssign)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = " "
        .TextArray(1) = "확정일자":     .ColWidth(1) = 1750:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "관리번호":     .ColWidth(2) = 1350:            .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "OrderNo":      .ColWidth(3) = 2000:            .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "생지폭":       .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "수주량":       .ColWidth(5) = 1000:            .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "단위":         .ColWidth(6) = 600:             .ColAlignment(6) = flexAlignCenterCenter
        
        .TextArray(7) = "소요량":       .ColWidth(7) = 1000:            .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "확정절수":     .ColWidth(8) = 1000:            .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "확정수량":     .ColWidth(9) = 1000:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "AssignSeq":   .ColWidth(10) = 0:               .ColAlignment(10) = flexAlignRightCenter
        
        
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        
        .Redraw = flexRDDirect
    
    End With
End Sub

Private Sub FillGridData()
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim i%, sStuffClss$
    Dim iNowRow%
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    m_bLoading = True
        
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    
    Set rs = oStuffIn.GetStuffInNotOrder(IIf(chkSearch(0) = vbChecked, 1, 0) _
                                , MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)) _
                                , IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag _
                                , IIf(chkSearch(3) = vbChecked, 1, 0), CboStuffClss.ItemData(CboStuffClss.ListIndex), _
                                 val(Left(cboAssignClss, 1)))
    Set oStuffIn = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        grdData.Rows = grdData.FixedRows
        grdOrder.Rows = grdOrder.FixedRows
        pnlProgress.Visible = False
        sfAssign.Enabled = False
        
        Exit Sub
    End If
    
    With grdData
        .Redraw = flexRDNone
        
        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            Select Case rs!StuffClss
                Case "1": sStuffClss = "생지"
'                Case "2": sStuffClss = "Shortage변상분"
                Case "3": sStuffClss = "반품 생지"
'                Case "4": sStuffClss = "반품 완제품"
            End Select
        
            .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & MakeDate(DF_LONG, rs!StuffDate) & vbTab & _
                sStuffClss & vbTab & rs!StuffSeq & vbTab & "" & vbTab & rs!TotRoll & vbTab & rs!TotQty & vbTab & _
                rs!AssignQty & vbTab & rs!TotQty - rs!AssignQty & vbTab & _
                Trim(rs!Custom) & vbTab & Trim(rs!CustomID) & vbTab & rs!ArticleID & vbTab & rs!StuffClss & vbTab & rs!OrderID & vbTab & rs!OrderCnt
                
            If Trim(rs!OrderID) <> "" Then
                .TextMatrix(.Rows - 1, 0) = "■"
            End If
                        
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways

            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
'            .HighLight = flexHighlightAlways
'            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
'            .TopRow = .Row
'            .Col = .FixedCols
'            .ColSel = .Cols - 1
            
            Call FillGridOrder
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        If Not m_bMode Then
            .SetFocus
        End If
    End With
    
    m_bLoading = False
    pnlProgress.Visible = False
    
    Exit Sub

ErrHandler:
    m_bLoading = False
    pnlProgress.Visible = False
    Set oStuffIn = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmStuffInByOrder.FillGridData", Err.Description)
End Sub


Private Sub FillGridAssign()
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
            
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    
    With grdData
        Set rs = oStuffIn.GetStuffINByAssignList(MakeDate(DF_SHORT, .TextMatrix(.Row, 3)), .TextMatrix(.Row, 14), .ValueMatrix(.Row, 5))
    End With
    Set oStuffIn = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        grdAssign.Rows = grdAssign.FixedRows
        Exit Sub
    End If
    
    With grdAssign
        .Redraw = flexRDNone
        .Rows = .FixedRows
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & MakeDate(DF_LONG, rs!AssignDate) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                     rs!OrderNo & vbTab & rs!StuffWidth & vbTab & SetCurrency(rs!OrderQty) & vbTab & rs!UnitClss & vbTab & _
                     SetCurrency(rs!RealQty) & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!AssignSeq
            
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If
        
        .Redraw = flexRDDirect
        If Not m_bMode Then
            .SetFocus
        End If
        
'        sfAssign.Enabled = True
    End With
    
    Exit Sub

ErrHandler:
    Set oStuffIn = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmStuffINOrder.FillGridAssign", Err.Description)

End Sub

Private Sub FillGridOrder()
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
            
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    
    Set rs = oStuffIn.GetStuffInOrder(grdData.TextMatrix(grdData.Row, 12), grdData.TextMatrix(grdData.Row, 13))

    Set oStuffIn = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        grdOrder.Rows = grdOrder.FixedRows
        sfAssign.Enabled = False
        Exit Sub
    End If
    
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!OrderNo & vbTab & rs!StuffWidth & vbTab & SetCurrency(rs!OrderQty, 0) & vbTab & SetCurrency(rs!NeedQty, 0) & vbTab & _
                0 & vbTab & SetCurrency(rs!StuffInQty, 0) & vbTab & SetCurrency(rs!NeedQty - rs!StuffInQty, 0)
            
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
        Else
            .HighLight = flexHighlightNever
        End If
        
        .Redraw = flexRDDirect
        If Not m_bMode Then
            .SetFocus
        End If
        
  '      sfAssign.Enabled = True
    End With
    
    Exit Sub

ErrHandler:
    Set oStuffIn = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmStuffInByOrder.FillGriOrder", Err.Description)
End Sub

Private Function SaveData() As Boolean
    Dim oStuffIn    As PlusLib2.cStuffIN
    Dim tItem As PlusLib2.TAssign
    Dim i%
    
    On Error GoTo ErrHandler
    
    SaveData = False
    
    With grdData
        tItem.JobFlag = "I"
        tItem.StuffDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 3))
        tItem.StuffClss = .TextMatrix(.Row, 14)
        tItem.StuffSeq = .TextMatrix(.Row, 5)
        tItem.OrderID = MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 3), OM_REDUCE)
        tItem.AssignSeq = 0
        tItem.Qty = val(txtStuffQty)
        tItem.Roll = val(txtStuffRoll)
        tItem.AssignDate = MakeDate(DF_SHORT, Now)
        
    End With
    
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    If oStuffIn.UpdateStuffINOrder(tItem) Then
        SaveData = True
    End If
    Set oStuffIn = Nothing
    Exit Function
    
ErrHandler:
    Set oStuffIn = Nothing
    SaveData = False
    Call ErrorBox(Err.Number, "frmStuffInOrder.SaveData", Err.Description)
End Function

Private Function UpdateData() As Boolean
    Dim oStuffIn    As PlusLib2.cStuffIN
    Dim tItem As PlusLib2.TAssign
    Dim i%
    Dim sOrderID As String, nSeq As Integer, sDate As String
    
    On Error GoTo ErrHandler
    
    UpdateData = False
    With grdAssign
        sOrderID = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)
        nSeq = .ValueMatrix(.Row, 10)
        sDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 1))
    End With
    
    With grdData
        tItem.JobFlag = "U"
        tItem.StuffDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 3))
        tItem.StuffClss = .TextMatrix(.Row, 14)
        tItem.StuffSeq = .TextMatrix(.Row, 5)
        tItem.OrderID = sOrderID
        tItem.AssignSeq = nSeq
        tItem.Qty = val(txtStuffQty)
        tItem.Roll = val(txtStuffRoll)
        tItem.AssignDate = sDate
    End With
    
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    If oStuffIn.UpdateStuffINOrder(tItem) Then
        UpdateData = True
    End If
    Set oStuffIn = Nothing
    Exit Function
    
ErrHandler:
    Set oStuffIn = Nothing
    UpdateData = False
    Call ErrorBox(Err.Number, "frmStuffInOrder.SaveData", Err.Description)

End Function

Private Sub txtStuffQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtStuffRoll_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub
