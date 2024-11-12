VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmInspect 
   ClientHeight    =   9255
   ClientLeft      =   1935
   ClientTop       =   1455
   ClientWidth     =   11865
   Icon            =   "frmInspect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.PictureBox picPrint 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2685
      ScaleHeight     =   570
      ScaleWidth      =   570
      TabIndex        =   87
      Top             =   8550
      Visible         =   0   'False
      Width           =   600
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   6210
      Left            =   0
      TabIndex        =   43
      Top             =   2250
      Width           =   3525
      _cx             =   6218
      _cy             =   10954
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
      Height          =   795
      Left            =   30
      TabIndex        =   78
      Top             =   8430
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   210
         Width           =   1200
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   495
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.Frame fraPrint 
      Height          =   705
      Left            =   7200
      TabIndex        =   77
      Top             =   8475
      Width           =   1200
      Begin VB.OptionButton optPrint 
         Caption         =   "로트 별"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   47
         Top             =   420
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "전체"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   46
         Top             =   165
         Width           =   885
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8445
      TabIndex        =   48
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      집계표(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin MSCommLib.MSComm comPrint 
      Left            =   1530
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   2100
      Top             =   8790
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraSearch 
      Height          =   2265
      Left            =   0
      TabIndex        =   71
      Top             =   -60
      Width           =   3530
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1380
         TabIndex        =   89
         Top             =   1530
         Width           =   1515
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1380
         TabIndex        =   10
         Top             =   1875
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   765
         Left            =   2685
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   165
         Width           =   765
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   1185
         Width           =   1515
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   300
         Index           =   2
         Left            =   75
         MousePointer    =   99  '사용자 정의
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   825
         Width           =   510
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   300
         Index           =   1
         Left            =   75
         MousePointer    =   99  '사용자 정의
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   495
         Width           =   510
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   630
         TabIndex        =   3
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   120389633
         CurrentDate     =   36271
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   630
         TabIndex        =   5
         Top             =   825
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   120389633
         CurrentDate     =   36271
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   72
         Top             =   165
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "검사일자"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   1
            Top             =   60
            Width           =   1050
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   73
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거  래  처"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   6
            Top             =   60
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   2
         Left            =   75
         TabIndex        =   74
         Top             =   1875
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   9
            Top             =   60
            Width           =   1125
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   2940
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1185
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         Enabled         =   0   'False
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   22
         Left            =   75
         TabIndex        =   90
         Top             =   1530
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품       명"
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   91
            Top             =   60
            Width           =   1125
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   2940
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1530
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         Enabled         =   0   'False
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   0
         Left            =   1950
         TabIndex        =   76
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   1
         Left            =   1950
         TabIndex        =   75
         Top             =   555
         Width           =   360
      End
   End
   Begin Threed.SSPanel pnlRollNo 
      Height          =   8430
      Left            =   3570
      TabIndex        =   51
      Top             =   30
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14870
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlEdit 
         Height          =   7560
         Left            =   0
         TabIndex        =   53
         Top             =   825
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   13335
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtDefect 
            Height          =   300
            Left            =   8040
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "NoTouch"
            Top             =   7230
            Visible         =   0   'False
            Width           =   300
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   420
            Left            =   7080
            TabIndex        =   38
            Top             =   45
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   741
            _Version        =   196609
            Caption         =   "불량 삭제"
         End
         Begin Threed.SSCommand cmdAddNew 
            Height          =   420
            Left            =   5880
            TabIndex        =   37
            Top             =   45
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   741
            _Version        =   196609
            Caption         =   "불량 추가"
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   300
            Left            =   1395
            TabIndex        =   18
            Top             =   1530
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            Format          =   120389633
            CurrentDate     =   37082
         End
         Begin VB.ComboBox cboGrade 
            Height          =   300
            Left            =   4590
            Style           =   2  '드롭다운 목록
            TabIndex        =   32
            Top             =   6060
            Width           =   1470
         End
         Begin VB.ComboBox cboTeam 
            Height          =   300
            Left            =   1395
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   2250
            Width           =   1380
         End
         Begin VB.ComboBox cboExamNo 
            Height          =   300
            Left            =   1395
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   1170
            Width           =   1380
         End
         Begin VSFlex7LCtl.VSFlexGrid grdDefect 
            Height          =   5490
            Left            =   3270
            TabIndex        =   39
            Top             =   495
            Width           =   4980
            _cx             =   8784
            _cy             =   9684
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
            Height          =   300
            Index           =   3
            Left            =   90
            TabIndex        =   54
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "Order No"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   4
            Left            =   90
            TabIndex        =   55
            Top             =   810
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "Roll No"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   56
            Top             =   1170
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "검사호기"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   6
            Left            =   90
            TabIndex        =   57
            Top             =   1530
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "검사일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   7
            Left            =   90
            TabIndex        =   58
            Top             =   2250
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "검사조"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   8
            Left            =   90
            TabIndex        =   59
            Top             =   2610
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "검사자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   60
            Top             =   2970
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "투입원단수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   10
            Left            =   90
            TabIndex        =   61
            Top             =   3330
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "실제검사수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   11
            Left            =   90
            TabIndex        =   62
            Top             =   3690
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "조정검사수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   12
            Left            =   90
            TabIndex        =   63
            Top             =   4050
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "견본수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   13
            Left            =   90
            TabIndex        =   64
            Top             =   4410
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "보상수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   14
            Left            =   90
            TabIndex        =   65
            Top             =   4770
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "난단수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   16
            Left            =   90
            TabIndex        =   66
            Top             =   5130
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "원단중량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   17
            Left            =   3330
            TabIndex        =   67
            Top             =   6060
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "등급"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   19
            Left            =   90
            TabIndex        =   68
            Top             =   6570
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "Lot No"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   3
            Left            =   2790
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   2610
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   20
            Left            =   3330
            TabIndex        =   69
            Top             =   6405
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "대표불량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   21
            Left            =   3330
            TabIndex        =   70
            Top             =   6765
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "난단대표불량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   4
            Left            =   6090
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   6405
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   5
            Left            =   6090
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   6765
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   0
            Left            =   1395
            TabIndex        =   14
            Top             =   90
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
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
            BackColor       =   16777152
         End
         Begin VB.PictureBox Picture1 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   79
            Top             =   0
            Width           =   0
         End
         Begin VB.PictureBox Picture2 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   80
            Top             =   0
            Width           =   0
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   2
            Left            =   1395
            TabIndex        =   16
            Top             =   810
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   3
            Left            =   1395
            TabIndex        =   21
            Top             =   2610
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   4
            Left            =   1395
            TabIndex        =   22
            Top             =   2970
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   5
            Left            =   1395
            TabIndex        =   23
            Top             =   3330
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   6
            Left            =   1395
            TabIndex        =   24
            Top             =   3690
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   7
            Left            =   1395
            TabIndex        =   25
            Top             =   4050
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   8
            Left            =   1395
            TabIndex        =   26
            Top             =   4410
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   9
            Left            =   1395
            TabIndex        =   27
            Top             =   4770
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   10
            Left            =   1395
            TabIndex        =   28
            Top             =   5130
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   14
            Left            =   1395
            TabIndex        =   31
            Top             =   6570
            Width           =   1380
            _ExtentX        =   2434
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   15
            Left            =   4560
            TabIndex        =   33
            Top             =   6405
            Width           =   1500
            _ExtentX        =   2646
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
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   16
            Left            =   4560
            TabIndex        =   35
            Top             =   6765
            Width           =   1500
            _ExtentX        =   2646
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
            Index           =   26
            Left            =   90
            TabIndex        =   81
            Top             =   5850
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "원단폭"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   11
            Left            =   1395
            TabIndex        =   29
            Top             =   5850
            Width           =   1380
            _ExtentX        =   2434
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
            Index           =   27
            Left            =   90
            TabIndex        =   82
            Top             =   6210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "원단밀도"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   12
            Left            =   1395
            TabIndex        =   30
            Top             =   6210
            Width           =   1380
            _ExtentX        =   2434
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
            Index           =   28
            Left            =   90
            TabIndex        =   83
            Top             =   450
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "색상명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   1
            Left            =   1395
            TabIndex        =   15
            Top             =   450
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
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
            BackColor       =   16777152
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   15
            Left            =   90
            TabIndex        =   88
            Top             =   1890
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "검사시간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSMask.MaskEdBox mskTime 
            Height          =   315
            Left            =   1395
            TabIndex        =   19
            Top             =   1890
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            Mask            =   "##시 ##분"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   18
            Left            =   90
            TabIndex        =   93
            Top             =   5490
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "단위당중량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtBox 
            Height          =   300
            Index           =   13
            Left            =   1395
            TabIndex        =   94
            Top             =   5490
            Width           =   1380
            _ExtentX        =   2434
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
      End
      Begin VSFlex7LCtl.VSFlexGrid grdRollNo 
         Height          =   4830
         Left            =   0
         TabIndex        =   85
         Top             =   3570
         Width           =   8295
         _cx             =   14631
         _cy             =   8520
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
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   4245
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   40
         ToolTipText     =   "자료 저장"
         Top             =   45
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   5835
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   11
         ToolTipText     =   "자료 추가"
         Top             =   45
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   7425
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   13
         ToolTipText     =   "자료 삭제"
         Top             =   45
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   6630
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   12
         ToolTipText     =   "자료 수정"
         Top             =   45
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   5040
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   41
         ToolTipText     =   "자료 취소"
         Top             =   45
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   585
         TabIndex        =   50
         Top             =   180
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   900
         _Version        =   196609
         BackColor       =   65535
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   2310
         Left            =   0
         TabIndex        =   84
         Top             =   855
         Width           =   8295
         _cx             =   14631
         _cy             =   4075
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
      Begin VSFlex7LCtl.VSFlexGrid grdColorTotal 
         Height          =   360
         Left            =   0
         TabIndex        =   86
         Top             =   3180
         Width           =   8295
         _cx             =   14631
         _cy             =   635
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
         FixedRows       =   0
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
      Left            =   10170
      TabIndex        =   49
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmInspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
' 변경이력
'-------------------------------------------------------------------------------------
' 요청ID: S_201301_태을염직_03
' 요청자 : 노과장
' 요청일자: 2013.01.22
' 요청내용 : 검사실적 조회에서 오더전체 자료 나오지 않고 검색조건에 입력된 날짜에 검색된 자료만 나오게 요청
'**********************************************************************************
 
Option Explicit

Private Const REPORTFILE1 = "\Report\Inspect.rpt"
Private Const REPORTFILE2 = "\Report\InspectByLot.rpt"
Private Const REPORTFILE3 = "\Report\InspectRollDetail.rpt"

Private Const LIMIT_ROW2 = 6
Private Const LIMIT_ROW4 = 18
Private Const LIMIT_WIDTH2 = 2340

Private m_bSortForward As Boolean

Private m_sOperate As String * 1
Private m_bloading As Boolean


Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 11970, 9660

    pnlEdit.Top = 870
    pnlEdit.Left = 30

    Call SetOperate(Me)
    Call ChangeMode(Me, True)

    dtpDate(0) = Now
    dtpDate(1) = Now

    Me.Show

    Call InitGrid

    For i = 1 To 5
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
'    cmdTag.Picture = LoadResPicture("BARCODE", vbResIcon)

    txtBox(0).Locked = True
    txtBox(1).Locked = True

    picPrint.Visible = False

    dtpDate(0).Enabled = False
    dtpDate(1).Enabled = False
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False

    With cboExamNo
        .AddItem "1 호"
        .AddItem "2 호"
        .AddItem "3 호"
        .AddItem "4 호"
        .AddItem "5 호"
        .AddItem "6 호"
        .AddItem "7 호"
        .ListIndex = 0
    End With
            
    
    chkSearch(0).Value = vbChecked

    Call MakeCodeCombo(cboGrade, CD_GRADE)
    Call MakeCodeCombo(cboTeam, CD_TEAM)
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(Index) Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True

            dtpDate(0).SetFocus
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False

            cmdSearch.SetFocus
        End If
    Else
        If chkSearch(Index) Then
            If Index = 1 Then cmdFind(1).Enabled = True
            If Index = 2 Then cmdFind(2).Enabled = True
            txtSearch(Index).Enabled = True

            txtSearch(Index).SetFocus
        Else
            If Index = 1 Then cmdFind(1).Enabled = False
            If Index = 2 Then cmdFind(2).Enabled = False
            txtSearch(Index).Enabled = False

            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
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
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, 0, False, txtSearch(2))
        End If
    Else
        KeyAscii = KeyPress(txtSearch(Index), KeyAscii)
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    ElseIf Index = 3 Then
        Call ReturnCode(LG_PERSON, , False, txtBox(3))
    ElseIf Index = 4 Then
        Call ReturnCode(LG_DEFECT, , False, txtBox(15))
    ElseIf Index = 5 Then
        Call ReturnCode(LG_DEFECT, , False, txtBox(16))
    End If
End Sub

Private Sub cmdSearch_Click()
    If Len(txtSearch(1)) = 0 Then chkSearch(1) = vbUnchecked
    If Len(txtSearch(2)) = 0 Then chkSearch(2) = vbUnchecked
    If Len(txtSearch(3)) = 0 Then chkSearch(3) = vbUnchecked
    
    Call FillGridOrder
End Sub

Private Sub grdOrder_RowColChange()
    If m_bloading Then Exit Sub

    Call FillGridColor
End Sub

Private Sub grdColor_RowColChange()
    If m_bloading Then Exit Sub
    Call FillGridRollNo
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Select Case Index
    Case ID_ADDNEW
        If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub

'        cmdTag.Visible = False
        Call ChangeMode(Me, False)
        fraPrint.Visible = False

        Call ClearData
        pnlEdit.Visible = True
        pnlMsg.Caption = LoadResString(302)
        m_sOperate = ID_ADDNEW

        txtBox(1) = grdColor.TextMatrix(grdColor.Row, 8)
        txtBox(1).Tag = grdColor.TextMatrix(grdColor.Row, 7)
        txtBox(2).Locked = False
        txtBox(2).SetFocus
    Case ID_UPDATE
        If grdRollNo.Rows = grdRollNo.FixedRows Then Exit Sub

'        cmdTag.Visible = False
        Call ChangeMode(Me, False)
        fraPrint.Visible = False

        pnlEdit.Visible = True
        pnlMsg.Caption = LoadResString(303)
        m_sOperate = ID_UPDATE

        txtBox(2).Locked = True
        cboExamNo.SetFocus
        
        Call ShowData
    Case ID_DELETE
        If grdRollNo.Rows = grdRollNo.FixedRows Then Exit Sub

        If Not QuestionBox(LoadResString(201)) Then Exit Sub

        If DeleteData() Then Call FillGridRollNo
    Case ID_SAVE
        If SaveData() Then
'            cmdTag.Visible = True
            Call ChangeMode(Me, True)
            fraPrint.Visible = True
            pnlEdit.Visible = False

            Call FillGridRollNo
        End If
    Case ID_CANCEL
'        cmdTag.Visible = True
        Call ChangeMode(Me, True)
        fraPrint.Visible = True
        pnlEdit.Visible = False
        grdRollNo.SetFocus

        If grdRollNo.Rows > grdRollNo.FixedRows Then
            Call ShowData
        Else
            Call ClearData
        End If
    End Select
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
    Select Case Index
    Case 4, 5, 6, 7, 9
        txtBox(Index) = SetCurrency(txtBox(Index), g_nPointPos)
    Case 8, 10, 11
        txtBox(Index) = SetCurrency(txtBox(Index), 1)
    Case 2, 12, 13, 14
        txtBox(Index) = SetCurrency(txtBox(Index))
    End Select
    Call GotFocusText(txtBox(Index))
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 3
        If KeyAscii = vbKeyReturn Then Call ReturnCode(LG_PERSON, , False, txtBox(Index))
    Case 4, 5, 6, 7, 9
        KeyAscii = KeyPress(txtBox(Index), KeyAscii, True, 7)
    Case 8, 10, 11
        KeyAscii = KeyPress(txtBox(Index), KeyAscii, True, 7)
    Case 2, 12, 13, 14
        KeyAscii = KeyPress(txtBox(Index), KeyAscii, True, 5)
    Case 15, 16
        If KeyAscii = vbKeyReturn Then Call ReturnCode(LG_DEFECT, , False, txtBox(Index))
    Case Else
        KeyAscii = KeyPress(txtBox(Index), KeyAscii, False)
    End Select
End Sub

Private Sub txtBox_LostFocus(Index As Integer)
    Select Case Index
    Case 4, 5, 6, 7, 9
        txtBox(Index) = SetCurrency(txtBox(Index), g_nPointPos)
    Case 8, 10, 11
        txtBox(Index) = SetCurrency(txtBox(Index), 1)
    Case 2, 12, 13, 14
        txtBox(Index) = SetCurrency(txtBox(Index))
    End Select
End Sub

Private Sub cboExamNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub dtpExamDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub mskTime_GotFocus()
    Call GotFocusText(mskTime)
End Sub

Private Sub mskTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call NextFocus
    Else
        Call MoveFocus(KeyCode)
    End If
End Sub

Private Sub cboTeam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub cboGrade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub cmdAddNew_Click()
    With grdDefect
        .SetFocus
        If .Rows = .FixedRows Then
            .AddItem CStr(1)
        Else
            .AddItem CStr(CInt(.TextMatrix(.Rows - 1, 0)) + 1)
        End If

        .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
        .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter

        .Select .Rows - 1, 1
    End With
End Sub

Private Sub cmdDelete_Click()
    With grdDefect
        If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub

        .RemoveItem .Row
    End With
End Sub

Private Sub grdDefect_Click()
    With grdDefect
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Or .MouseCol <> 2 Then Exit Sub

        Dim irow%
        irow = .MouseRow

        txtDefect = ""

        If ReturnCode(LG_DEFECT, , True, txtDefect) Then
            .TextMatrix(.Row, 1) = txtDefect
            .TextMatrix(.Row, 7) = txtDefect.Tag

            .Select .Row, 5
            .EditCell
        Else
            .TextMatrix(.Row, 1) = ""
            .TextMatrix(.Row, 7) = ""
        End If
    End With
End Sub

Private Sub grdDefect_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdDefect
        If Col >= 5 And Col <= 6 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Then .TextMatrix(Row, Col) = 0
        End If
    End With
End Sub

Private Sub grdDefect_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With grdDefect
        If KeyAscii = vbKeyReturn Then
            If Col = 1 Then
                txtDefect = .EditText

                If ReturnCode(LG_DEFECT, , False, txtDefect) Then
                    .TextMatrix(Row, 1) = txtDefect
                    .TextMatrix(Row, 7) = txtDefect.Tag

                    .Select .Row, 5
                    .EditCell
                Else
                    .TextMatrix(Row, 1) = ""
                    .TextMatrix(Row, 7) = ""
                End If
            ElseIf Col = 5 Then
                .Select Row, Col + 1
                .EditCell
            ElseIf Col = 6 Then
                If .Row = .Rows - 1 Then
                    If (MsgBox("불량을 계속 추가하시겠습니까 ?", vbQuestion + vbYesNo) = vbYes) Then
                        Call cmdAddNew_Click
                    Else
                        Call NextFocus
                    End If
                Else
                    .Select .Row + 1, 1
                End If
            End If
        End If
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        grdOrder.ColWidth(1) = 1485
        grdOrder.ColWidth(2) = 0
        chkSearch(3).Caption = "Order No"
        pnlName(3).Caption = "Order No"
    Else
        grdOrder.ColWidth(1) = 0
        grdOrder.ColWidth(2) = 1485
        chkSearch(3).Caption = "관리번호"
        pnlName(3).Caption = "관리번호"
    End If

    With grdOrder
        If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub

        txtBox(0) = .TextMatrix(.Row, Index + 1)
    End With
End Sub

'Private Sub cmdTag_Click()
'    Dim sPrint$, i%, nBuf(0) As Byte, sUnit$, sBarCode$, sBarCode1$
'    Dim vData(), vDefect()
'    Dim oInspect As PlusLib2.CInspect
'    Dim rs As ADODB.Recordset
'
'    On Error GoTo ErrHandler
'
'    If grdRollNo.Rows > grdRollNo.FixedRows Then
'        Set oInspect = New PlusLib2.CInspect
'        oInspect.Connection = g_adoCon
'        oInspect.UserName = g_sUserName
'        Set rs = oInspect.GetInspect(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE), grdColor.TextMatrix(grdColor.Row, 7), 1, grdRollNo.TextMatrix(grdRollNo.Row, 17))
'        Set oInspect = Nothing
'
'        ReDim vData(0 To 17)
'        '바코드 = 관리번호(10) + 색상번호(2) + 호기(2) + 절순위(4) + 수량(3) + 보상(2) + LotNo(5)
'        sUnit = IIf(rs!QtyUnit = "0", " Y", " M")
'            sBarCode = rs!OrderID & rs!ColorID & rs!ExamNO & Format(rs!RollID, "0000") & Format(rs!CtrlQty, "000") & Format(rs!LossQty, "00") & Format(rs!LotNo, "@@@@@")
'            sBarCode1 = rs!OrderID & " " & rs!ColorID & " " & rs!ExamNO & " " & Format(rs!RollNO, "@@@@") & " " & _
'                        Format(rs!CtrlQty, "@@@") & " " & Format(rs!LossQty, "@@") & " " & Format(rs!LotNo, "@@@@@")
'
'        vData(0) = MakeDate(DF_LONG, rs!ExamDate)
'        vData(1) = CheckNull(rs!TagArticle)
'        vData(2) = CheckNull(rs!TagOrderNo)
'        vData(3) = CheckNull(rs!TagArticle)
'        vData(4) = CheckNull(rs!TagArticle)
'        vData(5) = CheckNull(rs!Work)
'        vData(6) = CheckNull(rs!Color)
'        vData(7) = CheckNull(rs!Width)
'        vData(8) = CheckNull(rs!DesignNo)
'        vData(9) = rs!CtrlQty & sUnit
'        vData(10) = rs!LossQty
'        vData(11) = rs!LotNo
'        vData(12) = sBarCode
'        vData(13) = sBarCode1
'        vData(14) = IIf(rs!UsedClss = 0, "SURFACE", "BACKSIDE")
'        vData(15) = rs!RollNO
'        vData(16) = CheckNull(rs!Person)
'        vData(17) = "First Inspection by"
'
'        rs.Close
'        Set rs = Nothing
'
'        With grdDefect
'            For i = .FixedRows To .Rows - .FixedRows
'                If i = 21 Then Exit For
'                ReDim Preserve vDefect(.Rows - .FixedRows - 1, 2)
'                vDefect(i - 1, 0) = .TextMatrix(i, 3)
'                vDefect(i - 1, 1) = .TextMatrix(i, 4)
'                vDefect(i - 1, 2) = .TextMatrix(i, 5)
'            Next i
'        End With
'
'    '       0: TagName 1: Ypos 2: Demerit
'        If grdDefect.Rows = grdDefect.FixedRows Then
'            sPrint = MakeCleverTagPrintString(vData, "001", vDefect, , 1, 0)
'        Else
'            sPrint = MakeCleverTagPrintString(vData, "001", vDefect, , 1, 1)
'        End If
'
'        With comPrint
'            .CommPort = g_nPrintPort
'            .Settings = "9600,n,8,1"    '바코드의 설정값들과 동일하게 Setting
'            .PortOpen = True
'            .InBufferCount = 0
'
'            For i = 1 To Len(sPrint)
'                nBuf(0) = Val(AscB(Mid(sPrint, i, 1)))
'                .Output = nBuf
'
'                DoEvents
'            Next i
'
'            .PortOpen = False
'        End With
'
'        Call MessageBox("TAG 발행을 마쳤습니다.")
'    End If
'
'    Exit Sub
'
'ErrHandler:
'    Call ErrorBox(Err.Number, "frmInspect.CmdTag", Err.Description)
'End Sub

Private Sub cmdPrint_Click()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim sParam() As String
    Dim nChkDate%, sSDate$, sEDate$, nChkCustomID%, sCustomID$, nChkArticleID%, sArticleID$, nChkOrder%, sOrder$
    Dim sReportFile$

    On Error GoTo ErrHandler

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    nChkDate = IIf(chkSearch(0) = vbChecked, 1, 0)
    sSDate = MakeDate(DF_SHORT, dtpDate(0))
    sEDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkCustomID = IIf(chkSearch(1) = vbChecked, 1, 0)
    sCustomID = IIf(Len(txtSearch(1).Tag) > 0, txtSearch(1).Tag, " ")
    nChkArticleID = IIf(chkSearch(2) = vbChecked, 1, 0)
    sArticleID = IIf(Len(txtSearch(2).Tag) > 0, txtSearch(2).Tag, " ")
    nChkOrder = IIf(chkSearch(3) = vbChecked, IIf(optOrder(0), 2, 1), 0)
    sOrder = IIf(Len(txtSearch(3)) > 0, txtSearch(3), " ")

    Set rs = oInspect.PrintInspect(IIf(optPrint(0), False, True), nChkDate, sSDate, sEDate, nChkCustomID, sCustomID, nChkArticleID, sArticleID, nChkOrder, sOrder)
    Set oInspect = Nothing

    ReDim sParam(3)

    If optPrint(0) Then
        sReportFile = REPORTFILE1
        sParam(0) = "검사 집계표"
    Else
        sReportFile = REPORTFILE2
        sParam(0) = "검사 집계표"
    End If

    sParam(1) = CompanyName
    If dtpDate(0) = dtpDate(1) Then
        sParam(2) = "검사일자  : " & IIf(chkSearch(0), MakeDate(DF_LONG, dtpDate(0)), "")
    Else
        sParam(2) = "검사일자  : " & MakeDate(DF_LONG, dtpDate(0)) & " ~ " & MakeDate(DF_LONG, dtpDate(1))
    End If
    sParam(3) = "거 래 처   : " & IIf(chkSearch(1), txtSearch(1), "(전체)")

    Call PrintReport(sReportFile, rs, sParam, PlusMDI.PrintPreview)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%

    With grdOrder
        .Cols = 6
        Call SetVSFlexGrid(grdOrder)

        .Redraw = flexRDNone

        .TextArray(0) = " "
        .TextArray(1) = "Order No":     .ColWidth(1) = 0:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "관리번호":     .ColWidth(2) = 1490:       .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "거래처명":     .ColWidth(3) = 15:      .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처":       .ColWidth(4) = 0
        .TextArray(5) = "단위":         .ColWidth(5) = 0

        .Redraw = flexRDDirect
    End With

    With grdColor
        .Cols = 9
        Call SetVSFlexGrid(grdColor)

        .Redraw = flexRDNone
        .FixedCols = 0
        
        .TextArray(0) = ""
        .TextArray(1) = "":             .ColWidth(1) = 250
        .TextArray(2) = "색상명":       .ColWidth(2) = LIMIT_WIDTH2:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "수주수량":     .ColWidth(3) = 1140:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "투입수량":     .ColWidth(4) = 1140:            .ColAlignment(4) = flexAlignRightCenter:    .ColFormat(4) = GetFormat(g_nPointPos)
        .TextArray(5) = "합격절수":     .ColWidth(5) = 1140:            .ColAlignment(5) = flexAlignRightCenter:    .ColFormat(5) = GetFormat(g_nPointPos)
        .TextArray(6) = "합격수량":     .ColWidth(6) = 1140:            .ColAlignment(6) = flexAlignRightCenter:    .ColFormat(6) = GetFormat(g_nPointPos)
        .TextArray(7) = "색상번호":     .ColWidth(7) = 0
        .TextArray(8) = "색상명":       .ColWidth(8) = 0

        .GridLines = flexGridNone
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 450
        .RowHeightMin = 275


        .Redraw = flexRDDirect
    End With
    
    With grdColorTotal
        .Cols = 5
        Call SetVSFlexGrid(grdColorTotal)

        .Redraw = flexRDNone

        .FixedCols = 1
        .FixedRows = 0
        .Rows = 1

        .RowHeight(0) = 300
        .ScrollBars = flexScrollBarNone

        .TextArray(0) = "합          계":   .ColWidth(0) = LIMIT_WIDTH2 + 610:  .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = " ":                .ColWidth(1) = 1140:                .ColAlignment(1) = flexAlignRightCenter:    .ColFormat(1) = GetFormat(0)
        .TextArray(2) = " ":                .ColWidth(2) = 1140:                .ColAlignment(2) = flexAlignRightCenter:    .ColFormat(2) = GetFormat(0)
        .TextArray(3) = " ":                .ColWidth(3) = 1140:                .ColAlignment(3) = flexAlignRightCenter:    .ColFormat(3) = GetFormat(0)
        .TextArray(4) = " ":                .ColWidth(4) = 1140:                .ColAlignment(4) = flexAlignRightCenter:    .ColFormat(4) = GetFormat(0)

         .Redraw = flexRDDirect
    End With

    With grdRollNo
        .Cols = 18
        Call SetVSFlexGrid(grdRollNo)

        .Redraw = flexRDNone

        .TextArray(1) = " ":                .ColWidth(1) = 250:         .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Roll No":          .ColWidth(2) = 480:         .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "검사호기":         .ColWidth(3) = 450:         .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "LOT":              .ColWidth(4) = 450:         .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "검사일자":         .ColWidth(5) = 600:         .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "검사원":           .ColWidth(6) = 630:         .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "투입수량":         .ColWidth(7) = 500:         .ColAlignment(7) = flexAlignRightCenter:    .ColFormat(7) = GetFormat(g_nPointPos)
        .TextArray(8) = "견본수량":         .ColWidth(8) = 500:         .ColAlignment(8) = flexAlignRightCenter:    .ColFormat(8) = GetFormat(g_nPointPos)
        .TextArray(9) = "난단수량":         .ColWidth(9) = 500:         .ColAlignment(9) = flexAlignRightCenter:    .ColFormat(9) = GetFormat(g_nPointPos)
        .TextArray(10) = "검사수량":        .ColWidth(10) = 500:        .ColAlignment(10) = flexAlignRightCenter:   .ColFormat(10) = GetFormat(g_nPointPos)
        .TextArray(11) = "중량":            .ColHidden(11) = True
        .TextArray(12) = "보상수량":        .ColWidth(12) = 450:        .ColAlignment(12) = flexAlignRightCenter:   .ColFormat(12) = GetFormat(1)
        .TextArray(13) = "불량갯수":        .ColWidth(13) = 450:        .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "벌점":            .ColWidth(14) = 450:        .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "대표불량":        .ColWidth(15) = 1000:       .ColAlignment(15) = flexAlignLeftCenter
        .TextArray(16) = "등급":            .ColWidth(16) = 350:        .ColAlignment(16) = flexAlignCenterCenter
        .TextArray(17) = "RollID":          .ColWidth(17) = 0
        
        .Redraw = flexRDDirect
    End With

    With grdDefect
        .Cols = 8
        Call SetVSFlexGrid(grdDefect)

        .Redraw = flexRDNone

        .TextArray(0) = " "
        .TextArray(1) = "불량명":       .ColWidth(1) = 2500:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "불량명":       .ColWidth(2) = 300:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "영문명":       .ColWidth(3) = 0
        .TextArray(4) = "TagName":      .ColWidth(4) = 0
        .TextArray(5) = "위치":         .ColWidth(5) = 1000:             .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "벌점":         .ColWidth(6) = 800:              .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "코드":         .ColWidth(7) = 0
       
        .ColFormat(5) = "##0"
        .ColFormat(6) = "##0"

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGridOrder()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%, iNowRow%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetOrder(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1), 1, 0), txtSearch(1).Tag, _
        IIf(chkSearch(2), 1, 0), txtSearch(2).Tag, _
        IIf(chkSearch(3), IIf(optOrder(0), 2, 1), 0), IIf(optOrder(0), txtSearch(3), MakeOrderID(txtSearch(3), OM_REDUCE)), 0)
    Set oInspect = Nothing

    With grdOrder
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            DoEvents

            .AddItem CStr(i) & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!kCustom & vbTab & rs!CustomID & vbTab & rs!UnitClss

            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If

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
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        .Redraw = flexRDDirect
        .SetFocus
    End With

    m_bloading = False
    Screen.MousePointer = vbDefault

    Call FillGridColor

    Exit Sub

ErrHandler:
    m_bloading = False
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGridColor()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%, nTotal(4) As Long, nTop As Integer
    Dim nOrderSeq%
    
    If grdOrder.Rows = grdOrder.FixedRows Then
        grdColor.Rows = grdColor.FixedRows
        grdColor.HighLight = flexHighlightNever
        Call ChangeScrollColor

        grdRollNo.Rows = grdRollNo.FixedRows
        grdRollNo.HighLight = flexHighlightNever

        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    m_bloading = True

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

''    'S_201301_태을염직_03 에 의한 수정-OLD
''    Set rs = oInspect.GetOrderSub(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE))
    
    'S_201301_태을염직_03 수정(검사일자 조건 추가)-NEW
    Set rs = oInspect.GetOrderSub(IIf(chkSearch(0), 1, 0), _
                                 MakeDate(DF_SHORT, dtpDate(0)), _
                                 MakeDate(DF_SHORT, dtpDate(1)), _
                                 MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE))
    
    Set oInspect = Nothing

    nOrderSeq = -1
    With grdColor
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            If rs!OrderSeq <> nOrderSeq Then
                .AddItem CStr(i) & vbTab & "" & vbTab & rs!Color & vbTab & SetCurrency(rs!ColorQty) & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & rs!OrderSeq & vbTab & rs!Color

                Call DoFlexGridGroup(grdColor, .Rows - 1, 0)
                Call GridCollapse(nTop)
                
                nTotal(0) = nTotal(0) + IIf(IsNull(rs!ColorQty), 0, rs!ColorQty)
                nTop = .Rows - 1
            End If
            
            .AddItem "" & vbTab & "" & vbTab & CStr(CheckNull(rs!LotNo)) & vbTab & "" & vbTab & CheckNum(rs!StuffQty) & vbTab & _
                CheckNum(rs!PassRoll) & vbTab & CheckNum(rs!PassQty) & vbTab & rs!OrderSeq & vbTab & rs!Color
            
            .TextMatrix(nTop, 4) = .TextMatrix(nTop, 4) + CheckNum(rs!StuffQty)
            .TextMatrix(nTop, 5) = .TextMatrix(nTop, 5) + CheckNum(rs!PassRoll)
            .TextMatrix(nTop, 6) = .TextMatrix(nTop, 6) + CheckNum(rs!PassQty)
            

            nTotal(1) = nTotal(1) + CheckNum(rs!StuffQty)
            nTotal(2) = nTotal(2) + CheckNum(rs!PassRoll)
            nTotal(3) = nTotal(3) + CheckNum(rs!PassQty)

            nOrderSeq = rs!OrderSeq
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        Call GridCollapse(nTop)

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways

            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If

        Call ChangeScrollColor

        .Redraw = flexRDDirect
    End With

    With grdColorTotal
        .TextMatrix(0, 1) = nTotal(0)
        .TextMatrix(0, 2) = nTotal(1)
        .TextMatrix(0, 3) = nTotal(2)
        .TextMatrix(0, 4) = nTotal(3)
    End With

    Screen.MousePointer = vbArrow

    m_bloading = False
    Call FillGridRollNo

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGridRollNo()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%, iNowRow%

    If grdColor.Rows = grdColor.FixedRows Then
        grdRollNo.Rows = grdRollNo.FixedRows
        grdRollNo.HighLight = flexHighlightNever

        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    m_bloading = True
    
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

''''    'S_201301_태을염직_03 수정-OLD
''    Set rs = oInspect.GetInspect(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE), grdColor.TextMatrix(grdColor.Row, 7))

    'S_201301_태을염직_03 수정(검사일자 조건 추가)-NEW
    Set rs = oInspect.GetInspect(IIf(chkSearch(0), 1, 0), _
                                 MakeDate(DF_SHORT, dtpDate(0)), _
                                 MakeDate(DF_SHORT, dtpDate(1)), _
                                 MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE), _
                                 grdColor.TextMatrix(grdColor.Row, 7))
    
    Set oInspect = Nothing

    With grdRollNo
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            DoEvents

            .AddItem CStr(i) & vbTab & "" & vbTab & rs!RollNo & vbTab & rs!ExamNO & vbTab & rs!LotNo & vbTab & _
                Mid(rs!ExamDate, 5, 2) & "/" & Right(rs!ExamDate, 2) & vbTab & CheckNull(rs!Person) & vbTab & _
                CheckNum(rs!StuffQty) & vbTab & CheckNum(rs!SampleQty) & vbTab & CheckNum(rs!CutQty) & vbTab & _
                CheckNum(rs!CtrlQty) & vbTab & CheckNum(rs!StuffWeightUnit) & vbTab & CheckNum(rs!LossQty) & vbTab & _
                CheckNum(rs!DefectQty) & vbTab & CheckNum(rs!DefectPoint) & vbTab & CheckNull(rs!Defect) & vbTab & _
                rs!Grade & vbTab & CStr(rs!RollSeq)

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
        Else
            .HighLight = flexHighlightNever

            Call ClearData
        End If

        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbArrow
    m_bloading = False
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGridDefect()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%

    If grdRollNo.Rows = grdRollNo.FixedRows Then Exit Sub

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetInspectSub(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE), grdRollNo.TextMatrix(grdRollNo.Row, 17))
    Set oInspect = Nothing

    With grdDefect
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            DoEvents

            .AddItem CStr(i) & vbTab & rs!KDefect & vbTab & vbTab & CheckNull(rs!EDefect) & vbTab & rs!TagName & vbTab & _
                CStr(CInt(rs!yPos)) & vbTab & rs!Demerit & vbTab & rs!DefectID
                        
            .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If

        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub ClearData()
    With grdOrder
        txtBox(0) = .TextMatrix(.Row, IIf(optOrder(0), 1, 2))
        txtBox(0).Tag = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)
    End With
    With grdColor
        txtBox(1) = .TextMatrix(.Row, 1)
        txtBox(1).Tag = .TextMatrix(.Row, 7)
    End With

    txtBox(2) = "0"
    txtBox(2).Tag = "0"
    cboExamNo.ListIndex = 0
    dtpExamDate = Now
    cboTeam.ListIndex = 0
    txtBox(3) = ""
    txtBox(3).Tag = ""
    txtBox(4) = SetCurrency("0", g_nPointPos)
    txtBox(5) = SetCurrency("0", g_nPointPos)
    txtBox(5).Tag = ""
    txtBox(6) = SetCurrency("0", g_nPointPos)
    txtBox(7) = SetCurrency("0", g_nPointPos)
    txtBox(8) = SetCurrency("0", 1)
    txtBox(9) = SetCurrency("0", g_nPointPos)

    txtBox(10) = SetCurrency("0", 1)
    txtBox(11) = SetCurrency("0", 1)
    txtBox(12) = "0"
    cboGrade.ListIndex = 0
    txtBox(13) = "0"
    txtBox(14) = ""
    txtBox(15) = ""
    txtBox(15).Tag = ""
    txtBox(16) = ""
    txtBox(16).Tag = ""
    mskTime = Format(time, "HHmm")

    pnlName(20).Tag = ""
    pnlName(21).Tag = ""
    grdDefect.Rows = grdDefect.FixedRows
End Sub

Private Sub ShowData()
    Dim oInspect As PlusLib2.CInspect
    Dim rs As ADODB.Recordset
    Dim sOrderID$, nOrderSeq%, nRollID%
    
    On Error GoTo ErrHandler

    sOrderID = MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE)
    nOrderSeq = grdColor.TextMatrix(grdColor.Row, 7)
    nRollID = grdRollNo.TextMatrix(grdRollNo.Row, 17)
    
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
    oInspect.UserName = g_sUserName
    
''    'S_201301_태을염직_03 수정-OLD
''    Set rs = oInspect.GetInspect(sOrderID, nOrderSeq, 1, nRollID)
    
    'S_201301_태을염직_03 수정(검사일자 조건 추가)-NEW
    Set rs = oInspect.GetInspect(IIf(chkSearch(0), 1, 0), _
                                 MakeDate(DF_SHORT, dtpDate(0)), _
                                 MakeDate(DF_SHORT, dtpDate(1)), _
                                 sOrderID, nOrderSeq, 1, nRollID)
    With grdRollNo
        txtBox(0) = rs!OrderNo   '오더넘버
        txtBox(0).Tag = sOrderID
        txtBox(1) = rs!Color        '색상명
        txtBox(1).Tag = rs!OrderSeq
        txtBox(2) = rs!RollNo                           '절번호
        txtBox(2).Tag = rs!RollSeq                       '일련순위
        cboExamNo.ListIndex = CInt(rs!ExamNO) - 1        '검사호기
        dtpExamDate = MakeDate(DF_LONG, rs!ExamDate)      '검사일자
        mskTime = rs!ExamTime                             '검사시간
        cboTeam.ListIndex = FindComboBox(cboTeam, CLng(rs!TeamID)) '작업조
        txtBox(3) = rs!Person                            '검사원
        txtBox(3).Tag = rs!PersonID                       '검사원코드
        txtBox(4) = SetCurrency(rs!StuffQty, g_nPointPos)  '투입수량
        txtBox(5) = SetCurrency(rs!RealQty, g_nPointPos) '실제검사수량
        txtBox(5).Tag = rs!UnitClss
        txtBox(6) = SetCurrency(rs!CtrlQty, g_nPointPos)  '검사수량
        txtBox(7) = SetCurrency(rs!SampleQty, g_nPointPos)  '견본수량
        txtBox(8) = SetCurrency(rs!LossQty, 1)           '보상수량
        txtBox(9) = SetCurrency(rs!CutQty, g_nPointPos)  '난단수량
        txtBox(10) = SetCurrency(rs!StuffWeight, 1)          '원단중량
        txtBox(13) = SetCurrency(rs!StuffWeightUnit, 1)       '원단단위당중량
        txtBox(11) = SetCurrency(rs!StuffWidth, 1)          '원단폭
        txtBox(12) = rs!Density                          '원단밀도
        If Len(rs!GradeID) > 0 Then
            cboGrade.ListIndex = FindComboBox(cboGrade, CLng(rs!GradeID))
        Else
            cboGrade.ListIndex = -1
        End If
        txtBox(14) = rs!LotNo   'QC Lot
        txtBox(15) = CheckNull(rs!Defect)  '대표불량
        txtBox(15).Tag = CheckNull(rs!DefectID)
        pnlName(20).Tag = CheckNull(rs!DefectClss) '대표불량 불량종류
        txtBox(16) = CheckNull(rs!CutDefect)  '난단대표불량
        txtBox(16).Tag = CheckNull(rs!CutDefectID)
        pnlName(20).Tag = CheckNull(rs!CutDefectClss) '난단대표불량 불량종류
    End With

    Call FillGridDefect

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmInspect.ShowData", Err.Description)

    Resume Next
End Sub

Private Function CheckData() As Boolean
    Dim oInspect As PlusLib2.CInspect
    Dim i%

    CheckData = False

    On Error GoTo ErrHandler

    If m_sOperate = ID_ADDNEW Then
        Set oInspect = New PlusLib2.CInspect
        oInspect.Connection = g_adoCon

        If oInspect.GetExistRollNo(2, txtBox(0).Tag, txtBox(1).Tag, txtBox(2), Format(cboExamNo.ItemData(cboExamNo.ListIndex), "00"), Trim(txtBox(14))) Then
            Call MessageBox("'" & Trim(txtBox(2)) & "' 의 Roll No가 이미 존재 합니다. 다시 입력하십시오.")
            txtBox(1).SetFocus

            Exit Function
        End If
    End If

    If Len(txtBox(3).Tag) = 0 Then
        Call MessageBox("'검사자'를 입력하십시오.")
        txtBox(2).SetFocus

        Exit Function
    End If

    With grdDefect
        For i = .FixedRows To .Rows - 1
            If Len(.TextMatrix(i, 7)) = 0 Then
                Call MessageBox("'불량명'을 입력하십시오.")
                .SetFocus
                .Select i, 1
                Exit Function
            End If
            If Not IsNumeric(.TextMatrix(i, 5)) Then
                Call MessageBox("'수직위치'를 정확히 입력하십시오.")
                .SetFocus
                .Select i, 5
                Exit Function
            End If
            If Not IsNumeric(.TextMatrix(i, 6)) Then
                Call MessageBox("'벌점'을 정확히 입력하십시오.")
                .SetFocus
                .Select i, 8
                Exit Function
            End If
        Next i
    End With

    CheckData = True

    Exit Function

ErrHandler:
    Set oInspect = Nothing

    CheckData = False
End Function

Private Function SaveData() As Boolean
    Dim tIns      As PlusLib2.TInspect
    Dim tInsSub() As PlusLib2.TInspectSub
    Dim oInspect  As PlusLib2.CInspect
    Dim i%, nInsSub%, nDemerit!

    SaveData = False
    If Not CheckData() Then Exit Function

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    With tIns
        .OrderID = txtBox(0).Tag
        .OrderSeq = txtBox(1).Tag
        If m_sOperate = ID_UPDATE Then
            .RollSeq = CInt(txtBox(2).Tag)
            .RollNo = CInt(txtBox(2))
        End If
        .ExamNO = Format(Left(cboExamNo, 1), "00")
        .ExamDate = MakeDate(DF_SHORT, dtpExamDate)
        .ExamTime = Format(mskTime, "0000")
        .TeamID = Format(cboTeam.ItemData(cboTeam.ListIndex), "00")
        .PersonID = txtBox(3).Tag
        .StuffQty = CSng(txtBox(4))
        .RealQty = CSng(txtBox(5))
        .CtrlQty = CSng(txtBox(6))
        .SampleQty = CSng(txtBox(7))
        .LossQty = CSng(txtBox(8))
        .CutQty = CSng(txtBox(9))
        If m_sOperate = ID_ADDNEW Then
            .UnitClss = grdOrder.TextMatrix(grdOrder.Row, 5)
        Else
            .UnitClss = txtBox(5).Tag
        End If
        .StuffWeight = CSng(txtBox(10))
        .StuffWeightUnit = CInt(txtBox(13))
        .StuffWidth = CSng(txtBox(11))
        .Density = CInt(txtBox(12))
        .GradeID = cboGrade.ItemData(cboGrade.ListIndex)
        .LotNo = txtBox(14)
        .DefectQty = grdDefect.Rows - grdDefect.FixedRows
        .DefectID = IIf(Len(Trim(txtBox(15))) > 0, txtBox(15).Tag, "")
        .DefectClss = IIf(Len(Trim(txtBox(15))) > 0, pnlName(20).Tag, "")
        .CutDefectID = IIf(Len(Trim(txtBox(16))) > 0, txtBox(16).Tag, "")
        .CutDefectClss = IIf(Len(Trim(txtBox(15))) > 0, pnlName(21).Tag, "")
    End With

    nInsSub = -1
    With grdDefect
        If .Rows > .FixedRows Then
            nInsSub = .Rows - .FixedRows - 1
            ReDim tInsSub(nInsSub)

            For i = 0 To nInsSub
                tInsSub(i).OrderID = txtBox(0).Tag
                tInsSub(i).RollSeq = txtBox(2).Tag
                tInsSub(i).DefectSeq = i + 1
                tInsSub(i).DefectID = .TextMatrix(.FixedRows + i, 7)
                tInsSub(i).yPos = CInt(.TextMatrix(.FixedRows + i, 5))
                tInsSub(i).Demerit = CSng(.TextMatrix(.FixedRows + i, 6))
                
                nDemerit = nDemerit + tInsSub(i).Demerit

            Next i
        End If
    End With
    tIns.DefectPoint = nDemerit
    
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
    oInspect.UserName = g_sUserName

    If m_sOperate = ID_ADDNEW Then
        SaveData = oInspect.AddNewInspect(tIns, nInsSub, tInsSub)
    Else
        SaveData = oInspect.UpdateInspect(tIns, nInsSub, tInsSub)
    End If
    Set oInspect = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Set oInspect = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Function

Private Function DeleteData() As Boolean
    Dim oInspect As PlusLib2.CInspect

    On Error GoTo ErrHandler

    DeleteData = False

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
    oInspect.UserName = g_sUserName

    DeleteData = oInspect.DeleteInspect(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE), grdRollNo.TextMatrix(grdRollNo.Row, 17))

    Set oInspect = Nothing

    Exit Function

ErrHandler:
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Function

Private Sub ChangeScrollColor()
    With grdColor
        .ColWidth(2) = LIMIT_WIDTH2 - IIf(.Rows > LIMIT_ROW2, 240, 0)

        grdColorTotal.ColWidth(0) = .ColWidth(2) + 610
    End With
End Sub

Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, irow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(irow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(irow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, irow, 0, irow, .Cols - 1) = &H0&        '&HE0E0E0
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, irow, 0, irow, .Cols - 1) = &HE0E0E0
        End Select
    End With
End Sub

Private Sub grdColor_DblClick()
    With grdColor
        If .Row < 1 Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub

Private Sub GridCollapse(Row As Integer)
   
    With grdColor
    
        If Row >= .FixedRows Then
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub

