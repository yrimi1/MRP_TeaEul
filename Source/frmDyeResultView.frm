VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDyeResultView 
   ClientHeight    =   9255
   ClientLeft      =   300
   ClientTop       =   450
   ClientWidth     =   15150
   Icon            =   "frmDyeResultView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15150
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   975
      Left            =   13260
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   23
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   1905
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Index           =   5
      Left            =   9060
      TabIndex        =   22
      Top             =   337
      Width           =   2085
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Index           =   4
      Left            =   9060
      TabIndex        =   21
      Top             =   15
      Width           =   2085
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Index           =   3
      Left            =   9060
      TabIndex        =   20
      Top             =   660
      Width           =   2085
   End
   Begin VB.ComboBox cboSearch 
      Height          =   300
      Index           =   2
      Left            =   12465
      Style           =   2  '드롭다운 목록
      TabIndex        =   19
      Top             =   45
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   2580
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1920
         MousePointer    =   99  '사용자 정의
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   735
         Width           =   615
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "전일"
         Height          =   315
         Index           =   0
         Left            =   1920
         MousePointer    =   99  '사용자 정의
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   420
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   1
         Top             =   720
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116785152
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   405
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116785152
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   270
         Index           =   3
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "실적 일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   975
      Left            =   2580
      TabIndex        =   4
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1720
      _Version        =   196609
      Begin VB.OptionButton optProcess 
         Caption         =   "설비별"
         Height          =   375
         Index           =   0
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   90
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optProcess 
         Caption         =   "공정별"
         Height          =   405
         Index           =   1
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   510
         Width           =   1020
      End
      Begin Threed.SSPanel pnlProcess 
         Height          =   390
         Left            =   1140
         TabIndex        =   7
         Top             =   510
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   1
            Left            =   2865
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   45
            Width           =   960
         End
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   0
            Left            =   780
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   45
            Width           =   1440
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   2280
            TabIndex        =   10
            Top             =   45
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "기계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch 
               Caption         =   "기계호기"
               Height          =   180
               Index           =   1
               Left            =   75
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   315
               Width           =   1035
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   45
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "공정명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   390
         Left            =   1140
         TabIndex        =   13
         Top             =   75
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   3
            Left            =   795
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   45
            Width           =   1440
         End
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   4
            Left            =   2865
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   45
            Width           =   960
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   2280
            TabIndex        =   16
            Top             =   45
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "기계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch 
               Caption         =   "기계호기"
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   345
               Width           =   1035
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   60
            TabIndex        =   18
            Top             =   45
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "설비명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Crystal.CrystalReport CryReport 
      Left            =   14520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSCommand cmdHTML 
      Height          =   690
      Left            =   8415
      TabIndex        =   27
      Top             =   8490
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      HTML(&H)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   10095
      TabIndex        =   28
      Top             =   8490
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11790
      TabIndex        =   29
      Top             =   8490
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
      Left            =   13500
      TabIndex        =   30
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdSum 
      Height          =   360
      Left            =   -15
      TabIndex        =   31
      Top             =   7875
      Width           =   15165
      _cx             =   26749
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
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   6885
      Left            =   -15
      TabIndex        =   32
      Top             =   990
      Width           =   15165
      _cx             =   26749
      _cy             =   12144
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483637
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   14737632
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   1
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
      Begin VB.CommandButton cmdPlus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         MaskColor       =   &H8000000B&
         Style           =   1  '그래픽
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         MaskColor       =   &H8000000B&
         Style           =   1  '그래픽
         TabIndex        =   44
         Top             =   345
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   11550
      TabIndex        =   33
      Top             =   45
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "작 업 조"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "작업조"
         Height          =   180
         Index           =   2
         Left            =   45
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   60
         Width           =   840
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   5
      Left            =   7740
      TabIndex        =   35
      Top             =   660
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "관리번호"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   60
         Width           =   1125
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   7740
      TabIndex        =   37
      Top             =   15
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "작 업 조"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   4
         Left            =   105
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   60
         Width           =   960
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   3
      Left            =   11160
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   660
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
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   4
      Left            =   11160
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   15
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
      Index           =   10
      Left            =   7740
      TabIndex        =   41
      Top             =   330
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   60
         Width           =   1050
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   5
      Left            =   11160
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   330
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
   Begin VSFlex7LCtl.VSFlexGrid grdSumModi 
      Height          =   960
      Left            =   3780
      TabIndex        =   48
      Top             =   8250
      Width           =   4065
      _cx             =   7170
      _cy             =   1693
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
      ScrollBars      =   2
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
      Height          =   720
      Left            =   11550
      TabIndex        =   24
      Top             =   270
      Width           =   1695
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   465
         Width           =   1155
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   285
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin Threed.SSCommand cmdRpRate 
      Height          =   690
      Left            =   60
      TabIndex        =   50
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      평량조회"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.CheckBox chkPlan 
      BackColor       =   &H00F2E8FF&
      Caption         =   "  계획"
      Height          =   375
      Left            =   13260
      TabIndex        =   49
      Top             =   600
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox chkResult 
      BackColor       =   &H00F1F1F1&
      Caption         =   "   실적"
      Height          =   375
      Left            =   14220
      TabIndex        =   51
      Top             =   600
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmDyeResultView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE0 = "\Report\ResultWithRW.rpt"
Private Const REPORTFILE1 = "\Report\ResultWithElse.rpt"
Private Const REPORTFILE2 = "\Report\ResultWithRefine.rpt"
Private Const REPORTFILE3 = "\Report\ResultWithDry.rpt"
Private Const REPORTFILE4 = "\Report\ResultWithTenter.rpt"
Private Const REPORTFILE5 = "\Report\ResultWithReduce.rpt"
Private Const REPORTFILE6 = "\Report\ResultWithInspect.rpt"
Private Const REPORTFILE7 = "\Report\ResultWithRapid.rpt"
Private Const REPORTFILE8 = "\Report\InspectRecord.rpt"

Private Const LIMIT_ROW = 23
Private Const LIMIT_WIDTH0 = 1365
Private Const LIMIT_WIDTH1 = 1890
Private Const LIMIT_WIDTH2 = 1270
Private Const LIMIT_WIDTH3 = 1740
Private Const LIMIT_WIDTH4 = 1200
Private Const LIMIT_WIDTH5 = 1755
Private Const LIMIT_WIDTH6 = 1780
Private Const LIMIT_WIDTH7 = 1740

Private m_iSortType As Integer
Private m_bloading  As Boolean
Private m_bSkip As Boolean


Private Sub cmdMinus_Click()
Dim iCount As Integer

    If grdData.Rows <= grdData.FixedRows Then
        Exit Sub
    End If
    
    With grdData
        For iCount = 2 To .Rows - 1
            If .IsSubtotal(iCount) Then
                .IsCollapsed(iCount) = flexOutlineCollapsed
            End If
        Next iCount
    End With

End Sub

Private Sub cmdPlus_Click()
Dim iCount As Integer

    If grdData.Rows <= grdData.FixedRows Then
        Exit Sub
    End If
    
    With grdData
        For iCount = 2 To .Rows - 1
            If .IsSubtotal(iCount) Then
                .IsCollapsed(iCount) = flexOutlineExpanded
            End If
        Next iCount
    End With

End Sub

Private Sub cmdPrint_Click()
Dim sMsg As String

    With grdData
        .Redraw = False
    
        .ExtendLastCol = True
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(.Rows - 1) = False
        
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .FontSize = 7
        
        If chkPlan.Value = 1 Then
            If chkResult.Value = 1 Then
                sMsg = "(실적 / 계획)"
            Else
                sMsg = "(계획)"
            End If
        Else
            If chkResult.Value = 1 Then
                sMsg = "(실적)"
            Else
                sMsg = ""
            End If
        End If
        
        .Cell(flexcpText, 0, 5, 0, .Cols - 1) = "염 색  일 지" & sMsg
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 1, 0, .Cols - 1) = True
        .Cell(flexcpText, 1, 4, 1, 23) = "▶ 실적일 : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD") & _
                                            "  [" & IIf(optProcess(0).Value = True, cboSearch(4).Text, cboSearch(1).Text) & "]"
        .Cell(flexcpAlignment, 1, 4, 1, 23) = flexAlignLeftCenter
        .Cell(flexcpText, 1, 31, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpAlignment, 1, 31, 1, .Cols - 1) = flexAlignRightCenter
        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite
        
        .ColWidth(4) = 500
        .ColWidth(6) = 400
        .ColWidth(11) = 1400
        .ColWidth(12) = 1000
        .ColWidth(14) = 800
        .ColWidth(15) = 1800
        .ColWidth(16) = 1300
        .ColWidth(21) = 1200
        .ColWidth(22) = 400
        .ColWidth(23) = 600
        .ColWidth(31) = 500
        .ColWidth(35) = 400
        .ColWidth(36) = 600
        .ColWidth(37) = 600
        .ColWidth(38) = 400
        .ColWidth(39) = 400
        .ColWidth(40) = 0
        
        .PrintGrid "태을염직", True, 2, 150, 500

        .ExtendLastCol = False
'        .GridLinesFixed = flexGridInset
'        .GridColorFixed = &H80000010
        .FixedAlignment(4) = flexAlignCenterCenter
        .FixedAlignment(31) = flexAlignCenterCenter
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(.Rows - 1) = True

        .FontSize = 9

        .ColWidth(4) = 600
        .ColWidth(6) = 500
        .ColWidth(11) = 1400
        .ColWidth(12) = 1000
        .ColWidth(14) = 1000
        .ColWidth(15) = 2000
        .ColWidth(16) = 1400
        .ColWidth(21) = 1300
        .ColWidth(22) = 500
        .ColWidth(23) = 700
        .ColWidth(31) = 600
        .ColWidth(32) = 500
        .ColWidth(33) = 500
        .ColWidth(35) = 500
        .ColWidth(36) = 700
        .ColWidth(37) = 1000
        .ColWidth(38) = 700
        .ColWidth(39) = 2000
        .ColWidth(40) = 2000
    
        .Redraw = True
    End With
End Sub

Private Sub cmdRpRate_Click()
    With grdData
        If .Rows > .FixedRows + 1 Then
            
            Call frmRecipeCalc.SetInstruction(CLng(.TextMatrix(.Row, 52)), CInt(.TextMatrix(.Row, 53)))
        
        
        
        End If
    End With
    
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpDate(0) = DateSerial(Year(Date), Month(Date), Day(Date) - 1)
            dtpDate(1) = DateSerial(Year(Date), Month(Date), Day(Date) - 1)
        Case 1
            dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
            dtpDate(1) = Date
    End Select
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15300, 9660

    Call SetOperate(Me)
    
    dtpDate(0) = Now:   dtpDate(1) = Now
    
    Call MakeProcessCombo
    Call MakeMachineCombo
    Call MakePlantCombo
    Call MakeMachineNOCombo

    With cboSearch(2)
        .AddItem "전체"
        .AddItem "A"
        .AddItem "B"
        .AddItem "C"
        .ListIndex = 0
    End With

    For i = 3 To 5
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
        txtSearch(i).Enabled = False
    Next i

    cmdRpRate.Picture = LoadResPicture("FIND", vbResIcon)
    Call InitGrid

    Call ModifyGrid
    
    Show
End Sub

Private Sub cmdExit_Click()
    PlusMDI.pnlMenu.Visible = True
    Unload Me
End Sub

Private Sub MakeProcessCombo()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass
    m_bloading = True

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon

    Set rs = oRapid.GetDyeWorkProcess
    Set oRapid = Nothing

    With cboSearch(0)
        .Clear
        Do Until rs.EOF
            .AddItem CStr(rs!Process)
            .ItemData(.NewIndex) = CLng(rs!ProcessID)
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
    Set oRapid = Nothing
    m_bloading = False
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmDyeResultView.MakeProcessCombo", Err.Description)
End Sub

Private Sub MakeMachineCombo()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon

    Set rs = oRapid.GetDyeMachine(Format(cboSearch(0).ItemData(cboSearch(0).ListIndex), "0000"))
    Set oRapid = Nothing

    With cboSearch(1)
        .Clear

        .AddItem "전체"
        .ItemData(.NewIndex) = 0
        Do Until rs.EOF
            .AddItem rs!MachineNO & "호기"
            .ItemData(.NewIndex) = CLng(rs!machineid)

            rs.MoveNext
        Loop
        .AddItem "12호기"
        .ItemData(.NewIndex) = 12

        rs.Close
        Set rs = Nothing

        .ListIndex = 0
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRapid = Nothing
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmDyeResultView.MakeMachineCombo", Err.Description)
End Sub

Private Sub MakePlantCombo()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
        
    Screen.MousePointer = vbHourglass
    m_bloading = True

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon

    Set rs = oRapid.GetDyePlant
    Set oRapid = Nothing

    With cboSearch(3)
        .Clear
        Do Until rs.EOF
            .AddItem rs!Machine
            .ItemData(.NewIndex) = CLng(rs!ProcessID)
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
    Set oRapid = Nothing
    m_bloading = False
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmDyeResultView.MakePlantCombo", Err.Description)
End Sub

Private Sub MakeMachineNOCombo()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
    Dim sPlant$, i%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon

    sPlant = cboSearch(3).Text
    
    Set rs = oRapid.GetDyeMachineByPlant(sPlant)
    Set oRapid = Nothing

    With cboSearch(4)
        .Clear

        .AddItem "전체"
        .ItemData(.NewIndex) = 0
        For i = 0 To rs.RecordCount - 1
            
            .AddItem rs!MachineNO & "호기"
            .ItemData(.NewIndex) = CSng(rs!machineid)
                
                rs.MoveNext
            Next i
        rs.Close
        Set rs = Nothing
        .AddItem "12호기"
        .ItemData(.NewIndex) = 12

        .ListIndex = 0
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRapid = Nothing
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmDyeResultView.MakeMachineNOCombo", Err.Description)
End Sub

Private Sub InitGrid()
    Dim iCount As Integer
    
    With grdSum
        .Redraw = flexRDNone
        
        .FixedRows = 0:     .FixedCols = 0
        .Rows = 1:          .Cols = 12
        
        .RowHeight(0) = 300
        .FontSize = 9
        .FontBold = True
        For iCount = 0 To .Cols - 1
            .ColWidth(iCount) = 1255
            If iCount Mod 3 = 0 Then
                .ColAlignment(iCount) = flexAlignCenterCenter
                .Cell(flexcpBackColor, 0, iCount) = &H8000000F
            Else
                .ColAlignment(iCount) = flexAlignRightCenter
            End If
        Next iCount
        
        .ScrollBars = flexScrollBarNone
        .HighLight = flexHighlightNever
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeNone
        .AllowBigSelection = False
        .GridColor = vbBlack

        .WordWrap = True
        .ExtendLastCol = True
        
        .TextMatrix(0, 0) = "본염":     .TextMatrix(0, 1) = "0 건":     .TextMatrix(0, 2) = "0 YDS"
        .TextMatrix(0, 3) = "수정":     .TextMatrix(0, 4) = "0 건":     .TextMatrix(0, 5) = "0 YDS"
        .TextMatrix(0, 6) = "추가":     .TextMatrix(0, 7) = "0 건":     .TextMatrix(0, 8) = "0 YDS"
        .TextMatrix(0, 9) = "총 합계":  .TextMatrix(0, 10) = "0 건":    .TextMatrix(0, 11) = "0 YDS"

        .Redraw = flexRDDirect
    End With
    
    With grdSumModi
        .Redraw = flexRDNone
        
        .FixedRows = 0:     .FixedCols = 0
        .Rows = 4:          .Cols = 3
        
        For iCount = 0 To .Rows - 1
            .RowHeight(iCount) = 280
        Next iCount
        .FontSize = 9
'        .FontBold = True
        For iCount = 0 To .Cols - 1
            .ColWidth(iCount) = 1255
            If iCount Mod 3 = 0 Then
                .ColAlignment(iCount) = flexAlignCenterCenter
                .Cell(flexcpBackColor, 0, iCount, .Rows - 1, iCount) = &H8000000F
            Else
                .ColAlignment(iCount) = flexAlignRightCenter
            End If
        Next iCount
        
        .ScrollBars = flexScrollBarVertical
        .HighLight = flexHighlightNever
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeNone
        .AllowBigSelection = False
        .GridColor = vbBlack

        .WordWrap = True
        .ExtendLastCol = True
        
        .TextMatrix(0, 0) = "얼룩수정":         .TextMatrix(0, 1) = "0 건":     .TextMatrix(0, 2) = "0 YDS"
        .TextMatrix(1, 0) = "오염수정":         .TextMatrix(1, 1) = "0 건":     .TextMatrix(1, 2) = "0 YDS"
        .TextMatrix(2, 0) = "색  수정":         .TextMatrix(2, 1) = "0 건":     .TextMatrix(2, 2) = "0 YDS"
        .TextMatrix(3, 0) = "시와수정":         .TextMatrix(3, 1) = "0 건":     .TextMatrix(3, 2) = "0 YDS"
''        .TextMatrix(4, 0) = "탈발후 색수정":    .TextMatrix(4, 1) = "0 건":     .TextMatrix(4, 2) = "0 YDS"
''        .TextMatrix(5, 0) = "탈발후 재염":      .TextMatrix(5, 1) = "0 건":     .TextMatrix(5, 2) = "0 YDS"
''        .TextMatrix(6, 0) = "탈색":             .TextMatrix(6, 1) = "0 건":     .TextMatrix(6, 2) = "0 YDS"
''        .TextMatrix(7, 0) = "감색":             .TextMatrix(7, 1) = "0 건":     .TextMatrix(7, 2) = "0 YDS"

        .Redraw = flexRDDirect
    End With
    
End Sub

Private Function ModifyGrid() As Integer
    Dim i%
    
    With grdData
        .WordWrap = False
        .Cols = 60:     .Rows = 5
        .FixedCols = 2: .FixedRows = 5
        
        .FontSize = 9
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
'        .FocusRect = flexFocusNone
        
        For i = 0 To 4
            .RowHeight(i) = 250
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 0
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .RowHeightMin = 0
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        ' 기본내역
        .TextMatrix(3, 0) = " ":                        .ColWidth(0) = 0
        .TextMatrix(3, 1) = "N" & vbCrLf & "O":         .ColWidth(1) = 300:         .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = " "
        .TextMatrix(3, 3) = "구" & vbCrLf & "분":       .ColWidth(3) = 300:         .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "실적" & vbCrLf & "일자":   .ColWidth(4) = 600:         .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "공정명":                   .ColWidth(5) = 0:           .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(3, 6) = "기계" & vbCrLf & "NO":     .ColWidth(6) = 500:         .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(3, 7) = " "
        .TextMatrix(3, 8) = "밧자" & vbCrLf & "NO"
        .TextMatrix(3, 9) = "작업" & vbCrLf & "단위"
        .TextMatrix(3, 10) = "단위" & vbCrLf & "순위"
        .TextMatrix(3, 11) = "염색패턴":                .ColWidth(11) = 1400:       .ColAlignment(11) = flexAlignLeftCenter
        .TextMatrix(3, 12) = "작업" & vbCrLf & "구분":  .ColWidth(12) = 1000:       .ColAlignment(12) = flexAlignCenterCenter
        .TextMatrix(3, 13) = "염색구분":                .ColWidth(13) = 0:          .ColAlignment(13) = flexAlignCenterCenter
        .TextMatrix(3, 14) = "거래처":                  .ColWidth(14) = 1000:       .ColAlignment(14) = flexAlignLeftCenter
        .TextMatrix(3, 15) = "품명":                    .ColWidth(15) = 2000:       .ColAlignment(15) = flexAlignLeftCenter
        .TextMatrix(3, 16) = "색상명":                  .ColWidth(16) = 1400:       .ColAlignment(16) = flexAlignLeftCenter
        .TextMatrix(3, 17) = "OrderNo":                 .ColWidth(17) = 0:          .ColAlignment(17) = flexAlignLeftCenter
        .TextMatrix(3, 18) = "관리번호":                .ColWidth(18) = 1200:       .ColAlignment(18) = flexAlignCenterCenter
        .TextMatrix(3, 19) = "가공" & vbCrLf & "방법":  .ColWidth(19) = 0:          .ColAlignment(19) = flexAlignCenterCenter
        .TextMatrix(3, 20) = " "
        
        ' 카드, 절수, 수량
        .TextMatrix(3, 21) = "카드번호":                .ColWidth(21) = 1500:       .ColAlignment(21) = flexAlignLeftCenter
        .TextMatrix(3, 22) = "작업량":                  .ColWidth(22) = 500:        .ColAlignment(22) = flexAlignRightCenter
        .TextMatrix(3, 23) = "작업량":                  .ColWidth(23) = 700:        .ColAlignment(23) = flexAlignRightCenter
        
        ' 작업일,시,조,자
        .TextMatrix(3, 31) = "작업" & vbCrLf & "일":    .ColWidth(31) = 600:        .ColAlignment(31) = flexAlignCenterCenter
        .TextMatrix(3, 32) = "작업시간":                .ColWidth(32) = 500:        .ColAlignment(32) = flexAlignCenterCenter
        .TextMatrix(3, 33) = "작업시간":                .ColWidth(33) = 500:        .ColAlignment(33) = flexAlignCenterCenter
        .TextMatrix(3, 34) = "작업시간":                .ColWidth(34) = 600:        .ColAlignment(34) = flexAlignCenterCenter
        .TextMatrix(3, 35) = "작업" & vbCrLf & "조":    .ColWidth(35) = 500:        .ColAlignment(35) = flexAlignCenterCenter
        .TextMatrix(3, 36) = "작업자":                  .ColWidth(36) = 700:        .ColAlignment(36) = flexAlignCenterCenter
        .TextMatrix(3, 37) = "염색구분":                .ColWidth(37) = 1000:       .ColAlignment(37) = flexAlignCenterCenter
        .TextMatrix(3, 38) = "밧쟈":                    .ColWidth(38) = 700:        .ColAlignment(37) = flexAlignLeftCenter
        .TextMatrix(3, 39) = "비고":                    .ColWidth(39) = 2000:       .ColAlignment(39) = flexAlignLeftCenter
        .TextMatrix(3, 40) = "보류원인":                .ColWidth(40) = 2000:       .ColAlignment(40) = flexAlignLeftCenter
        
        ' 각종 코드
        .TextMatrix(3, 41) = "useclss(카드)"
        .TextMatrix(3, 42) = "공정코드"
        .TextMatrix(3, 43) = "염색패턴번호"
        .TextMatrix(3, 44) = "거래처코드"
        .TextMatrix(3, 45) = "품명코드"
        .TextMatrix(3, 46) = "색상코드"
        .TextMatrix(3, 47) = "관리번호"
        .TextMatrix(3, 48) = "카드번호"
        .TextMatrix(3, 49) = "분할번호"
        .TextMatrix(3, 50) = "작업조코드"
        .TextMatrix(3, 51) = "작업자코드"
        .TextMatrix(3, 52) = "스케쥴번호"
        .TextMatrix(3, 53) = "스케쥴차수"
        
        
        .TextMatrix(4, 0) = " "
        .TextMatrix(4, 1) = "N" & vbCrLf & "O"
        .TextMatrix(4, 2) = " "
        .TextMatrix(4, 3) = "구" & vbCrLf & "분"
        .TextMatrix(4, 4) = "실적" & vbCrLf & "일자"
        .TextMatrix(4, 5) = "공정명"
        .TextMatrix(4, 6) = "기계" & vbCrLf & "NO"
        .TextMatrix(4, 7) = " "
        .TextMatrix(4, 8) = "밧자" & vbCrLf & "NO"
        .TextMatrix(4, 9) = "작업" & vbCrLf & "단위"
        .TextMatrix(4, 10) = "단위" & vbCrLf & "순위"
        .TextMatrix(4, 11) = "염색패턴"
        .TextMatrix(4, 12) = "작업" & vbCrLf & "구분"
        .TextMatrix(4, 13) = "염색구분"
        .TextMatrix(4, 14) = "거래처"
        .TextMatrix(4, 15) = "품명"
        .TextMatrix(4, 16) = "색상명"
        .TextMatrix(4, 17) = "OrderNo"
        .TextMatrix(4, 18) = "관리번호"
        .TextMatrix(4, 19) = "가공" & vbCrLf & "방법"
        .TextMatrix(4, 20) = " "
        
        ' 카드, 절수, 수량
        .TextMatrix(4, 21) = "카드번호"
        .TextMatrix(4, 22) = "절수"
        .TextMatrix(4, 23) = "수량"

        ' 작업일,시,조,자
        .TextMatrix(4, 31) = "작업" & vbCrLf & "일"
        .TextMatrix(4, 32) = "시작"
        .TextMatrix(4, 33) = "종료"
        .TextMatrix(4, 34) = "소요"
        .TextMatrix(4, 35) = "작업" & vbCrLf & "조"
        .TextMatrix(4, 36) = "작업자"
        .TextMatrix(4, 37) = "염색구분"
        .TextMatrix(4, 38) = "밧쟈"
        .TextMatrix(4, 39) = "비고"
        .TextMatrix(4, 40) = "보류원인"
        
        ' 각종 코드
        .TextMatrix(4, 41) = "useclss(카드)"
        .TextMatrix(4, 42) = "공정코드"
        .TextMatrix(4, 43) = "염색패턴번호"
        .TextMatrix(4, 44) = "거래처코드"
        .TextMatrix(4, 45) = "품명코드"
        .TextMatrix(4, 46) = "색상코드"
        .TextMatrix(4, 47) = "관리번호"
        .TextMatrix(4, 48) = "카드번호"
        .TextMatrix(4, 49) = "분할번호"
        .TextMatrix(4, 50) = "작업조코드"
        .TextMatrix(4, 51) = "작업자코드"
        .TextMatrix(4, 52) = "스케쥴번호"
        .TextMatrix(4, 53) = "스케쥴차수"
           
        .MergeCells = flexMergeFree
        
        For i = 0 To 4
            .MergeRow(i) = True
        Next i
        
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i
        
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .Redraw = flexRDDirect
    End With


    dtpDate(0).Tag = MakeDate(DF_SHORT, dtpDate(0))
    cboSearch(0).Tag = ModifyGrid
    If cboSearch(1).ListCount > 0 Then cboSearch(1).Tag = cboSearch(1).ItemData(cboSearch(1).ListIndex)
    If cboSearch(2).ListCount > 0 Then cboSearch(2).Tag = cboSearch(2).ItemData(cboSearch(2).ListIndex)
    pnlCaption(1).Tag = cboSearch(1)
    pnlCaption(2).Tag = cboSearch(2)
    chkSearch(1).Tag = IIf(cboSearch(1).ListIndex > 0, 1, 0)
    chkSearch(2).Tag = IIf(cboSearch(2).ListIndex > 0, 1, 0)
End Function

Private Sub cmdExcel_Click()
    If grdData.Rows = 1 Then
        MsgBox LoadResString(203), vbInformation
        Exit Sub
    End If
    Call MakeExcelGrid(grdData)
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 3 Then
        Call ReturnCode(LG_ORDER, , False, txtSearch(Index))
    ElseIf Index = 4 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 5 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    End If
End Sub

Private Sub cmdHTML_Click()
    If grdData.Rows = 1 Then
        MsgBox LoadResString(203), vbInformation
        cmdSearch.SetFocus
        Exit Sub
    End If

    If MakeHtmlGrid(grdData, "C:\" & Me.Caption & ".html") Then
        Call RelateOpen(Me.hWnd, "C:\" & Me.Caption & ".html")
    End If
End Sub

Private Sub cboSearch_Click(Index As Integer)
    If m_bloading Then Exit Sub
    
    If Index = 1 Or Index = 4 Then Exit Sub

    If Index = 0 Then
        Call MakeMachineCombo

'        Call FillGridData
    ElseIf Index = 3 Then
        Call MakeMachineNOCombo
        
'        Call FillGridData
    
    Else
        If cboSearch(Index).ListIndex > 0 Then chkSearch(Index) = vbChecked
    End If
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 1 Or Index = 2 Then
        If chkSearch(Index) = vbUnchecked Then cboSearch(Index).ListIndex = 0
    ElseIf Index = 0 Then
        cboSearch(0).ListIndex = 0
    
    Else
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            cmdFind(Index).Enabled = True
            txtSearch(Index).SetFocus
        Else
            txtSearch(Index).Enabled = False
            cmdFind(Index).Enabled = False
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub optOrder_Click(Index As Integer)
    Dim iPrevCol%, iCol%, nSize%
    
    With grdData
        If Index = 0 Then
            chkSearch(3).Caption = "Order No."
            .ColWidth(17) = 1200
            .ColWidth(18) = 0
        Else
            chkSearch(3).Caption = "관리 번호"
            .ColWidth(17) = 0
            .ColWidth(18) = 1200
        End If
    End With
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 3 Or Index = 5 Then
            Call ReturnCode(LG_ORDER, , False, txtSearch(Index))
        ElseIf Index = 4 Or Index = 6 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 7 Or Index = 8 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    End If
End Sub

Private Function MakeTime(ByVal sTime As String) As String

    If Len(sTime) = 0 Then
        MakeTime = ":"
    Else
        MakeTime = Left(sTime, 2) & ":" & Right(sTime, 2)
    End If
End Function

Public Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub FillGridData()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
    Dim i%, iCntRec%, nClss%
    Dim sDate$, eDate$
    Dim nChkProcessID%, sProcessID$
    Dim nChkMachineID%, sMachineID$
    Dim nChkTeamID%, sTeamID$
    Dim nChkOrder%, sOrder$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim nRoll(3) As Long, nQty(3) As Long
    Dim nModiRoll(8) As Long, nModiQty(8) As Long   ' 얼룩수정, 주름수정, 오염수정, 색수정, 탈발후 색수정, 탈발후 재염, 탈색, 감색
    Dim sResultDate$, sProcID$, sMachID$
    Dim iDyeSch%, iDyeSeq%
    Dim bSub As Boolean
    

    On Error GoTo ErrHandler

    If chkPlan.Value = 0 And chkResult.Value = 0 Then
        MsgBox "계획 또는 실적 구분을 체크하여 검색하여 주십시오", vbInformation + vbOKOnly, "체크항목 선택 요구"
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    pnlCaption(2).Enabled = True
    cboSearch(2).Enabled = True

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon

    m_bSkip = True
    ' 공정별, 설비별 검색 구분
    nClss = IIf(optProcess(0).Value = True, 4, 1)


    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    
    If optProcess(0).Value Then
        nChkProcessID = 0
        sProcessID = cboSearch(3).ItemData(cboSearch(3).ListIndex)
    Else
        nChkProcessID = 1
        sProcessID = cboSearch(0).ItemData(cboSearch(0).ListIndex)
    End If
    
    nChkMachineID = IIf(cboSearch(nClss).ListIndex = 0, 0, 1)
    sMachineID = cboSearch(nClss).ItemData(cboSearch(nClss).ListIndex)
    
    nChkTeamID = IIf(cboSearch(2).ListIndex = 0, 0, 1)
    sTeamID = Format((IIf(cboSearch(2).ListIndex > 0, cboSearch(2).ListIndex, 0)), "00")
    nChkOrder = IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0)
    sOrder = txtSearch(3)
    nChkCustom = IIf(chkSearch(4).Value = vbChecked, 1, 0)
    sCustom = txtSearch(4).Tag
    nChkArticle = IIf(chkSearch(5).Value = vbChecked, 1, 0)
    sArticle = txtSearch(5).Tag

    bSub = True
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
            
        Set rs = oRapid.GetRapidResultByPlant(sDate, eDate, nChkProcessID, Trim(sProcessID), nChkMachineID, _
                         sMachineID, nChkTeamID, sTeamID, _
                        nChkOrder, sOrder, nChkCustom, sCustom, nChkArticle, sArticle, chkPlan.Value, chkResult.Value)
        Set oRapid = Nothing

        For i = 1 To rs.RecordCount
            If sResultDate <> rs!wkResultDT Or sProcID <> rs!wkProcID Or sMachID <> rs!wkMachID Or iDyeSch <> rs!DyeSchID Or iDyeSeq <> rs!DyeSeq Then
                bSub = True
            End If

            If i = 1 Then
                sMachID = rs!wkMachID
            End If
            
            If sMachID <> rs!wkMachID Then
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 25
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
            End If
            
            '------merge 때문에 중간에 빈라인 삽입( 04/02/04 최현숙 )
            .Rows = .Rows + 1
            .RowHidden(.Rows - 1) = True
            '---------------------------------------
            
            If rs!CardCnt < 2 Then      ' 카드수가 1개이거나 없는것
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 350
                If rs!wkResultDT <> "99999999" Then
                    iCntRec = iCntRec + 1
                    .TextMatrix(.Rows - 1, 1) = IIf(rs!wkResultDT = "99999999", " ", CStr(iCntRec))
                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = &HF1F1F1
                Else
                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = &HF2E8FF
                    .TextMatrix(.Rows - 1, 1) = ""
                End If
                
                .TextMatrix(.Rows - 1, 0) = " "
'                .TextMatrix(.Rows - 1, 1) = IIf(rs!wkresultdt = "99999999", " ", CStr(iCntRec))
                .TextMatrix(.Rows - 1, 2) = " "
                .TextMatrix(.Rows - 1, 3) = IIf(rs!ReWorkClss = "*", "■", "")
                If Trim(rs!RapidClss) = "" And rs!wkResultDT <> "99999999" Then
                    .Cell(flexcpBackColor, .Rows - 1, 12) = IIf(Trim(rs!EndDate) = "", vbBlue, 0)
                    .Cell(flexcpForeColor, .Rows - 1, 12) = IIf(Trim(rs!EndDate) = "", vbWhite, 0)
                End If
                

                
                .TextMatrix(.Rows - 1, 4) = IIf(rs!wkResultDT = "99999999", "계획", MakeDate(DF_MD, rs!wkResultDT))
                .TextMatrix(.Rows - 1, 5) = rs!Process
                .TextMatrix(.Rows - 1, 6) = rs!wkMachID
                .TextMatrix(.Rows - 1, 7) = " "
                .TextMatrix(.Rows - 1, 8) = " "
                .TextMatrix(.Rows - 1, 9) = rs!WorkUnitId
                .TextMatrix(.Rows - 1, 10) = rs!WorkUnitSeq
                .TextMatrix(.Rows - 1, 11) = CheckNull(rs!PtName)
                .TextMatrix(.Rows - 1, 12) = rs!WorkClss
                .TextMatrix(.Rows - 1, 13) = rs!RapidClss
                .TextMatrix(.Rows - 1, 14) = rs!kCustom
                .TextMatrix(.Rows - 1, 15) = rs!Article
                .TextMatrix(.Rows - 1, 16) = Trim(rs!Color)
                .TextMatrix(.Rows - 1, 17) = rs!OrderNo
                .TextMatrix(.Rows - 1, 18) = IIf(Trim(rs!OrderID) = "", "", MakeOrderID(rs!OrderID, OM_EXPAND))
                .TextMatrix(.Rows - 1, 19) = " "
                .TextMatrix(.Rows - 1, 20) = " "
                .TextMatrix(.Rows - 1, 21) = IIf(Trim(rs!CardID) = "", "", MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID))
                
                If rs!CardCnt = 1 Then
                    Select Case rs!UseClss
                        Case "작업":
                            .Cell(flexcpBackColor, .Rows - 1, 21) = vbBlue
                            .Cell(flexcpForeColor, .Rows - 1, 21) = vbWhite
                        Case "보류":
                            .Cell(flexcpBackColor, .Rows - 1, 21) = vbRed
                            .Cell(flexcpForeColor, .Rows - 1, 21) = vbWhite
                    End Select
                End If
                
                .TextMatrix(.Rows - 1, 22) = Format(rs!WkRoll, "#,###")
                .TextMatrix(.Rows - 1, 23) = Format(rs!WkQty, "#,###,###")
                
                ' 작업일,시,조,자
                .TextMatrix(.Rows - 1, 31) = IIf(Trim(rs!StartDate) = "", "", MakeDate(DF_MD, rs!StartDate))
                .TextMatrix(.Rows - 1, 32) = IIf(Trim(rs!StartTime) = "", "", MakeTime(rs!StartTime))
                .TextMatrix(.Rows - 1, 33) = IIf(Trim(rs!EndTime) = "", "", MakeTime(rs!EndTime))
                .TextMatrix(.Rows - 1, 34) = IIf(rs!requiredtime = "0", "", rs!requiredtime)
                If rs!wkResultDT <> "99999999" Then
                    .TextMatrix(.Rows - 1, 35) = IIf(Format(Trim(rs!TeamID), "00") = "01", "A", IIf(Format(Trim(rs!TeamID), "00") = "02", "B", "C"))
                Else
                    .TextMatrix(.Rows - 1, 35) = " "
                End If
                .TextMatrix(.Rows - 1, 36) = rs!Name
                .TextMatrix(.Rows - 1, 37) = IIf(rs!RapidClss = "본염", "", rs!RapidClss)
                .TextMatrix(.Rows - 1, 38) = rs!BatJaNO
                .TextMatrix(.Rows - 1, 39) = IIf(Trim(rs!ReWorkClss) = "*", Trim(rs!RapidClss), rs!Remark)
                .TextMatrix(.Rows - 1, 40) = rs!HoldReason
                
                ' 각종 코드
                .TextMatrix(.Rows - 1, 41) = rs!UseClss
                .TextMatrix(.Rows - 1, 42) = rs!ProcessID
                .TextMatrix(.Rows - 1, 43) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 44) = rs!CustomID
                .TextMatrix(.Rows - 1, 45) = rs!ArticleID
                .TextMatrix(.Rows - 1, 46) = ""
                .TextMatrix(.Rows - 1, 47) = rs!OrderID
                .TextMatrix(.Rows - 1, 48) = Trim(rs!CardID)
                .TextMatrix(.Rows - 1, 49) = rs!SplitID
                .TextMatrix(.Rows - 1, 50) = rs!TeamID
                .TextMatrix(.Rows - 1, 51) = rs!PersonID
                .TextMatrix(.Rows - 1, 52) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 53) = rs!DyeSeq
                
                If rs!wkResultDT <> "99999999" Then
                    Select Case rs!RapidClss
                    
                        Case "본염":    nRoll(0) = nRoll(0) + rs!workroll
                                        nQty(0) = nQty(0) + rs!workqty
                        Case "추가":    nRoll(2) = nRoll(2) + rs!workroll
                                        nQty(2) = nQty(2) + rs!workqty
                        Case Else:
'                            If Trim(rs!RapidClss) <> "" Then
'                                nRoll(1) = nRoll(1) + rs!workroll
'                                nQty(1) = nQty(1) + rs!workqty
'                            End If
                            
                            Select Case rs!RapidClss
                                Case "얼룩수정":
                                        nModiRoll(0) = nModiRoll(0) + rs!workroll
                                        nModiQty(0) = nModiQty(0) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                Case "오염수정":
                                        nModiRoll(1) = nModiRoll(1) + rs!workroll
                                        nModiQty(1) = nModiQty(1) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                Case "색수정":
                                        nModiRoll(2) = nModiRoll(2) + rs!workroll
                                        nModiQty(2) = nModiQty(2) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                Case "시와수정":
                                        nModiRoll(3) = nModiRoll(3) + rs!workroll
                                        nModiQty(3) = nModiQty(3) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                
''                                Case "색수정":
''                                        nModiRoll(3) = nModiRoll(3) + rs!workroll
''                                        nModiQty(3) = nModiQty(3) + rs!workqty
''                                Case "탈발후 색수정":
''                                        nModiRoll(4) = nModiRoll(4) + rs!workroll
''                                        nModiQty(4) = nModiQty(4) + rs!workqty
''                                Case "탈발후 재염":
''                                        nModiRoll(5) = nModiRoll(5) + rs!workroll
''                                        nModiQty(5) = nModiQty(5) + rs!workqty
''                                Case "탈색":
''                                        nModiRoll(6) = nModiRoll(6) + rs!workroll
''                                        nModiQty(6) = nModiQty(6) + rs!workqty
''                                Case "감색":
''                                        nModiRoll(7) = nModiRoll(7) + rs!workroll
''                                        nModiQty(7) = nModiQty(7) + rs!workqty
                            End Select
                    End Select
                
                End If
                
                bSub = True
            Else                        ' 카드개수가 2개 이상인것
                If bSub = True Then
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = 350
                    
                    If rs!wkResultDT <> "99999999" Then
                        iCntRec = iCntRec + 1
                        .TextMatrix(.Rows - 1, 1) = IIf(rs!wkResultDT = "99999999", " ", CStr(iCntRec))
                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = &HF1F1F1
                    Else
                        .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = &HF2E8FF
                        .TextMatrix(.Rows - 1, 1) = ""
                    End If
                    
                    .TextMatrix(.Rows - 1, 0) = " "
'                    .TextMatrix(.Rows - 1, 1) = IIf(rs!wkresultdt = "99999999", " ", CStr(iCntRec))
                    .TextMatrix(.Rows - 1, 2) = " "
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!ReWorkClss = "*", "■", "")
                    If Trim(rs!RapidClss) = "" And rs!wkResultDT <> "99999999" Then
                        .Cell(flexcpBackColor, .Rows - 1, 12) = IIf(Trim(rs!EndDate) = "", vbBlue, 0)
                        .Cell(flexcpForeColor, .Rows - 1, 12) = IIf(Trim(rs!EndDate) = "", vbWhite, 0)
                    End If
                    .TextMatrix(.Rows - 1, 4) = IIf(rs!wkResultDT = "99999999", "계획", MakeDate(DF_MD, rs!wkResultDT))
                    .TextMatrix(.Rows - 1, 5) = rs!Process
                    .TextMatrix(.Rows - 1, 6) = rs!wkMachID
                    .TextMatrix(.Rows - 1, 7) = " "
                    .TextMatrix(.Rows - 1, 8) = " "
                    .TextMatrix(.Rows - 1, 9) = rs!WorkUnitId
                    .TextMatrix(.Rows - 1, 10) = rs!WorkUnitSeq
                    .TextMatrix(.Rows - 1, 11) = CheckNull(rs!PtName)
                    .TextMatrix(.Rows - 1, 12) = rs!WorkClss
                    .TextMatrix(.Rows - 1, 13) = rs!RapidClss
                    .TextMatrix(.Rows - 1, 22) = Format(rs!WkRoll, "#,###")
                    .TextMatrix(.Rows - 1, 23) = Format(rs!WkQty, "#,###,###")
                    
                    ' 작업일,시,조,자
                    .TextMatrix(.Rows - 1, 31) = IIf(Trim(rs!StartDate) = "", "", MakeDate(DF_MD, rs!StartDate))
                    .TextMatrix(.Rows - 1, 32) = IIf(Trim(rs!StartTime) = "", "", MakeTime(rs!StartTime))
                    .TextMatrix(.Rows - 1, 33) = IIf(Trim(rs!EndTime) = "", "", MakeTime(rs!EndTime))
                    .TextMatrix(.Rows - 1, 34) = IIf(rs!requiredtime = "0", "", rs!requiredtime)
                    If rs!wkResultDT <> "99999999" Then
                        .TextMatrix(.Rows - 1, 35) = IIf(Format(Trim(rs!TeamID), "00") = "01", "A", IIf(Format(Trim(rs!TeamID), "00") = "02", "B", "C"))
                    Else
                        .TextMatrix(.Rows - 1, 35) = " "
                    End If
                    .TextMatrix(.Rows - 1, 36) = rs!Name
                    .TextMatrix(.Rows - 1, 37) = IIf(rs!RapidClss = "본염", "", rs!RapidClss)
                    .TextMatrix(.Rows - 1, 38) = rs!BatJaNO
                    .TextMatrix(.Rows - 1, 39) = rs!Remark
                    
                    ' 각종 코드
                    .TextMatrix(.Rows - 1, 41) = rs!UseClss
                    .TextMatrix(.Rows - 1, 42) = rs!ProcessID
                    .TextMatrix(.Rows - 1, 43) = rs!DyeSchID
                    .TextMatrix(.Rows - 1, 44) = rs!CustomID
                    .TextMatrix(.Rows - 1, 45) = rs!ArticleID
                    .TextMatrix(.Rows - 1, 46) = ""
                    .TextMatrix(.Rows - 1, 47) = rs!OrderID
                    .TextMatrix(.Rows - 1, 48) = rs!CardID
                    .TextMatrix(.Rows - 1, 49) = rs!SplitID
                    .TextMatrix(.Rows - 1, 50) = rs!TeamID
                    .TextMatrix(.Rows - 1, 51) = rs!PersonID
                    .TextMatrix(.Rows - 1, 52) = rs!DyeSchID
                    .TextMatrix(.Rows - 1, 53) = rs!DyeSeq
                    
'                    .IsSubtotal(.Rows - 1) = True
                    
                    bSub = False
'                    '------merge 때문에 중간에 빈라인 삽입( 04/02/04 최현숙 )
                    
                    .Rows = .Rows + 1
                    .RowHidden(.Rows - 1) = True
                    '---------------------------------------
                End If
                
                    
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 300
                    
                .TextMatrix(.Rows - 1, 0) = " "
        '        .TextMatrix(.Rows - 1, 1) = IIf(rs!wkresultdt = "99999999", " ", CStr(iCntRec))
                .TextMatrix(.Rows - 1, 14) = rs!kCustom
                .TextMatrix(.Rows - 1, 15) = rs!Article
                .TextMatrix(.Rows - 1, 16) = Trim(rs!Color)
                .TextMatrix(.Rows - 1, 17) = rs!OrderNo
                .TextMatrix(.Rows - 1, 18) = IIf(Trim(rs!OrderID) = "", "", MakeOrderID(rs!OrderID, OM_EXPAND))
                .TextMatrix(.Rows - 1, 21) = IIf(Trim(rs!CardID) = "", "", MakeCardID(rs!CardID, OM_EXPAND)) & IIf(Trim(rs!SplitID) = "", "", "(" & Trim(rs!SplitID) & ")")
                .TextMatrix(.Rows - 1, 22) = Format(rs!workroll, "#,###")
                .TextMatrix(.Rows - 1, 23) = Format(rs!workqty, "#,###,###")
                
                ' 각종 코드
                .TextMatrix(.Rows - 1, 41) = rs!UseClss
                .TextMatrix(.Rows - 1, 42) = rs!ProcessID
                .TextMatrix(.Rows - 1, 43) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 44) = rs!CustomID
                .TextMatrix(.Rows - 1, 45) = rs!ArticleID
                .TextMatrix(.Rows - 1, 46) = ""
                .TextMatrix(.Rows - 1, 47) = rs!OrderID
                .TextMatrix(.Rows - 1, 48) = rs!CardID
                .TextMatrix(.Rows - 1, 49) = rs!SplitID
                .TextMatrix(.Rows - 1, 50) = rs!TeamID
                .TextMatrix(.Rows - 1, 51) = rs!PersonID
                .TextMatrix(.Rows - 1, 52) = rs!DyeSchID
                .TextMatrix(.Rows - 1, 53) = rs!DyeSeq
                
'                If rs!wkresultdt = "99999999" Then
'                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = &HF2E8FF
'                Else
'                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = &HF1F1F1
'                End If
                
                Select Case rs!UseClss
                    Case "작업":
                        .Cell(flexcpBackColor, .Rows - 1, 21) = vbBlue
                        .Cell(flexcpForeColor, .Rows - 1, 21) = vbWhite
                    Case "보류":
                        .Cell(flexcpBackColor, .Rows - 1, 21) = vbRed
                        .Cell(flexcpForeColor, .Rows - 1, 21) = vbWhite
                End Select
                
                If rs!wkResultDT <> "99999999" Then
                    Select Case rs!RapidClss
                        Case "본염":    nRoll(0) = nRoll(0) + rs!workroll
                                        nQty(0) = nQty(0) + rs!workqty
                        Case "추가":    nRoll(2) = nRoll(2) + rs!workroll
                                        nQty(2) = nQty(2) + rs!workqty
                        Case Else:
''                        nRoll(1) = nRoll(1) + rs!workroll
''                                        nQty(1) = nQty(1) + rs!workqty
                            Select Case rs!RapidClss
                                Case "얼룩수정":
                                        nModiRoll(0) = nModiRoll(0) + rs!workroll
                                        nModiQty(0) = nModiQty(0) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                Case "오염수정":
                                        nModiRoll(1) = nModiRoll(1) + rs!workroll
                                        nModiQty(1) = nModiQty(1) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                Case "색수정":
                                        nModiRoll(2) = nModiRoll(2) + rs!workroll
                                        nModiQty(2) = nModiQty(2) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty
                                Case "시와수정":
                                        nModiRoll(3) = nModiRoll(3) + rs!workroll
                                        nModiQty(3) = nModiQty(3) + rs!workqty
                                        nRoll(1) = nRoll(1) + rs!workroll
                                        nQty(1) = nQty(1) + rs!workqty

''                                Case "색수정":
''                                        nModiRoll(3) = nModiRoll(3) + rs!workroll
''                                        nModiQty(3) = nModiQty(3) + rs!workqty
''                                Case "탈발후 색수정":
''                                        nModiRoll(4) = nModiRoll(4) + rs!workroll
''                                        nModiQty(4) = nModiQty(4) + rs!workqty
''                                Case "탈발후 재염":
''                                        nModiRoll(5) = nModiRoll(5) + rs!workroll
''                                        nModiQty(5) = nModiQty(5) + rs!workqty
''                                Case "탈색":
''                                        nModiRoll(6) = nModiRoll(6) + rs!workroll
''                                        nModiQty(6) = nModiQty(6) + rs!workqty
''                                Case "감색":
''                                        nModiRoll(7) = nModiRoll(7) + rs!workroll
''                                        nModiQty(7) = nModiQty(7) + rs!workqty
                            End Select
                    End Select
                End If
                '------merge 때문에 중간에 빈라인 삽입( 04/02/04 최현숙 )
                .Rows = .Rows + 1
                .RowHidden(.Rows - 1) = True
                '---------------------------------------
                
            End If
            
            sResultDate = rs!wkResultDT
            sProcID = rs!wkProcID
            sMachID = rs!wkMachID
            iDyeSch = rs!DyeSchID
            iDyeSeq = rs!DyeSeq
            
            rs.MoveNext
        Next i

        If .Rows > .FixedRows Then
           ' cmdPrint.Enabled = True
        Else
            .HighLight = flexHighlightNever
           ' cmdPrint.Enabled = False
        End If

        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    rs.Close
    Set rs = Nothing

    m_bSkip = False

'    Call cmdMinus_Click
    
    With grdSum
        .TextMatrix(0, 1) = SetCurrency(nRoll(0)) & " 절"
        .TextMatrix(0, 2) = SetCurrency(nQty(0)) & " YDS"
        .TextMatrix(0, 4) = SetCurrency(nRoll(1)) & " 절"
        .TextMatrix(0, 5) = SetCurrency(nQty(1)) & " YDS"
        .TextMatrix(0, 7) = SetCurrency(nRoll(2)) & " 절"
        .TextMatrix(0, 8) = SetCurrency(nQty(2)) & " YDS"
        .TextMatrix(0, 10) = SetCurrency(nRoll(0) + nRoll(1) + nRoll(2)) & " 절"
        .TextMatrix(0, 11) = SetCurrency(nQty(0) + nQty(1) + nQty(2)) & " YDS"
        .Cell(flexcpForeColor, 0, 9, 0, 11) = vbRed
    End With
    
    With grdData
        
        .AddItem ""
        .RowHeight(.Rows - 1) = 350
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 14) = "본염: " & SetCurrency(nRoll(0)) & "절     " & SetCurrency(nQty(0)) & " YDS"
        .Cell(flexcpText, .Rows - 1, 15, .Rows - 1, 18) = "수정: " & SetCurrency(nRoll(1)) & "절     " & SetCurrency(nQty(1)) & " YDS"
        .Cell(flexcpText, .Rows - 1, 19, .Rows - 1, 27) = "추가: " & SetCurrency(nRoll(2)) & "절     " & SetCurrency(nQty(2)) & " YDS"
        .Cell(flexcpText, .Rows - 1, 28, .Rows - 1, 39) = "총 합계: " & SetCurrency(nRoll(0) + nRoll(1) + nRoll(2)) & "절     " & SetCurrency(nQty(0) + nQty(1) + nQty(2)) & " YDS"
        
        .MergeRow(.Rows - 1) = True
        .RowHidden(.Rows - 1) = True
        .Redraw = flexRDDirect
    End With
    
    With grdSumModi
        For i = 0 To 3
            .TextMatrix(i, 1) = SetCurrency(nModiRoll(i)) & " 절"
            .TextMatrix(i, 2) = SetCurrency(nModiQty(i)) & " YDS"
        Next i
    End With

    If cboSearch(1).ListCount > 0 Then cboSearch(1).Tag = cboSearch(1).ItemData(cboSearch(1).ListIndex)
    If cboSearch(2).ListCount > 0 Then cboSearch(2).Tag = cboSearch(2).ListIndex - 1
    pnlCaption(1).Tag = cboSearch(1)
    pnlCaption(2).Tag = cboSearch(2)
    chkSearch(1).Tag = IIf(cboSearch(1).ListIndex > 0, 1, 0)
    chkSearch(2).Tag = IIf(cboSearch(2).ListIndex > 0, 1, 0)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRapid = Nothing
    Screen.MousePointer = vbDefault
    m_bSkip = False
    
    Call ErrorBox(Err.Number, "frmDyeResultView.FillGridData", Err.Description)
End Sub




'Private Sub PrintinspectRecord()
'    Dim oCard As PlusLib2.CCard
'    Dim rs As ADODB.Recordset
'    Dim sParam() As String
'    Dim SDate$, EDate$
'
'    On Error GoTo ErrHandler
'
'    Set oCard = New PlusLib2.CCard
'    oCard.Connection = g_adoCon
'    oCard.UserName = g_sUserName
'
'    SDate = MakeDate(DF_SHORT, dtpDate(0))
'    EDate = MakeDate(DF_SHORT, dtpDate(1))
'
'    'IIf(chkSearch(3), 1, 0), txtSearch(3).Tag
'    Set rs = oCard.GetInspectRecordByDate(SDate, EDate, IIf(chkSearch(4), 1, 0), txtSearch(4).Tag, IIf(chkSearch(7), 1, 0), txtSearch(7).Tag)
'    Set oCard = Nothing
'
'    ReDim sParam(2)
'    sParam(0) = "일일 검사현황"
'    sParam(1) = CompanyName
'
'    If SDate = EDate Then
'        sParam(2) = "검사일자  : " & MakeDate(DF_LONG, dtpDate(0))
'    Else
'        sParam(2) = "검사일자  : " & MakeDate(DF_LONG, dtpDate(0)) & " ~ " & MakeDate(DF_LONG, dtpDate(1))
'    End If
'    Call PrintReport(REPORTFILE8, rs, sParam, PlusMDI.PrintPreview)
'
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'ErrHandler:
'    Screen.MousePointer = vbDefault
'    Set oCard = Nothing
'    Set rs = Nothing
'    Call ErrorBox(Err.Number, "frmRecord.cmdPrint_Click", Err.Description)
'End Sub
'
