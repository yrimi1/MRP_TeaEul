VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProcessResultModify 
   ClientHeight    =   9765
   ClientLeft      =   15
   ClientTop       =   630
   ClientWidth     =   15180
   ForeColor       =   &H00808080&
   Icon            =   "frmProcessResultModify.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   15180
   Begin VB.TextBox txtCardID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   12450
      MaxLength       =   12
      TabIndex        =   46
      Top             =   60
      Width           =   1605
   End
   Begin VB.Frame fraOrder 
      Height          =   465
      Left            =   7950
      TabIndex        =   32
      Top             =   -90
      Width           =   2865
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   195
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   1425
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   210
         Width           =   1155
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Index           =   5
      Left            =   12450
      TabIndex        =   31
      Top             =   735
      Width           =   1605
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Index           =   4
      Left            =   9210
      TabIndex        =   30
      Top             =   720
      Width           =   1605
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Index           =   3
      Left            =   9210
      TabIndex        =   29
      Top             =   390
      Width           =   1605
   End
   Begin VB.ComboBox cboSearch 
      Height          =   300
      Index           =   2
      Left            =   12465
      Style           =   2  '드롭다운 목록
      TabIndex        =   28
      Top             =   405
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   0
      TabIndex        =   9
      Top             =   -90
      Width           =   2010
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
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
         Format          =   113639424
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   435
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
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
         Format          =   113639424
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "실적 일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlData 
      Height          =   1845
      Left            =   2835
      TabIndex        =   3
      Top             =   4590
      Visible         =   0   'False
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   3254
      _Version        =   196609
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdCancel 
         Height          =   615
         Left            =   8610
         TabIndex        =   4
         Top             =   1125
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1085
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdEdit 
         Height          =   900
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   9915
         _cx             =   17489
         _cy             =   1587
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
         Editable        =   1
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
      Begin Threed.SSCommand cmdSave 
         Height          =   615
         Left            =   7125
         TabIndex        =   6
         Top             =   1125
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1085
         _Version        =   196609
         Caption         =   "저장"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   1695
         Left            =   90
         Top             =   90
         Width           =   9975
      End
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   690
      Left            =   11685
      TabIndex        =   2
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "편집"
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   735
      Left            =   14160
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   1
      ToolTipText     =   "자료 저장"
      Top             =   0
      Width           =   990
   End
   Begin Crystal.CrystalReport CryReport 
      Left            =   14070
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13500
      TabIndex        =   0
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdSum 
      Height          =   990
      Left            =   0
      TabIndex        =   7
      Top             =   7485
      Width           =   15135
      _cx             =   26696
      _cy             =   1746
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
      Height          =   6435
      Left            =   0
      TabIndex        =   8
      Top             =   1050
      Width           =   15195
      _cx             =   26802
      _cy             =   11351
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   1035
      Left            =   2010
      TabIndex        =   13
      Top             =   0
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1826
      _Version        =   196609
      Begin VB.OptionButton optProcess 
         Caption         =   "설비별"
         Height          =   375
         Index           =   0
         Left            =   60
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optProcess 
         Caption         =   "공정별"
         Height          =   405
         Index           =   1
         Left            =   60
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   540
         Width           =   1020
      End
      Begin Threed.SSPanel pnlProcess 
         Height          =   405
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   540
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   714
         _Version        =   196609
         Enabled         =   0   'False
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   1
            Left            =   3705
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   45
            Width           =   1020
         End
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   0
            Left            =   1080
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   45
            Width           =   1500
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   2700
            TabIndex        =   19
            Top             =   45
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "기    계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch 
               Caption         =   "기계호기"
               Height          =   180
               Index           =   1
               Left            =   75
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   315
               Width           =   1035
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   21
            Top             =   45
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "공 정 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel pnlProcess 
         Height          =   390
         Index           =   0
         Left            =   1080
         TabIndex        =   22
         Top             =   105
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   688
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   3
            Left            =   1065
            Style           =   2  '드롭다운 목록
            TabIndex        =   24
            Top             =   45
            Width           =   1500
         End
         Begin VB.ComboBox cboSearch 
            Height          =   300
            Index           =   4
            Left            =   3705
            Style           =   2  '드롭다운 목록
            TabIndex        =   23
            Top             =   45
            Width           =   1020
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   2700
            TabIndex        =   25
            Top             =   45
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "기    계"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch 
               Caption         =   "기계호기"
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   345
               Width           =   1035
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   60
            TabIndex        =   27
            Top             =   45
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "설 비 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   11190
      TabIndex        =   35
      Top             =   390
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "작 업 조"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "작 업 조"
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   60
         Width           =   960
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   5
      Left            =   7950
      TabIndex        =   37
      Top             =   390
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "Order No."
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   60
         Width           =   1125
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   7950
      TabIndex        =   39
      Top             =   720
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "작 업 조"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   4
         Left            =   75
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   60
         Width           =   960
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   3
      Left            =   10830
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   390
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
      Left            =   10830
      TabIndex        =   42
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
      Index           =   10
      Left            =   11190
      TabIndex        =   43
      Top             =   720
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   180
         Index           =   5
         Left            =   75
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   60
         Width           =   1050
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   5
      Left            =   14070
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   750
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
      Index           =   8
      Left            =   11190
      TabIndex        =   47
      Top             =   45
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkCardSearch 
         Caption         =   "카드번호"
         Height          =   180
         Left            =   75
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   60
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmProcessResultModify"
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


Private Sub chkCardSearch_Click()
    txtCardID.Enabled = IIf(chkCardSearch.Value = vbChecked, True, False)
End Sub

Private Sub cmdCancel_Click()
    pnlData.Visible = False
End Sub


Private Sub cmdEdit_Click()
    Dim nProcess%
    Dim sCardID$, sProcess$, nWorkSeq%, sSplitID$
    Dim nPosition%
    
    If optProcess(0).Value = True Then
        nProcess = cboSearch(3).ItemData(cboSearch(3).ListIndex)
    Else
        nProcess = cboSearch(0).ItemData(cboSearch(0).ListIndex)
    End If
    
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    If grdData.ColWidth(44) = 0 Then Exit Sub
    
    
    ' Airo 나 Calender 공정은 수정 불가...
    If nProcess = PC_Airo Or nProcess = PC_Calender Or nProcess = PC_CPB Then
        Exit Sub
    End If
    
    Call InitEditGrid(nProcess)
    ' 텐터, 건조, 수세, 모소, m/c, cpb전처리,  peach, 샴푸
        
    If grdData.Row - grdData.TopRow > 10 Then
        nPosition = 1
    Else
        nPosition = 2
    End If
    
    pnlData.Move 1920, IIf(nPosition = 1, 1850, 5625)
    pnlData.Visible = True
    
    With grdEdit
        .TextMatrix(1, 1) = MakeCardID(grdData.TextMatrix(grdData.Row, 38), OM_EXPAND) ' 카드번호
        .TextMatrix(1, 2) = cboSearch(0).Text                   ' 공정명
        .TextMatrix(1, 3) = grdData.TextMatrix(grdData.Row, 40)     ' 공정순위
        .TextMatrix(1, 4) = grdData.TextMatrix(grdData.Row, 39)     ' 분할번호
    
        Select Case nProcess
            ' 텐터 - '셋팅 '폭줄  '가공
            Case PC_Setting, PC_WidthLine, PC_FinalSetting
                ' 작업조건 - 온도, 속도, OverFeed,  위사밀도, Setting, 작업조건, 건조정도, 불량내역, 불량내역코드
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' 온도
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 속도
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' OverFeed
                .TextMatrix(1, 8) = grdData.TextMatrix(grdData.Row, 47)     ' 위사 밀도
                .TextMatrix(1, 9) = grdData.TextMatrix(grdData.Row, 48)     ' 작업조건
                .TextMatrix(1, 10) = grdData.TextMatrix(grdData.Row, 49)     ' WR사용량
                .TextMatrix(1, 11) = grdData.TextMatrix(grdData.Row, 50)     ' WR사용액
                
            
            ' 건조
            Case PC_Dry
                ' 작업조건 - 온도, 속도, OverFeed, 건조정도, 불량내역, 불량내역코드
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' 온도
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 속도
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' OverFeed
                .TextMatrix(1, 8) = grdData.TextMatrix(grdData.Row, 47)     ' 밀도
                
                
            ' 수세
            Case PC_REFINE, PC_SK
                ' 작업조건 - 온도, 속도, 정련구분, 셋팅
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' 온도
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 속도
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' 밀도
                            
            ' 모소
            Case PC_Moso  ' 모소
                ' 작업조건 - 단면/양면구분, 풍량, 가스량, 속도
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' 단면, 양면구분
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 풍량
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' 가스
                .TextMatrix(1, 8) = grdData.TextMatrix(grdData.Row, 47)     ' 속도
             
                
            ' m/c
            Case PC_SK, PC_NewST, PC_OBoxSK
                ' 작업조건 - RPM,온도, 조제명, 조제코드, 조제농도
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' RPM
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 온도
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' 조제명
                .TextMatrix(1, 8) = grdData.TextMatrix(grdData.Row, 47)     ' 조제코드
                .TextMatrix(1, 9) = grdData.TextMatrix(grdData.Row, 48)     ' 농도
            '    .TextMatrix(1, 10) = grdData.TextMatrix(grdData.Row, 28)    ' 카드 분할번호
                
            ' CPBPre - 전처리, 1차호발, 정련, 1차정련, 2차정련, 2차감량, LBOX 전처리, CPB 전처리, 신 ST 전처리
            Case PC_Pre, PC_1stHobal, PC_Pufiry, PC_1stPurify, PC_2ndPurify, PC_2ndReduce, PC_LBoxPre, PC_CPBPre, PC_NewSTPre
                ' 작업조건 - 속도, 정련구분
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' 속도
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 밀도
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' 베이스 온도
'                .TextMatrix(1, 8) = grdData.TextMatrix(grdData.Row, 47)     ' 에이징 온도
                             
            ' Peach
            Case PC_Peach
                ' 작업조건 - 폭, 속도, 페파본1, 페파본2, 페파본3, 밀도, 장력, 압력1, 압력2, 압력 3
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)     ' 속도
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 페파본1
                .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 46)     ' 페파본2
                .TextMatrix(1, 8) = grdData.TextMatrix(grdData.Row, 47)     ' 페파본3
                .TextMatrix(1, 9) = grdData.TextMatrix(grdData.Row, 48)     ' 페파본4
                .TextMatrix(1, 10) = grdData.TextMatrix(grdData.Row, 49)    ' 밀도
                .TextMatrix(1, 11) = grdData.TextMatrix(grdData.Row, 50)    ' 장력
                .TextMatrix(1, 12) = grdData.TextMatrix(grdData.Row, 51)    ' 압력1
                .TextMatrix(1, 13) = grdData.TextMatrix(grdData.Row, 52)    ' 압력2
                .TextMatrix(1, 14) = grdData.TextMatrix(grdData.Row, 53)    ' 압력3
             '   .TextMatrix(1, 14) = grdData.TextMatrix(grdData.Row, 28)    ' 카드분할번호
                
            ' 샴푸
            Case PC_Shampu
                ' 작업조건 - 속도, 실측율
                
                .TextMatrix(1, 5) = grdData.TextMatrix(grdData.Row, 44)    ' 속도
                .TextMatrix(1, 6) = grdData.TextMatrix(grdData.Row, 45)     ' 실측율
            '    .TextMatrix(1, 7) = grdData.TextMatrix(grdData.Row, 28)     ' 카드 분할번호
                
        End Select
    End With
        
End Sub



Private Sub InitEditGrid(nProcess As Integer)

    With grdEdit
        Call SetVSFlexGrid(grdEdit)
        
        .Redraw = flexRDNone
        .Rows = 2
        .RowHeightMin = 370
        .SelectionMode = flexSelectionFree
        .FixedCols = 5
        
       
            
            .TextArray(1) = "카드번호":                 .ColWidth(1) = 1700:            .ColAlignment(1) = flexAlignCenterCenter
            .TextArray(2) = "공정명":                   .ColWidth(2) = 0:               .ColAlignment(2) = flexAlignLeftCenter
            .TextArray(3) = "작업순위":                 .ColWidth(3) = 0:               .ColAlignment(3) = flexAlignLeftCenter
            .TextArray(4) = "분할" & vbCrLf & "번호":   .ColWidth(4) = 600:             .ColAlignment(4) = flexAlignCenterCenter
            
        Select Case nProcess
            
            ' 텐터 - '셋팅 '폭줄  '가공
            Case PC_Setting, PC_WidthLine, PC_FinalSetting
                .Cols = 12
                ' 작업조건 - 온도, 속도, OverFeed,  위사밀도, Setting, 작업조건, 건조정도, 불량내역, 불량내역코드
                .TextArray(5) = "온도":                     .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "속도":                     .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "Over" & vbCrLf & "Feed":   .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
                .TextArray(8) = "위사" & vbCrLf & "밀도":   .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignCenterCenter
                .TextArray(9) = "작업" & vbCrLf & "구분":   .ColWidth(9) = 800:             .ColAlignment(9) = flexAlignCenterCenter
                .TextArray(10) = "WR" & vbCrLf & "사용량":  .ColWidth(10) = 800:            .ColAlignment(10) = flexAlignCenterCenter
                .TextArray(11) = "WR" & vbCrLf & "사용액":  .ColWidth(11) = 800:            .ColAlignment(11) = flexAlignCenterCenter
                
            ' 건조
            Case PC_Dry
                .Cols = 9
                ' 작업조건 - 온도, 속도, OverFeed, 건조정도, 불량내역, 불량내역코드
                .TextArray(5) = "온도":                     .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "속도":                     .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "Over" & vbCrLf & "Feed":   .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
                .TextArray(8) = "밀도":                     .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignCenterCenter
                
            ' 수세
            Case PC_REFINE, PC_SK
                .Cols = 8
                ' 작업조건 - 온도, 속도, 정련구분, 셋팅
                .TextArray(5) = "온도":                     .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "속도":                     .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "밀도":                     .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
                
                            
            ' 모소
            Case PC_Moso  ' 모소
                ' 작업조건 - 단면/양면구분, 풍량, 가스량, 속도
                .Cols = 9
                
                .TextArray(5) = "단면/" & vbCrLf & "양면구분":  .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "풍량":                         .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "가스량":                       .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
                .TextArray(8) = "속도":                         .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignCenterCenter
                
            ' m/c
            Case PC_SK, PC_NewST, PC_OBoxSK
                ' 작업조건 - RPM,온도, 조제명, 조제코드, 조제농도
                .Cols = 10
                .TextArray(5) = "RPM":                      .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "온도":                     .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "조제명":                   .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
                .TextArray(8) = "조제" & vbCrLf & "코드":   .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignCenterCenter
                .TextArray(9) = "조제" & vbCrLf & "농도":   .ColWidth(9) = 800:             .ColAlignment(9) = flexAlignCenterCenter
                    
            ' CPBPre - 전처리, 1차호발, 정련, 1차정련, 2차정련, 2차감량, LBOX 전처리, CPB 전처리, 신 ST 전처리
            Case PC_Pre, PC_1stHobal, PC_Pufiry, PC_1stPurify, PC_2ndPurify, PC_2ndReduce, PC_LBoxPre, PC_CPBPre, PC_NewSTPre
                ' 작업조건 - 속도, 정련구분
                .Cols = 8
                .TextArray(5) = "속도":                     .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "베이스" & vbCrLf & "온도": .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "NaOH" & vbCrLf & "량(ℓ/)": .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
                                
            ' Peach
            Case PC_Peach
                ' 작업조건 - 폭, 속도, 페파본1, 페파본2, 페파본3, 밀도, 장력, 압력1, 압력2, 압력 3
                .Cols = 15
                
                .TextArray(5) = "속도":             .ColWidth(5) = 750:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "페파본1":          .ColWidth(6) = 750:             .ColAlignment(6) = flexAlignCenterCenter
                .TextArray(7) = "페파본2":          .ColWidth(7) = 750:             .ColAlignment(7) = flexAlignCenterCenter
                .TextArray(8) = "페파본3":          .ColWidth(8) = 750:             .ColAlignment(8) = flexAlignCenterCenter
                .TextArray(9) = "페파본4":          .ColWidth(9) = 750:             .ColAlignment(9) = flexAlignCenterCenter
                .TextArray(10) = "밀도":            .ColWidth(10) = 750:            .ColAlignment(10) = flexAlignCenterCenter
                .TextArray(11) = "장력":            .ColWidth(11) = 750:            .ColAlignment(11) = flexAlignCenterCenter
                .TextArray(12) = "압력1":           .ColWidth(12) = 750:            .ColAlignment(12) = flexAlignCenterCenter
                .TextArray(13) = "압력2":           .ColWidth(13) = 750:            .ColAlignment(13) = flexAlignCenterCenter
                .TextArray(14) = "압력3":           .ColWidth(14) = 750:            .ColAlignment(14) = flexAlignCenterCenter
                            
            ' 샴푸
            Case PC_Shampu
                ' 작업조건 - 속도, 실측율
                .Cols = 7
                
                .TextArray(5) = "속도":                 .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
                .TextArray(6) = "실측율":               .ColWidth(6) = 800:             .ColAlignment(6) = flexAlignCenterCenter
        End Select
        
        Call FixedColAlignMentSetting(grdEdit)
        .Editable = flexEDKbdMouse
        .ExtendLastCol = False
        .Redraw = flexRDDirect
    
    End With
End Sub

Private Sub FixedColAlignMentSetting(vsGrid As VSFlexGrid)
    Dim iCount As Integer
    For iCount = 0 To vsGrid.Cols - 1
        vsGrid.FixedAlignment(iCount) = flexAlignCenterCenter
    Next iCount
End Sub

Private Sub cmdSave_Click()
'''    Dim NewResult As PlusLib2.TProcessResult
'''    Dim oProcess As PlusLib2.CProcess
'''    Dim nProcess%, sCard$
'''
'''
'''    On Error GoTo ErrHandler
'''
'''    nProcess = grdData.TextMatrix(grdData.Row, 7)
'''    ' 텐터, 건조, 수세, 모소, m/c, cpb전처리,  peach, 샴푸
'''
'''
'''    NewResult.sCardID = MakeCardID(grdEdit.TextMatrix(1, 1), OM_REDUCE)
'''    NewResult.nWorkSeq = grdEdit.TextMatrix(1, 3)
'''    NewResult.sProcessID = Format(nProcess, "0000")
'''    NewResult.sSplitID = grdEdit.TextMatrix(1, 4)
'''
'''    With grdEdit
'''
'''        Select Case nProcess
'''            ' 텐터 - '셋팅 '폭줄  '가공
'''            Case PC_Setting, PC_WidthLine, PC_FinalSetting
'''                ' 작업조건 - 온도, 속도, OverFeed,  위사밀도, Setting, 작업조건, 건조정도, 불량내역, 불량내역코드
'''
'''                NewResult.nTemper = .TextMatrix(1, 5)
'''                NewResult.nVelocity = .TextMatrix(1, 6)
'''                NewResult.nOverFeed = .TextMatrix(1, 7)
'''                NewResult.nDensity = .TextMatrix(1, 8)
'''                NewResult.sSettingClss = " "
'''                NewResult.sWorkCon = .TextMatrix(1, 9)
'''                NewResult.sDryID = " "
'''                NewResult.nWRQty = .TextMatrix(1, 10)
'''                NewResult.nWRPrice = .TextMatrix(1, 11)
'''
'''            ' 건조
'''            Case PC_Dry
'''                ' 작업조건 - 온도, 속도, OverFeed, 건조정도, 불량내역, 불량내역코드
'''
'''                NewResult.nTemper = .TextMatrix(1, 5)
'''                NewResult.nVelocity = .TextMatrix(1, 6)
'''                NewResult.nOverFeed = .TextMatrix(1, 7)
'''                NewResult.nDensity = .TextMatrix(1, 8)
'''                NewResult.sDryID = " "
'''
'''            ' 수세
'''            Case PC_REFINE, PC_SK
'''                ' 작업조건 - 온도, 속도, 정련구분, 셋팅
'''
'''                NewResult.nTemper = .TextMatrix(1, 5)
'''                NewResult.nVelocity = .TextMatrix(1, 6)
'''                NewResult.nDensity = .TextMatrix(1, 7)
'''
'''
'''
'''            ' 모소
'''            Case PC_Moso  ' 모소
'''                ' 작업조건 - 단면/양면구분, 풍량, 가스량, 속도
'''
'''                NewResult.sSideClss = .TextMatrix(1, 5)
'''                NewResult.nWind = .TextMatrix(1, 6)
'''                NewResult.nGas = .TextMatrix(1, 7)
'''                NewResult.nVelocity = .TextMatrix(1, 8)
'''
'''
'''            ' m/c
'''            Case PC_SK, PC_NewST, PC_OBoxSK
'''                ' 작업조건 - RPM,온도, 조제명, 조제코드, 조제농도
'''
'''                NewResult.nRPM = .TextMatrix(1, 5)
'''                NewResult.nTemper = .TextMatrix(1, 6)
'''                NewResult.sDyeAuxID = .TextMatrix(1, 8)
'''                NewResult.nDensity = .TextMatrix(1, 9)
'''
'''            ' CPBPre - 전처리, 1차호발, 정련, 1차정련, 2차정련, 2차감량, LBOX 전처리, CPB 전처리, 신 ST 전처리
'''            Case PC_Pre, PC_1stHobal, PC_Pufiry, PC_1stPurify, PC_2ndPurify, PC_2ndReduce, PC_LBoxPre, PC_CPBPre, PC_NewSTPre
'''                ' 작업조건 - 속도, 정련구분
'''
'''                NewResult.nVelocity = .TextMatrix(1, 5)
'''                NewResult.nDensity = 0
'''                NewResult.nBaseTemp = .TextMatrix(1, 6)
'''                NewResult.nAgingTemp = .TextMatrix(1, 7)
'''
'''            ' Peach
'''            Case PC_Peach
'''                ' 작업조건 - 폭, 속도, 페파본1, 페파본2, 페파본3, 밀도, 장력, 압력1, 압력2, 압력 3
'''
'''                NewResult.nVelocity = .TextMatrix(1, 5)
'''                NewResult.nPepaBon1 = .TextMatrix(1, 6)
'''                NewResult.nPepaBon2 = .TextMatrix(1, 7)
'''                NewResult.nPepaBon3 = .TextMatrix(1, 8)
'''                NewResult.nPepaBon4 = .TextMatrix(1, 9)
'''                NewResult.nDensity = .TextMatrix(1, 10)
'''                NewResult.nPressure1 = .TextMatrix(1, 11)
'''                NewResult.nPressure2 = .TextMatrix(1, 12)
'''                NewResult.nPressure3 = .TextMatrix(1, 13)
'''
'''            ' 샴푸
'''            Case PC_Shampu
'''                ' 작업조건 - 속도, 실측율
'''
'''                NewResult.nVelocity = .TextMatrix(1, 5)
'''                NewResult.nRealLoss = .TextMatrix(1, 6)
'''
'''        End Select
'''    End With
'''
'''
'''    Set oProcess = New PlusLib2.CProcess
'''    oProcess.Connection = g_adoCon
'''
'''    If oProcess.UpdateProcessResult(NewResult) Then
'''        MessageBox "변경내역이 저장되었습니다"
'''    End If
'''
'''    Set oProcess = Nothing
'''
'''    pnlData.Visible = False
'''    Call FillGridData
'''
'''    Exit Sub
'''
'''ErrHandler:
'''    Set oProcess = Nothing
'''    Screen.MousePointer = vbDefault
'''
'''    Call ErrorBox(Err.Number, "frmProcessResultModify.cmdSave_Click", Err.Description)
'''
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdEdit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With grdEdit
        If Not IsNumeric(.TextMatrix(Row, Col)) Then
            .TextMatrix(Row, Col) = "0"
        End If
    
    End With
End Sub

Private Sub grdEdit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If grdEdit.Col < 5 Then Exit Sub
    
End Sub


Private Sub grdEdit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

    If grdEdit.Col < 5 Then Exit Sub
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



Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15300, 9660

    Call SetOperate(Me)
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    
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
    
    Call InitGrid
    
    i = ModifyGrid

    Show

End Sub


Private Sub cmdSearch_Click()
    Call FillGridData

End Sub

Private Sub cboSearch_Click(Index As Integer)
    If m_bloading Then Exit Sub
    
    If Index = 1 Or Index = 4 Then Exit Sub

    If Index = 0 Then
        Call MakeMachineCombo

        Call FillGridData
    ElseIf Index = 3 Then
        Call MakeMachineNOCombo
        
        Call FillGridData
    
    Else
        If cboSearch(Index).ListIndex > 0 Then chkSearch(Index) = vbChecked
    End If
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 1 Or Index = 2 Then
        If chkSearch(Index) = vbUnchecked Then cboSearch(Index).ListIndex = 0
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




Private Sub optOrder_Click(Index As Integer)
    Dim iPrevCol%, iCol%, nSize%
    
    If Index = 0 Then
        chkSearch(3).Caption = "Order No."
    Else
        chkSearch(3).Caption = "관리 번호"
    End If

End Sub



Private Sub cmdExit_Click()
    PlusMDI.pnlMenu.Visible = True
    Unload Me
End Sub



Private Sub FillGridData()
    Dim oProcess As PlusLib2.CProcess
    Dim rs As Recordset
    Dim i%, iNowRow%, iProcess%
    Dim nFlag%, sFlag$
    Dim sWorkUnitID$, nWorkUnitSeq%
    Dim nRollCount&, nRollQty&
    Dim nReworkRoll&, nReworkQty&
    Dim nTotalRoll&, nTotalQty As Long, nWorkRoll%, nWorkQty As Long
    Dim sDate$, eDate$, sProcessID As EPROCESSCODE
    Dim nChkMachineID%, sMachineID$
    Dim nChkTeamID%, sTeamID$
    Dim nChkOrder%, sOrder$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim sTeam$, nClss%, sCard$
    Dim sCardID$, sTemp$, sSplitID$
    Dim bChange As Boolean, nColorSeq%
    

    Screen.MousePointer = vbHourglass

    iProcess = ModifyGrid()
    
    pnlCaption(2).Enabled = True
    cboSearch(2).Enabled = True

    On Error GoTo ErrHandler
    
    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon

    
    m_bSkip = True
    ' 공정별, 설비별 검색 구분
    nClss = IIf(optProcess(0).Value = True, 4, 1)

    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    sProcessID = Format(iProcess, "0000")
    
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

    '-----------------------------------------------------------
    ' 공정 카드로 검색
    sCardID = Left(txtCardID, 8)
    sTemp = Trim(Mid(txtCardID, 9, Len(txtCardID)))
    sSplitID = IIf(Len(sTemp) = 0, " ", sTemp)

    nColorSeq = 1


    With grdData
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        If chkCardSearch.Value = vbChecked Then
            Set rs = oProcess.GetResultByCard(sCardID, sSplitID)
        
        Else
            If optProcess(0).Value = True Then
                ' 설비별 검색
                Set rs = oProcess.GetResultByPlant(sDate, eDate, sProcessID, nChkMachineID, sMachineID, nChkTeamID, sTeamID, _
                                nChkOrder, sOrder, nChkCustom, sCustom, nChkArticle, sArticle)
            Else
                ' 공정별 검색
                Set rs = oProcess.GetResultByProcess(sDate, eDate, sProcessID, nChkMachineID, sMachineID, nChkTeamID, sTeamID, _
                                nChkOrder, sOrder, nChkCustom, sCustom, nChkArticle, sArticle)
            End If
        End If
        
        
        Set oProcess = Nothing

        ' 텐터, 건조, 수세, 모소, m/c, cpb전처리,  peach, 샴푸
        For i = 1 To rs.RecordCount
            If sWorkUnitID = rs!WorkUnitId Then
                nWorkUnitSeq = nWorkUnitSeq + 1
            
                bChange = False
            Else
                nWorkUnitSeq = 1
                sWorkUnitID = rs!WorkUnitId
                
                bChange = True
            End If
            
            sCard = MakeCardID(rs!CardID, OM_EXPAND)
            sCard = sCard & IIf(Len(Trim(rs!SplitID)) = 0, "", " (" & rs!SplitID & ")")

            Select Case rs!TeamID
                Case 1
                    sTeam = "A"
                Case 2
                    sTeam = "B"
                Case Is = 3
                    sTeam = "C"
            End Select
            
            ' *****************************************************************************
            ' *    카드별 실적
            ' *
            ' *     변경일 2003-12-01
            ' *     날씨맑음....
            ' ******************************************************************************
            If chkCardSearch.Value = vbChecked Then
                .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                        rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                        rs!WorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                        rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                        " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                        " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                        MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                        rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                        rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " "
            Else
                
                ' 설비 및 공정별 실적
                Select Case iProcess
                
                    ' Airo, 카렌다
                    Case PC_Airo, PC_Calender
                        .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " "
                                    
                    ' CPB 염색
                    Case PC_CPB
                       .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Winding & vbTab & rs!Vinyl
                
                    ' 셋팅, 가공, 폭줄,
                    Case PC_Setting, PC_WidthLine, PC_FinalSetting
                       .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Temper & vbTab & rs!Velocity & vbTab & rs!OverFeed & vbTab & rs!Density & vbTab & _
                                    rs!WorkCond & vbTab & CheckNull(rs!WRQty) & vbTab & CheckNull(rs!WRPrice) & vbTab & CheckNull(rs!HoldReason) & vbTab & CheckNull(rs!CodeID)
                
                    '-------------------------------------------------------------------------------------------------------------
                    ' 건조
                    Case PC_Dry
                        ' 작업조건 - 온도, 속도, OverFeed, 건조정도, 불량내역, 불량내역코드
                        .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Temper & vbTab & rs!Velocity & vbTab & rs!OverFeed & vbTab & rs!WorkDensity & vbTab & CheckNull(rs!HoldReason)
                                
                    '-------------------------------------------------------------------------------------------------------------
                    ' 수세
                    Case PC_REFINE, PC_SK
                        ' 작업조건 - 온도, 속도, OverFeed, 건조정도, 불량내역, 불량내역코드
                         .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Temper & vbTab & rs!Velocity & vbTab & rs!WorkDensity
                        
                    '-------------------------------------------------------------------------------------------------------------
                    ' 모소
                    Case PC_Moso  ' 모소
                        ' 작업조건 - 단면/양면구분, 풍량, 가스량, 속도
                        .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!SideClss & vbTab & rs!Wind & vbTab & rs!Gas & vbTab & rs!Velocity
                                 
                    '-------------------------------------------------------------------------------------------------------------
                    ' m/c
                    Case PC_SK  ' M/C 일지
                        ' 작업조건 - RPM,온도, 조제명, 조제코드, 조제농도
                        .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Rpm & vbTab & rs!Temper & vbTab & rs!DyeAux & vbTab & rs!DyeAuxID & vbTab & rs!Density
        
                    '-------------------------------------------------------------------------------------------------------------
                    ' CPBPre - 전처리, 1차호발, 정련, 1차정련, 2차정련, 2차감량
                    Case PC_Pre, PC_1stHobal, PC_Pufiry, PC_1stPurify, PC_2ndPurify, PC_2ndReduce
                        ' 작업조건 - 속도, 정련구분
                        .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Velocity & vbTab & rs!BaseTemp & vbTab & rs!AgingTemp
        
                    '-------------------------------------------------------------------------------------------------------------
                    ' Peach
                    Case PC_Peach
                        ' 작업조건 - 폭, 속도, 페파본1, 페파본2, 페파본3, 밀도, 장력, 압력1, 압력2, 압력 3
                       .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Velocity & vbTab & rs!PePaBon1 & vbTab & rs!PePaBon2 & vbTab & rs!PePaBon3 & vbTab & rs!PePaBon4 & vbTab & rs!Density & vbTab & _
                                    rs!Tension & vbTab & rs!Pressure1 & vbTab & rs!Pressure2 & vbTab & rs!Pressure3
        
                    '-------------------------------------------------------------------------------------------------------------
                    ' 샴푸
                    Case PC_Shampu
                        ' 작업조건 - 속도, 실측율
                        .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_LONG, rs!ResultDate) & vbTab & _
                                    rs!Process & vbTab & rs!MachineNO & vbTab & rs!ProcessID & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId & vbTab & _
                                    nWorkUnitSeq & vbTab & sCard & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                                    rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!WorkName & vbTab & _
                                    " " & vbTab & rs!workroll & vbTab & SetCurrency(rs!workqty) & vbTab & rs!UnitPrice & vbTab & rs!workqty * rs!UnitPrice & vbTab & _
                                    " " & vbTab & rs!PreWidth & vbTab & rs!OrderWidth & vbTab & rs!WorkWidth & vbTab & " " & vbTab & _
                                    MakeDate(DF_LONG, rs!StartDate) & vbTab & MakeTime(rs!StartTime) & vbTab & MakeTime(CheckNull(rs!EndTime)) & vbTab & _
                                    rs!requiredtime & vbTab & " " & vbTab & sTeam & vbTab & rs!Name & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!OrderSeq & vbTab & rs!CardID & vbTab & rs!SplitID & vbTab & rs!WorkSeq & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                                    rs!Velocity & vbTab & rs!RealLoss
                End Select
                
            End If


            '' 작업 단위별로 색상 변경
            If bChange Then
                nColorSeq = IIf(nColorSeq = 1, 2, 1)
            
            End If
            
            .Row = .FixedRows + i - 1
            .Col = .FixedCols
            .ColSel = .Cols - 1
            If nColorSeq = 1 Then
                .CellBackColor = &HE0E0E0
            Else
                .CellBackColor = &HFFFFFF
            End If
            
            nTotalRoll = nTotalRoll + CheckNull(rs!workroll)
            nTotalQty = nTotalQty + CheckNull(rs!workqty)

            If Trim(rs!ReWorkClss) = "*" Then
                nReworkRoll = nReworkRoll + CheckNull(rs!workroll)
                nReworkQty = nReworkQty + CheckNull(rs!workqty)
            Else
                nWorkRoll = nWorkRoll + CheckNull(rs!workroll)
                nWorkQty = nWorkQty + CheckNull(rs!workqty)
            End If



            rs.MoveNext
        Next i

        
        .Redraw = flexRDDirect
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
            .HighLight = flexHighlightAlways
        Else
            .HighLight = flexHighlightNever

           MsgBox LoadResString(203), vbInformation
        End If

        .SetFocus
    End With
    
    rs.Close
    Set rs = Nothing

    m_bSkip = False


    With grdSum
        .RowHeightMin = 300
        .Rows = 0
        .AddItem ""
        .TextMatrix(0, 0) = "총 생산"
        .TextMatrix(0, 1) = SetCurrency(nTotalRoll) & "  절"
        .TextMatrix(0, 2) = SetCurrency(nTotalQty) & "  YDS"
        .Cell(flexcpFontSize, 0, 1, 0, 2) = 12
        .Cell(flexcpFontBold, 0, 1, 0, 2) = True
        
        .AddItem ""
        .TextMatrix(1, 0) = " 생 산"
        .TextMatrix(1, 1) = SetCurrency(nWorkRoll) & "  절"
        .TextMatrix(1, 2) = SetCurrency(nWorkQty) & "  YDS"
        .Cell(flexcpFontSize, 1, 1, 1, 2) = 12
        .Cell(flexcpFontBold, 1, 1, 1, 2) = True
        
        .AddItem ""
        .TextMatrix(2, 0) = " 수 정"
        .TextMatrix(2, 1) = SetCurrency(nReworkRoll) & "  절"
        .TextMatrix(2, 2) = SetCurrency(nReworkQty) & "  YDS"
        .Cell(flexcpFontSize, 2, 1, 2, 2) = 12
        .Cell(flexcpFontBold, 2, 1, 2, 2) = True
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
    Set oProcess = Nothing
    Screen.MousePointer = vbDefault
    m_bSkip = False
    
    Call ErrorBox(Err.Number, "frmProcessResultView.FillGridData", Err.Description)
End Sub



Private Function MakeTime(ByVal sTime As String) As String

    If Len(sTime) = 0 Then
        MakeTime = ":"
    Else
        MakeTime = Left(sTime, 2) & ":" & Right(sTime, 2)
    End If
    
End Function



Private Sub MakeProcessCombo()
    Dim oProcess As PlusLib2.CProcess
    Dim rs As Recordset
        

    Screen.MousePointer = vbHourglass
    m_bloading = True

    On Error GoTo ErrHandler

    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon

    Set rs = oProcess.GetWorkProcess
    
    Set oProcess = Nothing

    With cboSearch(0)
        .Clear

        Do Until rs.EOF
            If rs!ProcessID < "4304" Or rs!ProcessID > "4306" Then
                .AddItem CStr(rs!Process)
                .ItemData(.NewIndex) = CLng(rs!ProcessID)
            End If

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
    Set oProcess = Nothing
    m_bloading = False
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmProcessResultModify.MakeProcessCombo", Err.Description)
End Sub

Private Sub MakePlantCombo()
    Dim oProcess As PlusLib2.CProcess
    Dim rs As Recordset
        

    Screen.MousePointer = vbHourglass
    m_bloading = True

    On Error GoTo ErrHandler

    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon

    Set rs = oProcess.GetPlant
    Set oProcess = Nothing

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
    Set oProcess = Nothing
    m_bloading = False
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmProcessResultView.MakePlantCombo", Err.Description)
End Sub


Private Sub MakeMachineCombo()
    Dim oProcess As PlusLib2.CProcess
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon

    Set rs = oProcess.GetMachine(Format(cboSearch(0).ItemData(cboSearch(0).ListIndex), "0000"))
    Set oProcess = Nothing

    With cboSearch(1)
        .Clear

        .AddItem "전체"
        .ItemData(.NewIndex) = 0
        Do Until rs.EOF
            .AddItem rs!MachineNO & "호기"
            .ItemData(.NewIndex) = CLng(rs!machineid)

            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing

        .ListIndex = 0
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oProcess = Nothing
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmProcessResultView.MakeMachineCombo", Err.Description)
End Sub



Private Sub MakeMachineNOCombo()
    Dim oProcess As PlusLib2.CProcess
    Dim rs As Recordset
    Dim sPlant$, i%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon

    sPlant = cboSearch(3).Text
    
    Set rs = oProcess.GetMachineByPlant(sPlant)
    Set oProcess = Nothing

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

        .ListIndex = 0
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oProcess = Nothing
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "frmProcessResultView.MakeMachineNOCombo", Err.Description)
End Sub



Private Function ModifyGrid() As Integer
    Dim i%
    Dim nProcess As EPROCESSCODE
    
    ' 설비별 검색
    If optProcess(0).Value = True Then
        ModifyGrid = cboSearch(3).ItemData(cboSearch(3).ListIndex)
        nProcess = cboSearch(3).ItemData(cboSearch(3).ListIndex)
    
    Else    ' 공정별 검색
        ModifyGrid = cboSearch(0).ItemData(cboSearch(0).ListIndex)
        nProcess = cboSearch(0).ItemData(cboSearch(0).ListIndex)
        ' 텐터, 건조, 수세, 모소, m/c, cpb전처리,  peach, 샴푸
    End If
    
    Call SetVSFlexGrid(grdData)
    
    With grdData
        .Cols = 57
        .Rows = 5
        .FixedRows = 5
        ' 0~2번 Row는 리포트 발행시 타이틀및 일자등 출력하는 부분
        ' 3,4번 Row는 실제 화면에서 컬럼명 출력부분
        
        For i = 0 To 3
            .RowHeight(i) = 300
        Next i
        
        .RowHeight(4) = 400
        
        .RowHeightMin = 300
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        ' 기본내역
        .TextMatrix(3, 0) = "NO"
        .TextMatrix(3, 1) = " ":                        .ColWidth(1) = 0
        .TextMatrix(3, 2) = " ":                        .ColWidth(2) = 0
        .TextMatrix(3, 3) = "구" & vbCrLf & "분":       .ColWidth(3) = 300:             .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "실적일자":                 .ColWidth(4) = 600:             .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "공정명":                   .ColWidth(5) = 900:             .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "기계" & vbCrLf & "NO":     .ColWidth(6) = 400:             .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(3, 7) = "ProcessID":                .ColWidth(7) = 0
        .TextMatrix(3, 8) = "밧자" & vbCrLf & "NO":     .ColWidth(8) = 600:            .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(3, 9) = "작업" & vbCrLf & "단위":   .ColWidth(9) = 0:               .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(3, 10) = "단위" & vbCrLf & "순위":  .ColWidth(10) = 400:            .ColAlignment(10) = flexAlignCenterCenter
        .TextMatrix(3, 11) = "  CardNO":                .ColWidth(11) = 1500:           .ColAlignment(11) = flexAlignLeftCenter
        .TextMatrix(3, 12) = "거래처":                  .ColWidth(12) = 920:           .ColAlignment(12) = flexAlignLeftCenter
        .TextMatrix(3, 13) = "품명":                    .ColWidth(13) = 2100:           .ColAlignment(13) = flexAlignLeftCenter
        .TextMatrix(3, 14) = "OrderNo":                 .ColWidth(14) = 1200:           .ColAlignment(14) = flexAlignLeftCenter
        .TextMatrix(3, 15) = "관리번호":                .ColWidth(15) = 1200:           .ColAlignment(15) = flexAlignCenterCenter
        .TextMatrix(3, 16) = "색상명":                  .ColWidth(16) = 1600:           .ColAlignment(16) = flexAlignLeftCenter
        .TextMatrix(3, 17) = "가공" & vbCrLf & "방법":  .ColWidth(17) = 1000:           .ColAlignment(17) = flexAlignCenterCenter
        .TextMatrix(3, 18) = " ":                       .ColWidth(18) = 0
        
        ' 절수, 수량
        .TextMatrix(3, 19) = "작업량":                  .ColWidth(19) = 600:            .ColAlignment(19) = flexAlignRightCenter
        .TextMatrix(3, 20) = "작업량":                  .ColWidth(20) = 800:            .ColAlignment(20) = flexAlignRightCenter
        .TextMatrix(3, 21) = "단가":                    .ColWidth(21) = 0:            .ColAlignment(21) = flexAlignRightCenter
        .TextMatrix(3, 22) = "금액":                    .ColWidth(22) = 0:            .ColAlignment(22) = flexAlignRightCenter
        .TextMatrix(3, 23) = "":                        .ColWidth(23) = 0
        
        ' 작업전 폭, 요구, 작업후 폭
        .TextMatrix(3, 24) = "폭":                      .ColWidth(24) = 700:            .ColAlignment(24) = flexAlignCenterCenter
        .TextMatrix(3, 25) = "폭":                      .ColWidth(25) = 700:            .ColAlignment(25) = flexAlignCenterCenter
        .TextMatrix(3, 26) = "폭":                      .ColWidth(26) = 700:            .ColAlignment(26) = flexAlignCenterCenter
        .TextMatrix(3, 27) = " ":                       .ColWidth(27) = 0
        
        
        ' 시작, 종료, 소요시간
        .TextMatrix(3, 28) = "작업일자":                .ColWidth(28) = 1000:           .ColAlignment(28) = flexAlignCenterCenter
        .TextMatrix(3, 29) = "작업시간":                .ColWidth(29) = 700:            .ColAlignment(29) = flexAlignCenterCenter
        .TextMatrix(3, 30) = "작업시간":                .ColWidth(30) = 700:            .ColAlignment(30) = flexAlignCenterCenter
        .TextMatrix(3, 31) = "작업시간":                .ColWidth(31) = 700:            .ColAlignment(31) = flexAlignCenterCenter
        .TextMatrix(3, 32) = " ":                       .ColWidth(32) = 0
        
        .TextMatrix(3, 33) = "작업조":                  .ColWidth(33) = 700:            .ColAlignment(33) = flexAlignCenterCenter
        .TextMatrix(3, 34) = "작업자":                  .ColWidth(34) = 700:            .ColAlignment(34) = flexAlignCenterCenter
        .TextMatrix(3, 35) = " ":                       .ColWidth(35) = 0
        
        .TextMatrix(3, 36) = "Alter":                   .ColWidth(36) = 0
        .TextMatrix(3, 37) = "ColorID":                 .ColWidth(37) = 0
        .TextMatrix(3, 38) = "CardID":                  .ColWidth(38) = 0
        .TextMatrix(3, 39) = "SplitID":                 .ColWidth(39) = 0
        .TextMatrix(3, 40) = "WorkSeq":                 .ColWidth(40) = 0
        .TextMatrix(3, 41) = " ":                       .ColWidth(41) = 0
        .TextMatrix(3, 42) = " ":                       .ColWidth(42) = 0
        .TextMatrix(3, 43) = " ":                       .ColWidth(43) = 0
        
        .TextMatrix(3, 56) = "비고":                    .ColWidth(56) = 650:            .ColAlignment(56) = flexAlignCenterCenter
        
        '///////////////////////////////////////////////////
        
        .TextMatrix(4, 0) = "NO"
        .TextMatrix(4, 1) = " "
        .TextMatrix(4, 2) = " "
        .TextMatrix(4, 3) = "구" & vbCrLf & "분"
        .TextMatrix(4, 4) = "실적일자"
        .TextMatrix(4, 5) = "공정명"
        .TextMatrix(4, 6) = "기계" & vbCrLf & "NO"
        .TextMatrix(4, 7) = ""
        .TextMatrix(4, 8) = "밧자" & vbCrLf & "NO"
        .TextMatrix(4, 9) = "작업" & vbCrLf & "단위"
        .TextMatrix(4, 10) = "단위" & vbCrLf & "순위"
        .TextMatrix(4, 11) = "  CardNO"
        .TextMatrix(4, 12) = "거래처"
        .TextMatrix(4, 13) = "품명"
        .TextMatrix(4, 14) = "OrderNo"
        .TextMatrix(4, 15) = "관리번호"
        .TextMatrix(4, 16) = "색상명"
        .TextMatrix(4, 17) = "가공" & vbCrLf & "방법"
        .TextMatrix(4, 18) = " "
        
        ' 절수, 수량
        .TextMatrix(4, 19) = "절수"
        .TextMatrix(4, 20) = "수량"
        .TextMatrix(4, 21) = "단가"
        .TextMatrix(4, 22) = "금액"
        .TextMatrix(4, 23) = " "
        
        ' 작업전 폭, 요구, 작업후 폭
        .TextMatrix(4, 24) = "작업전"
        .TextMatrix(4, 25) = "요구"
        .TextMatrix(4, 26) = "작업후"
        .TextMatrix(4, 27) = " "
        
        ' 시작, 종료, 소요시간
        .TextMatrix(4, 28) = "작업일자"
        .TextMatrix(4, 29) = "시작"
        .TextMatrix(4, 30) = "종료"
        .TextMatrix(4, 31) = "소요"
        .TextMatrix(4, 32) = " "
        
        .TextMatrix(4, 33) = "작업조"
        .TextMatrix(4, 34) = "작업자"
        .TextMatrix(4, 35) = " "
        
        .TextMatrix(4, 36) = "Alter"
        .TextMatrix(4, 37) = "ColorID"
        .TextMatrix(4, 38) = "SplitID"
        .TextMatrix(4, 39) = "WorkSeq"
        .TextMatrix(4, 40) = " "
        .TextMatrix(4, 41) = " "
        .TextMatrix(4, 42) = " "
        .TextMatrix(4, 43) = " "
       
        .TextMatrix(4, 56) = "비고"
    
        ' ******************************************************************
        ' *    카드별 공정실적 검색
        ' *
        ' *     변경일자 2003-12-01
        ' ****************************************************************&
        If chkCardSearch.Value = vbChecked Then
            
            .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 0
            .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 0
            .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 0
            .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
            .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
            .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
            .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
            .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
            .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
            .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
            .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
            .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                      
            ' 작업조건
            .TextMatrix(4, 44) = ""
            .TextMatrix(4, 45) = ""
            .TextMatrix(4, 46) = ""
            .TextMatrix(4, 47) = ""
            .TextMatrix(4, 48) = ""
            .TextMatrix(4, 49) = ""
            .TextMatrix(4, 50) = ""
            .TextMatrix(4, 51) = ""
            .TextMatrix(4, 52) = ""
            .TextMatrix(4, 53) = ""
            .TextMatrix(4, 54) = ""
            .TextMatrix(4, 55) = ""
        
        Else
        
            ' ******************************************************************
            ' *    ' 공정별, 설비별 실적 검색
            ' *
            ' *     변경일자 2003-12-01
            ' ****************************************************************&
            Select Case nProcess
            
                Case PC_Airo, PC_Calender
                     ' 작업조건 - 없음
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 0
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 0
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 0
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                              
                    ' 작업조건
                    .TextMatrix(4, 44) = ""
                    .TextMatrix(4, 45) = ""
                    .TextMatrix(4, 46) = ""
                    .TextMatrix(4, 47) = ""
                    .TextMatrix(4, 48) = ""
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                    
                 
                Case PC_CPB
                     ' 작업조건 - 없음
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 0
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                              
                    ' 작업조건
                    .TextMatrix(4, 44) = "와인딩"
                    .TextMatrix(4, 45) = "비닐"
                    .TextMatrix(4, 46) = ""
                    .TextMatrix(4, 47) = ""
                    .TextMatrix(4, 48) = ""
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                    
                ' 텐터 - '셋팅 '폭줄  '가공
                Case PC_Setting, PC_WidthLine, PC_FinalSetting
                    ' 작업조건 - 온도, 속도, OverFeed,  위사밀도, Setting, 작업조건, 건조정도, 불량내역, 불량내역코드
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 900:            .ColAlignment(47) = flexAlignCenterCenter
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 900:            .ColAlignment(48) = flexAlignCenterCenter
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 900:            .ColAlignment(49) = flexAlignCenterCenter
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 900:            .ColAlignment(50) = flexAlignCenterCenter
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 900:            .ColAlignment(51) = flexAlignCenterCenter
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                              
                    ' 작업조건
                    .TextMatrix(4, 44) = "온도(℃)"
                    .TextMatrix(4, 45) = "속도(M)"
                    .TextMatrix(4, 46) = "Over" & vbCrLf & "Feed(%)"
                    .TextMatrix(4, 47) = "위사" & vbCrLf & "밀도(T)"
                    .TextMatrix(4, 48) = "작업구분"
                    .TextMatrix(4, 49) = "WR사용" & vbCrLf & "량(kg)"
                    .TextMatrix(4, 50) = "WR사용" & vbCrLf & "금액(원)"
                    .TextMatrix(4, 51) = "불량" & vbCrLf & "내역"
                    .TextMatrix(4, 52) = "불량내역 코드"
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                    
                '-------------------------------------------------------------------------------------------------------------
                ' 건조
                Case PC_Dry
                    ' 작업조건 - 온도, 속도, OverFeed, 건조정도, 불량내역, 불량내역코드
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 900:            .ColAlignment(47) = flexAlignCenterCenter
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 900:            .ColAlignment(48) = flexAlignCenterCenter
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                            
                    ' 작업조건
                     .TextMatrix(4, 44) = "온도(℃)"
                     .TextMatrix(4, 45) = "속도(M)"
                     .TextMatrix(4, 46) = "Over" & vbCrLf & "Feed(%)"
                     .TextMatrix(4, 47) = "밀도(T)"
                     .TextMatrix(4, 48) = "불량" & vbCrLf & "내역"
                     .TextMatrix(4, 49) = "불량내역코드"
                     .TextMatrix(4, 50) = ""
                     .TextMatrix(4, 51) = ""
                     .TextMatrix(4, 52) = ""
                     .TextMatrix(4, 53) = ""
                     .TextMatrix(4, 54) = ""
                     .TextMatrix(4, 55) = ""
                
                
                '-------------------------------------------------------------------------------------------------------------
                ' 수세
                Case PC_REFINE, PC_SK
                    ' 작업조건 - 온도, 속도, 정련구분, 셋팅
                    
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                              
                    ' 작업조건
                    .TextMatrix(4, 44) = "온도(℃)"
                    .TextMatrix(4, 45) = "속도(M)"
                    .TextMatrix(4, 46) = "밀도(T)"
                    .TextMatrix(4, 47) = ""
                    .TextMatrix(4, 48) = ""
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                    
                '-------------------------------------------------------------------------------------------------------------
                ' 모소
                Case PC_Moso  ' 모소
                    ' 작업조건 - 단면/양면구분, 풍량, 가스량, 속도
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 900:            .ColAlignment(47) = flexAlignCenterCenter
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                    
                    ' 작업조건
                    .TextMatrix(4, 44) = "단면/" & vbCrLf & "양면구분"
                    .TextMatrix(4, 45) = "풍량"
                    .TextMatrix(4, 46) = "가스량"
                    .TextMatrix(4, 47) = "속도(M)"
                    .TextMatrix(4, 48) = ""
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                              
                
                '-------------------------------------------------------------------------------------------------------------
                ' m/c
                Case PC_SK, PC_NewST, PC_OBoxSK
                    ' 작업조건 - RPM,온도, 조제명, 조제코드, 조제농도
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 900:            .ColAlignment(48) = flexAlignCenterCenter
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                    
                              
                    ' 작업조건
                    .TextMatrix(4, 44) = "RPM"
                    .TextMatrix(4, 45) = "온도(℃)"
                    .TextMatrix(4, 46) = "조제명"
                    .TextMatrix(4, 47) = "조제" & vbCrLf & "코드"
                    .TextMatrix(4, 48) = "조제" & vbCrLf & "농도"
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                                    
                
                '-------------------------------------------------------------------------------------------------------------
                ' CPBPre - 전처리, 1차호발, 정련, 1차정련, 2차정련, 2차감량, LBOX 전처리, CPB 전처리, 신 ST 전처리
                Case PC_Pre, PC_1stHobal, PC_Pufiry, PC_1stPurify, PC_2ndPurify, PC_2ndReduce, PC_LBoxPre, PC_CPBPre, PC_NewSTPre
                    ' 작업조건 - 속도, 정련구분
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                    
                    
                    ' 작업조건
                    .TextMatrix(4, 44) = "속도(M)"
                    .TextMatrix(4, 45) = "베이스" & vbCrLf & "온도(℃)"
                    .TextMatrix(4, 46) = "NaOH" & vbCrLf & "량(ℓ/)"
                    .TextMatrix(4, 47) = ""
                    .TextMatrix(4, 48) = ""
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                                                              
                
                '-------------------------------------------------------------------------------------------------------------
                ' Peach
                Case PC_Peach
                    ' 작업조건 - 폭, 속도, 페파본1, 페파본2, 페파본3, 밀도, 장력, 압력1, 압력2, 압력 3
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 900:            .ColAlignment(46) = flexAlignCenterCenter
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 900:            .ColAlignment(47) = flexAlignCenterCenter
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 900:            .ColAlignment(48) = flexAlignCenterCenter
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 900:            .ColAlignment(49) = flexAlignCenterCenter
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 900:            .ColAlignment(50) = flexAlignCenterCenter
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 900:            .ColAlignment(51) = flexAlignCenterCenter
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 900:            .ColAlignment(52) = flexAlignCenterCenter
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 900:            .ColAlignment(53) = flexAlignCenterCenter
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                    
                              
                    ' 작업조건
                    .TextMatrix(4, 44) = "속도(M)"
                    .TextMatrix(4, 45) = "페파본" & vbCrLf & "1(本)"
                    .TextMatrix(4, 46) = "페파본" & vbCrLf & "2(本)"
                    .TextMatrix(4, 47) = "페파본" & vbCrLf & "3(本)"
                    .TextMatrix(4, 48) = "페파본" & vbCrLf & "4(本)"
                    .TextMatrix(4, 49) = "밀도(T)"
                    .TextMatrix(4, 50) = "장력" & vbCrLf & "(m/m)"
                    .TextMatrix(4, 51) = "압력1(K)"
                    .TextMatrix(4, 52) = "압력2(K)"
                    .TextMatrix(4, 53) = "압력3(K)"
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
                    
                                       
                '-------------------------------------------------------------------------------------------------------------
                ' 샴푸
                Case PC_Shampu
                    ' 작업조건 - 속도, 실측율
                    .TextMatrix(3, 44) = "작업조건":                .ColWidth(44) = 900:            .ColAlignment(44) = flexAlignCenterCenter
                    .TextMatrix(3, 45) = "작업조건":                .ColWidth(45) = 900:            .ColAlignment(45) = flexAlignCenterCenter
                    .TextMatrix(3, 46) = "작업조건":                .ColWidth(46) = 0
                    .TextMatrix(3, 47) = "작업조건":                .ColWidth(47) = 0
                    .TextMatrix(3, 48) = "작업조건":                .ColWidth(48) = 0
                    .TextMatrix(3, 49) = "작업조건":                .ColWidth(49) = 0
                    .TextMatrix(3, 50) = "작업조건":                .ColWidth(50) = 0
                    .TextMatrix(3, 51) = "작업조건":                .ColWidth(51) = 0
                    .TextMatrix(3, 52) = "작업조건":                .ColWidth(52) = 0
                    .TextMatrix(3, 53) = "작업조건":                .ColWidth(53) = 0
                    .TextMatrix(3, 54) = "작업조건":                .ColWidth(54) = 0
                    .TextMatrix(3, 55) = "작업조건":                .ColWidth(55) = 0
                    
                    
                    ' 작업조건
                    .TextMatrix(4, 44) = "속도(M)"
                    .TextMatrix(4, 45) = "실측율"
                    .TextMatrix(4, 46) = ""
                    .TextMatrix(4, 47) = ""
                    .TextMatrix(4, 48) = ""
                    .TextMatrix(4, 49) = ""
                    .TextMatrix(4, 50) = ""
                    .TextMatrix(4, 51) = ""
                    .TextMatrix(4, 52) = ""
                    .TextMatrix(4, 53) = ""
                    .TextMatrix(4, 54) = ""
                    .TextMatrix(4, 55) = ""
            End Select
        End If
        
        .MergeCells = flexMergeFixedOnly
        
        For i = 0 To 4
            .MergeRow(i) = True
        Next i
        
        For i = 0 To 56
            .MergeCol(i) = True
        Next i
        
        Call FixedColAlignMentSetting(grdData)
        .WordWrap = False
        .ExtendLastCol = True
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


Private Sub InitGrid()
    Dim iCount As Integer
    
    With grdSum
    
        .Redraw = flexRDNone
        
        .Rows = 1
        .FixedRows = 0
        .Cols = 3
        .FixedCols = 1
        
        .RowHeight(0) = 350
        .ColWidth(0) = 5000

        .ScrollBars = flexScrollBarNone
        .HighLight = flexHighlightNever
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False

        .RowHeightMin = 275
        .WordWrap = False
        .ExtendLastCol = True
        
        .ColAlignment(0) = flexAlignCenterCenter
        
        For iCount = 0 To .Cols - 1
            .FixedAlignment(iCount) = flexAlignCenterCenter
        Next iCount

        .Redraw = True
        
        .TextArray(0) = "합계"
        .TextArray(1) = "0 건":         .ColWidth(1) = 7000
        .TextArray(2) = "0 YDS"
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub optProcess_Click(Index As Integer)
    If optProcess(0).Value = True Then
        pnlProcess(1).Enabled = False
    Else
        pnlProcess(0).Enabled = False
    
    End If
    pnlProcess(Index).Enabled = optProcess(Index).Value

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

