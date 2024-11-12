VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInstCondition 
   Caption         =   "작업조건지시"
   ClientHeight    =   9285
   ClientLeft      =   2130
   ClientTop       =   2205
   ClientWidth     =   15240
   Icon            =   "frmInstCondition.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15240
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   330
      Left            =   15
      TabIndex        =   59
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
      Height          =   6015
      Left            =   15
      TabIndex        =   58
      Top             =   2040
      Width           =   3900
      _cx             =   6879
      _cy             =   10610
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
      Height          =   2070
      Left            =   30
      TabIndex        =   51
      Top             =   -45
      Width           =   3900
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   3390
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1305
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
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   45
         Top             =   1305
         Width           =   1905
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   2985
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   48
         ToolTipText     =   "자료 저장"
         Top             =   210
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   47
         Top             =   1680
         Width           =   1905
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   345
         MousePointer    =   99  '사용자 정의
         TabIndex        =   41
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   540
         Width           =   600
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   1005
         TabIndex        =   42
         Top             =   540
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   566231041
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1005
         TabIndex        =   43
         Top             =   900
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   566231041
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   1305
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
            Caption         =   "거 래 처"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   1680
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
            Index           =   2
            Left            =   60
            TabIndex        =   46
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   165
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
            Caption         =   "수주 일자"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   39
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   2295
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   615
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13575
      TabIndex        =   49
      Top             =   8490
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
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
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   9180
      Left            =   3990
      TabIndex        =   50
      Top             =   60
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   16193
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
      Begin VB.CommandButton cmdCopy 
         Caption         =   "복사"
         Height          =   795
         Left            =   10410
         TabIndex        =   110
         Top             =   60
         Width           =   795
      End
      Begin VB.TextBox txtSelRecs 
         Height          =   315
         Left            =   4320
         TabIndex        =   109
         Text            =   "10"
         Top             =   450
         Width           =   615
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   795
         Index           =   3
         Left            =   6420
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
         Left            =   8010
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
         Left            =   9600
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
         Left            =   8805
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
         Left            =   7215
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   70
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   765
         Left            =   2430
         TabIndex        =   63
         Top             =   60
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1349
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton opProcess 
            Caption         =   "공정별"
            Height          =   180
            Left            =   180
            TabIndex        =   65
            Top             =   450
            Width           =   1275
         End
         Begin VB.OptionButton opMachine 
            Caption         =   "설비별"
            Height          =   255
            Left            =   180
            TabIndex        =   64
            Top             =   120
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid GrdCondition 
         Height          =   3285
         Left            =   60
         TabIndex        =   66
         Top             =   900
         Width           =   11130
         _cx             =   19632
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
         ScrollBars      =   1
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
         Index           =   1
         Left            =   4320
         TabIndex        =   67
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "최근레코드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin TabDlg.SSTab tabProc 
         Height          =   4215
         Left            =   60
         TabIndex        =   69
         Top             =   4200
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   7435
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   8
         TabHeight       =   600
         TabCaption(0)   =   "  정련 C.P.B기  "
         TabPicture(0)   =   "frmInstCondition.frx":000C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "pnlCPB"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSPanel2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "  수세기  "
         TabPicture(1)   =   "frmInstCondition.frx":0028
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSPanel3(0)"
         Tab(1).Control(1)=   "pnlRefine"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "  텐터기  "
         TabPicture(2)   =   "frmInstCondition.frx":0044
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "pnlTenter"
         Tab(2).Control(1)=   "SSPanel3(1)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "  Peach기  "
         TabPicture(3)   =   "frmInstCondition.frx":0060
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "pnlPeach"
         Tab(3).Control(1)=   "SSPanel3(2)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "건조기  "
         TabPicture(4)   =   "frmInstCondition.frx":007C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "pnlDry"
         Tab(4).Control(1)=   "SSPanel3(3)"
         Tab(4).ControlCount=   2
         Begin Threed.SSPanel SSPanel3 
            Height          =   345
            Index           =   0
            Left            =   -73410
            TabIndex        =   126
            Top             =   780
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   609
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboRefineProc 
               Height          =   300
               Left            =   1260
               Style           =   2  '드롭다운 목록
               TabIndex        =   130
               Top             =   30
               Width           =   2775
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   2
               Left            =   30
               TabIndex        =   131
               Top             =   30
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
               Caption         =   "공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Left            =   210
            TabIndex        =   123
            Top             =   870
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   661
            _Version        =   196609
            Caption         =   "SSPanel2"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboCPBProc 
               Height          =   300
               Left            =   1425
               Style           =   2  '드롭다운 목록
               TabIndex        =   124
               Top             =   60
               Width           =   2775
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   0
               Left            =   180
               TabIndex        =   125
               Top             =   60
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
               Caption         =   "공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel pnlCPB 
            Height          =   2235
            Left            =   360
            TabIndex        =   75
            Top             =   1230
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   3942
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txtCPBPersonID 
               BackColor       =   &H00E0E0E0&
               Height          =   345
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   1410
               Width           =   1275
            End
            Begin VB.TextBox txtCPBVelocity 
               Height          =   315
               Left            =   1260
               TabIndex        =   0
               Top             =   30
               Width           =   1545
            End
            Begin VB.TextBox txtCPBRemark 
               Height          =   1005
               Left            =   1260
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   1
               Top             =   375
               Width           =   2775
            End
            Begin VB.ComboBox cboCPBRefineClss 
               Height          =   300
               Left            =   1260
               Style           =   2  '드롭다운 목록
               TabIndex        =   2
               Top             =   1815
               Visible         =   0   'False
               Width           =   2775
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   27
               Left            =   30
               TabIndex        =   76
               Top             =   30
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "속도(㎜)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   37
               Left            =   30
               TabIndex        =   77
               Top             =   1815
               Visible         =   0   'False
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
               Caption         =   "정련구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   345
               Index           =   41
               Left            =   30
               TabIndex        =   78
               Top             =   1410
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   609
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
               Caption         =   "작성자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   39
               Left            =   30
               TabIndex        =   79
               Top             =   375
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
               Caption         =   "비고 사항"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel pnlRefine 
            Height          =   2655
            Left            =   -73410
            TabIndex        =   80
            Top             =   1110
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   4683
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txtRefineTemper 
               Height          =   345
               Left            =   1260
               TabIndex        =   4
               Top             =   30
               Width           =   1485
            End
            Begin VB.TextBox txtRefineVelocity 
               Height          =   345
               Left            =   1260
               TabIndex        =   5
               Top             =   390
               Width           =   1485
            End
            Begin VB.TextBox txtRefinePersonID 
               BackColor       =   &H00E0E0E0&
               Height          =   345
               Left            =   1245
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   1785
               Width           =   1485
            End
            Begin VB.TextBox txtRefineSettingClss 
               Height          =   345
               Left            =   2550
               TabIndex        =   9
               Top             =   2760
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.ComboBox cboRefineClss 
               Height          =   300
               Left            =   1245
               Style           =   2  '드롭다운 목록
               TabIndex        =   7
               Top             =   2160
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.TextBox txtRefineRemark 
               Height          =   1005
               Left            =   1260
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   6
               Top             =   765
               Width           =   2775
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   46
               Left            =   15
               TabIndex        =   81
               Top             =   2160
               Visible         =   0   'False
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
               Caption         =   "정련구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   48
               Left            =   30
               TabIndex        =   82
               Top             =   765
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
               Caption         =   "비고 사항"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   50
               Left            =   480
               TabIndex        =   83
               Top             =   2700
               Visible         =   0   'False
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
               Caption         =   "Setting구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   345
               Index           =   42
               Left            =   30
               TabIndex        =   111
               Top             =   390
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   609
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
               Caption         =   "속도(㎜)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   345
               Index           =   47
               Left            =   15
               TabIndex        =   112
               Top             =   1785
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   609
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
               Caption         =   "작성자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   345
               Index           =   49
               Left            =   30
               TabIndex        =   113
               Top             =   30
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   609
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
               Caption         =   "온도(℃)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel pnlTenter 
            Height          =   2535
            Left            =   -72450
            TabIndex        =   84
            Top             =   1170
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   4471
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboTenterCodeID 
               Height          =   300
               Left            =   5610
               Style           =   2  '드롭다운 목록
               TabIndex        =   17
               Top             =   1455
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.TextBox txtTenterSettingClss 
               Height          =   315
               Left            =   5730
               TabIndex        =   16
               Top             =   1980
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.TextBox txtTenterDensity 
               Height          =   315
               Left            =   5025
               TabIndex        =   13
               Top             =   375
               Width           =   1485
            End
            Begin VB.TextBox txtTenterOverFeed 
               Height          =   315
               Left            =   5025
               TabIndex        =   12
               Top             =   15
               Width           =   1485
            End
            Begin VB.TextBox txtTenterVelocity 
               Height          =   315
               Left            =   1260
               TabIndex        =   11
               Top             =   390
               Width           =   1485
            End
            Begin VB.TextBox txtTenterTemper 
               Height          =   315
               Left            =   1260
               TabIndex        =   10
               Top             =   30
               Width           =   1485
            End
            Begin VB.ComboBox cboTenterWorkClss 
               Height          =   300
               Left            =   1260
               Style           =   2  '드롭다운 목록
               TabIndex        =   14
               Top             =   765
               Width           =   2775
            End
            Begin VB.ComboBox cboTenterDryID 
               Height          =   300
               Left            =   1290
               Style           =   2  '드롭다운 목록
               TabIndex        =   18
               Top             =   2760
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.TextBox txtTenterPersonID 
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   1770
               Width           =   1305
            End
            Begin VB.TextBox txtTenterRemark 
               Height          =   615
               Left            =   1230
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   15
               Top             =   1110
               Width           =   2805
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   51
               Left            =   30
               TabIndex        =   85
               Top             =   390
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "속도(㎜)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   52
               Left            =   30
               TabIndex        =   86
               Top             =   765
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
               Caption         =   "작업구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   53
               Left            =   30
               TabIndex        =   87
               Top             =   1770
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
               Caption         =   "작성자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   54
               Left            =   30
               TabIndex        =   88
               Top             =   1110
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
               Caption         =   "비고 사항"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   55
               Left            =   30
               TabIndex        =   89
               Top             =   30
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "온도(℃)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   56
               Left            =   4500
               TabIndex        =   90
               Top             =   1980
               Visible         =   0   'False
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "Setting구분"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   4
               Left            =   3795
               TabIndex        =   94
               Top             =   15
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "Over Feed"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   5
               Left            =   3795
               TabIndex        =   95
               Top             =   375
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "위사밀도"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   6
               Left            =   4380
               TabIndex        =   96
               Top             =   1455
               Visible         =   0   'False
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
               Caption         =   "불량코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   7
               Left            =   60
               TabIndex        =   97
               Top             =   2760
               Visible         =   0   'False
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
               Caption         =   "건조정도"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel pnlPeach 
            Height          =   3015
            Left            =   -71430
            TabIndex        =   98
            Top             =   900
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   5318
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txtPeachPePaBon4 
               Height          =   300
               Left            =   1170
               TabIndex        =   24
               Top             =   1350
               Width           =   1185
            End
            Begin VB.TextBox txtPeachVelocity 
               Height          =   300
               Left            =   1170
               TabIndex        =   20
               Top             =   30
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPePaBon1 
               Height          =   300
               Left            =   1170
               TabIndex        =   21
               Top             =   360
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPePaBon2 
               Height          =   300
               Left            =   1170
               TabIndex        =   22
               Top             =   690
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPePaBon3 
               Height          =   300
               Left            =   1170
               TabIndex        =   23
               Top             =   1020
               Width           =   1185
            End
            Begin VB.TextBox txtPeachDensity 
               Height          =   300
               Left            =   4035
               TabIndex        =   25
               Top             =   30
               Width           =   1185
            End
            Begin VB.TextBox txtPeachTension 
               Height          =   300
               Left            =   4035
               TabIndex        =   26
               Top             =   360
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPressure1 
               Height          =   300
               Left            =   4035
               TabIndex        =   27
               Top             =   690
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPressure2 
               Height          =   300
               Left            =   4035
               TabIndex        =   28
               Top             =   1020
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPressure3 
               Height          =   300
               Left            =   4035
               TabIndex        =   29
               Top             =   1350
               Width           =   1185
            End
            Begin VB.TextBox txtPeachPersonID 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1170
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   2430
               Width           =   1155
            End
            Begin VB.TextBox txtPeachRemark 
               Height          =   660
               Left            =   1170
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   30
               Top             =   1665
               Width           =   4050
            End
            Begin Threed.SSPanel pnlName 
               Height          =   285
               Index           =   10
               Left            =   30
               TabIndex        =   99
               Top             =   2430
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
               Caption         =   "작성자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   11
               Left            =   30
               TabIndex        =   100
               Top             =   1695
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "비고 사항"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   8
               Left            =   2895
               TabIndex        =   114
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "장력"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   9
               Left            =   2895
               TabIndex        =   115
               Top             =   30
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "밀도"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   12
               Left            =   30
               TabIndex        =   116
               Top             =   30
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "속도(㎜)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   13
               Left            =   30
               TabIndex        =   117
               Top             =   1020
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "페파본3"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   15
               Left            =   30
               TabIndex        =   118
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "페파본1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   16
               Left            =   30
               TabIndex        =   119
               Top             =   690
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "페파본2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   19
               Left            =   2895
               TabIndex        =   120
               Top             =   1350
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "압력3"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   20
               Left            =   2895
               TabIndex        =   121
               Top             =   690
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "압력1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   21
               Left            =   2895
               TabIndex        =   122
               Top             =   1020
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "압력2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   18
               Left            =   30
               TabIndex        =   138
               Top             =   1350
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "페파본4"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel pnlDry 
            Height          =   2805
            Left            =   -70350
            TabIndex        =   101
            Top             =   1020
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4948
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboDryCodeID 
               Height          =   300
               Left            =   1230
               Style           =   2  '드롭다운 목록
               TabIndex        =   36
               Top             =   2400
               Visible         =   0   'False
               Width           =   2745
            End
            Begin VB.TextBox cboDryPersonID 
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   1950
               Width           =   945
            End
            Begin VB.TextBox cboDryOverFeed 
               Height          =   315
               Left            =   1290
               TabIndex        =   34
               Top             =   780
               Width           =   1335
            End
            Begin VB.TextBox cboDryVelocity 
               Height          =   315
               Left            =   1290
               TabIndex        =   33
               Top             =   420
               Width           =   1335
            End
            Begin VB.TextBox cboDryTemper 
               Height          =   315
               Left            =   1290
               TabIndex        =   32
               Top             =   60
               Width           =   1335
            End
            Begin VB.TextBox cboDryRemark 
               Height          =   795
               Left            =   1290
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   35
               Top             =   1125
               Width           =   2865
            End
            Begin VB.ComboBox cboDryDryID 
               Height          =   300
               Left            =   2640
               Style           =   2  '드롭다운 목록
               TabIndex        =   38
               Top             =   2730
               Visible         =   0   'False
               Width           =   945
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   17
               Left            =   60
               TabIndex        =   102
               Top             =   420
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "속도(㎜)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   22
               Left            =   60
               TabIndex        =   103
               Top             =   1950
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "작성자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   23
               Left            =   60
               TabIndex        =   104
               Top             =   1140
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
               Caption         =   "비고 사항"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   24
               Left            =   60
               TabIndex        =   105
               Top             =   60
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "온도(℃)"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   315
               Index           =   28
               Left            =   60
               TabIndex        =   106
               Top             =   780
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
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
               Caption         =   "Over Feed"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   30
               Left            =   0
               TabIndex        =   107
               Top             =   2400
               Visible         =   0   'False
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
               Caption         =   "불량코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   31
               Left            =   1320
               TabIndex        =   108
               Top             =   2700
               Visible         =   0   'False
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
               Caption         =   "건조정도"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Index           =   1
            Left            =   -72450
            TabIndex        =   127
            Top             =   810
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   661
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboTenterProc 
               Height          =   300
               Left            =   1260
               Style           =   2  '드롭다운 목록
               TabIndex        =   132
               Top             =   60
               Width           =   2775
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   3
               Left            =   30
               TabIndex        =   133
               Top             =   60
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
               Caption         =   "공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Index           =   2
            Left            =   -71430
            TabIndex        =   128
            Top             =   510
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   661
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboPeachProc 
               Height          =   300
               Left            =   1170
               Style           =   2  '드롭다운 목록
               TabIndex        =   134
               Top             =   30
               Width           =   2295
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   14
               Left            =   30
               TabIndex        =   135
               Top             =   30
               Width           =   1095
               _ExtentX        =   1931
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
               Caption         =   "공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   405
            Index           =   3
            Left            =   -70350
            TabIndex        =   129
            Top             =   600
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   714
            _Version        =   196609
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboDryProc 
               Height          =   300
               Left            =   1290
               Style           =   2  '드롭다운 목록
               TabIndex        =   136
               Top             =   60
               Width           =   2775
            End
            Begin Threed.SSPanel pnlName 
               Height          =   300
               Index           =   26
               Left            =   60
               TabIndex        =   137
               Top             =   60
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
               Caption         =   "공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   60
         TabIndex        =   91
         Top             =   60
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1349
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton opCustomItem 
            Caption         =   "동일업체 + 동일품명"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   420
            Width           =   2175
         End
         Begin VB.OptionButton opOrderID 
            Caption         =   "관리번호"
            Height          =   180
            Left            =   120
            TabIndex        =   92
            Top             =   150
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Label Label1 
         Caption         =   "건 가져오기"
         Height          =   165
         Left            =   5010
         TabIndex        =   68
         Top             =   540
         Width           =   1005
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   795
      Left            =   45
      TabIndex        =   60
      Top             =   8415
      Width           =   1500
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   62
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
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   210
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmInstCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------'
'Private Const REPORTFILE = "\Report\Order.rpt"

'----------------------------------------------------------------'
Private m_nBaseX As Long
Private m_nBaseY As Long
Private m_nBaseBlank As Long

'----------------------------------------------------------------'
Private m_iFlag    As Integer   ' 현재 상태 (추가/수정/삭제/검색)
Private m_bloading As Boolean
Private mOrderID As String      ' OrderID
Private mIndiDate As String, mIndiTime As String   '신규이면 현재시스템 일자, 수정이면 원래 데이터

Private Type mDelType
        xpProName      As String       '프로시저 명
        OrderID        As String       '관리번호
        Process        As String       '[] 공정명
        IndiDate       As String       '지시일자
        IndiTime       As String       '지시시간
End Type
Private m_SelOrderID As String, mProcID As String

'-- 텐터기 레코드 신규/삭제 function
Private Function AddNewTenter_bol() As Boolean
    Dim TTenter As PlusLib2.TTenter
    Dim oInstCondition As PlusLib2.CInstCondition
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHandler

    Set oInstCondition = New PlusLib2.CInstCondition
    oInstCondition.Connection = g_adoCon

    With TTenter
        If m_iFlag = ID_ADDNEW Then
            .JobFlag = "I"
             Call GetNowDate(mIndiDate, mIndiTime)
            .PersonID = g_sUserName
        Else
            .JobFlag = "U"
            .PersonID = txtTenterPersonID.Tag
        End If
        
        .IndiDate = mIndiDate       '[2] 계획일자
        .IndiTime = mIndiTime
        .OrderID = mOrderID
        .Process = Trim(cboTenterProc)
        
        .Temper = GetNumeric(txtTenterTemper)
        .Velocity = GetNumeric(txtTenterVelocity)
        .OverFeed = GetNumeric(txtTenterOverFeed)
        .Density = GetNumeric(txtTenterDensity)
        .SettingClss = txtTenterSettingClss
        
        .WorkCond = cboTenterWorkClss
        .CodeID = cboTenterCodeID.Tag
        .DryID = cboTenterDryID
        .Remark = txtTenterRemark
    End With
    '-----------------------------------------------------------------------------------------
    
    AddNewTenter_bol = oInstCondition.AddNewTenter(TTenter)
    
    Set oInstCondition = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrHandler:
End Function

Private Function AddNewDry_bol() As Boolean
    Dim TDry As PlusLib2.TDry
    Dim oInstCondition As PlusLib2.CInstCondition
    
    Screen.MousePointer = vbHourglass
    
'    On Error GoTo ErrHandler

    Set oInstCondition = New PlusLib2.CInstCondition
    oInstCondition.Connection = g_adoCon

    With TDry
        If m_iFlag = ID_ADDNEW Then
            .JobFlag = "I"
             Call GetNowDate(mIndiDate, mIndiTime)
            .PersonID = g_sUserName
        Else
            .JobFlag = "U"
            .PersonID = cboDryPersonID.Tag
        End If
        
        .IndiDate = mIndiDate       '[2] 계획일자
        .IndiTime = mIndiTime
        .OrderID = mOrderID
        .Process = cboDryProc
        
        .Temper = GetNumeric(cboDryTemper)
        
        .Velocity = GetNumeric(cboDryVelocity)
        
        .OverFeed = GetNumeric(cboDryOverFeed)
        .CodeID = cboDryCodeID
        .Remark = cboDryRemark
    End With
    '-----------------------------------------------------------------------------------------
    
    AddNewDry_bol = oInstCondition.AddNewDry(TDry)
    
    Set oInstCondition = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function

End Function
'-- Refine(수세) 레코드 신규/삭제 function
Private Function AddNewRefine_bol() As Boolean
    Dim TRefine As PlusLib2.TRefine
    Dim oInstCondition As PlusLib2.CInstCondition
    
    Screen.MousePointer = vbHourglass
    
'    On Error GoTo ErrHandler

    Set oInstCondition = New PlusLib2.CInstCondition
    oInstCondition.Connection = g_adoCon

    With TRefine
        If m_iFlag = ID_ADDNEW Then
            .JobFlag = "I"
             Call GetNowDate(mIndiDate, mIndiTime)
            .PersonID = g_sUserName
        Else
            .JobFlag = "U"
            .PersonID = txtRefinePersonID.Tag
        End If
        
        .IndiDate = mIndiDate       '[2] 계획일자
        .IndiTime = mIndiTime
        .OrderID = mOrderID
        .Process = Trim(cboRefineProc)
        .Temper = GetNumeric(txtRefineTemper)
        .Velocity = GetNumeric(txtRefineVelocity)
        .RefineClss = cboRefineClss
        .SettingClss = txtRefineSettingClss
        .Remark = txtRefineRemark
    
    End With
    '-----------------------------------------------------------------------------------------
    
    AddNewRefine_bol = oInstCondition.AddNewRefine(TRefine)
    
    Set oInstCondition = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
End Function


Private Function AddNewCPBPre_bol() As Boolean
    
    Dim TCPBPre As PlusLib2.TCPBPre
    Dim oTCPBPre As PlusLib2.CInstCondition
    
    Screen.MousePointer = vbHourglass
    
'    On Error GoTo ErrHandler

    Set oTCPBPre = New PlusLib2.CInstCondition
    oTCPBPre.Connection = g_adoCon
    
    With TCPBPre
        If m_iFlag = ID_ADDNEW Then
            .JobFlag = "I"
             Call GetNowDate(mIndiDate, mIndiTime)
            .PersonID = g_sUserName
        Else
            .JobFlag = "U"
            .PersonID = Trim(txtCPBPersonID.Tag)
        End If
        
        .IndiDate = mIndiDate       '[2] 계획일자
        .IndiTime = mIndiTime
        .OrderID = mOrderID
        .Process = Trim(cboCPBProc)
        .RefineClss = Trim(cboCPBRefineClss)
        .Remark = Trim(txtCPBRemark.Text)
        .Velocity = GetNumeric(txtCPBVelocity)
    
    End With
    
    AddNewCPBPre_bol = oTCPBPre.AddNewCPBPre(TCPBPre)
    
    Set oTCPBPre = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
End Function

Private Sub cboCPBProc_Click()
    If m_iFlag = -1 And opProcess.Value = True Then
        Call ShowData
    End If
'GetInstDefectList
End Sub

Private Sub cboCPBProc_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboCPBRefineClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub


Private Sub cboDryCodeID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryDryID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryOverFeed_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryPersonID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryProc_Click()
    Select Case m_iFlag
        Case -1
            If opProcess.Value = True Then
                Call ShowData
            End If
        Case Else
            Call SetInstDefct(cboDryCodeID, cboDryProc.Text)
            If cboDryCodeID.ListCount > 0 Then
                cboDryCodeID.ListIndex = 0
            Else
                cboDryCodeID.ListIndex = -1
            End If
    End Select
End Sub

Private Sub cboDryProc_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryTemper_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboDryVelocity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboPeachProc_Click()
    If m_iFlag = -1 And opProcess.Value = True Then
        Call ShowData
    End If

End Sub

Private Sub cboPeachProc_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboRefineClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub



Private Sub cboRefineProc_Click()
    If m_iFlag = -1 And opProcess.Value = True Then
        Call ShowData
    End If

End Sub

Private Sub cboRefineProc_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboTenterCodeID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboTenterDryID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboTenterProc_Click()

    
    Select Case m_iFlag
        Case -1
            If opProcess.Value = True Then
                Call ShowData
            End If
        Case Else
            Call SetInstDefct(cboTenterCodeID, cboTenterProc.Text)
            If cboTenterCodeID.ListCount > 0 Then
                cboTenterCodeID.ListIndex = 0
            Else
                cboTenterCodeID.ListIndex = -1
            End If
    End Select
End Sub

Private Sub cboTenterProc_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboTenterWorkClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub



Private Sub cmdCopy_Click()
    m_iFlag = ID_ADDNEW
    Call cmdOperate_Click(3)   '저장을 선택 한 것 과 같은 프로시저를 적용한다.
End Sub

Private Sub Form_Load()
    Dim i%
    
'    Me.Move 0, 0, 11975, 9660
    PlusMDI.pnlMenu.Visible = False

    Me.Move 0, 0, 15360, 9840
    
    m_iFlag = -1
    
    Call InitGrid
    Call SetOperate(Me)
    Call SetModeForEditor
    
    Show
'
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
    
    dtpDate(0) = Date
    dtpDate(1) = Date
    
    '--- 정련CPB기 정련구분
    With cboCPBRefineClss
        .Clear
        .AddItem "1차완"
        .AddItem "2차완"
        .AddItem "수정"
        .AddItem "효소발"
    End With
    
    '---- 수세기 정련구분
    With cboRefineClss
        .Clear
        .AddItem "1차완"
        .AddItem "2차완"
        .AddItem "수정"
        .AddItem "효소발"
    End With
    
    '--- 텐터 작업조건
    With cboTenterWorkClss
        .Clear
        .AddItem "WR"
        .AddItem "PD"
        .AddItem "ST"
        .AddItem "TF"
        .AddItem "TP"
    End With
    
    Call FillComboBox(cboCPBProc, GetMachineProcID("정련 C.P.B"))       '정련 C.P.B
    Call FillComboBox(cboRefineProc, GetMachineProcID("수세기"))        '수세
    Call FillComboBox(cboTenterProc, GetMachineProcID("텐터기"))        '텐터
    Call FillComboBox(cboPeachProc, GetMachineProcID("Peach기"))        'Peach
    Call FillComboBox(cboDryProc, GetMachineProcID("건조기"))           'Dry
    
    '최근레코드 10건 가져오기
    txtSelRecs.Text = "10"
    
    Call ClearData
    
    
End Sub


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index
    Case 0  '수주일자 선택
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Case 1  '거래처
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(1).Enabled = True
            txtSearch(1).SetFocus
            cmdFind(0).Visible = True
        Else
            txtSearch(1).Enabled = False
            cmdFind(0).Visible = False
        End If
    Case 2  '관리번호 & Order NO
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(2).Enabled = True
            txtSearch(2).SetFocus
        Else
            txtSearch(2).Enabled = False
        End If
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub


Private Sub NonEditMode(NewValue As Boolean, Optional PreTab_int As Integer = -1)
    Dim TabIndex As Integer
    
    If PreTab_int = -1 Then
        TabIndex = tabProc.Tab    '현재의 Tab Editmode설정
    Else
        TabIndex = PreTab_int          '이전 Tab Editmode 설정
    End If
    
    Select Case TabIndex
        Case 0                     '정련CPB기
            cboCPBProc.Locked = NewValue
            txtCPBVelocity.Locked = NewValue
            cboCPBRefineClss.Locked = NewValue
            txtCPBPersonID.Locked = NewValue
            txtCPBRemark.Locked = NewValue
            
        Case 1                      '수세기
            cboRefineProc.Locked = NewValue

            txtRefineTemper.Locked = NewValue
            txtRefineVelocity.Locked = NewValue
            cboRefineClss.Locked = NewValue
            txtRefineSettingClss.Locked = NewValue
            txtRefinePersonID.Locked = NewValue
            txtRefinePersonID.Locked = NewValue
            txtRefineRemark.Locked = NewValue
            
        Case 2
            cboTenterProc.Locked = NewValue
            txtTenterTemper.Locked = NewValue
            txtTenterVelocity.Locked = NewValue
            txtTenterOverFeed.Locked = NewValue
            txtTenterDensity.Locked = NewValue
            txtTenterSettingClss.Locked = NewValue
            cboTenterWorkClss.Locked = NewValue
            cboTenterCodeID.Locked = NewValue
            cboTenterDryID.Locked = NewValue
            txtTenterPersonID.Locked = NewValue
            txtTenterRemark.Locked = NewValue
            
        Case 3
            cboPeachProc.Locked = NewValue
            txtPeachVelocity.Locked = NewValue
            txtPeachPePaBon1.Locked = NewValue
            txtPeachPePaBon2.Locked = NewValue
            txtPeachPePaBon3.Locked = NewValue
            txtPeachPePaBon4.Locked = NewValue
            txtPeachDensity.Locked = NewValue
            txtPeachTension.Locked = NewValue
            txtPeachPressure1.Locked = NewValue
            txtPeachPressure2.Locked = NewValue
            txtPeachPressure3.Locked = NewValue
            txtPeachPersonID.Locked = NewValue
            txtPeachRemark.Locked = NewValue
        
        Case 4
            cboDryProc.Locked = NewValue
            cboDryTemper.Locked = NewValue
            cboDryVelocity.Locked = NewValue
            cboDryOverFeed.Locked = NewValue
            cboDryCodeID.Locked = NewValue
            cboDryDryID.Locked = NewValue
            cboDryPersonID.Locked = NewValue
            cboDryRemark.Locked = NewValue
    End Select
    
End Sub


Private Sub cmdFind_Click(Index As Integer)

    Select Case Index
        Case 0   '거래처 코드
            Call ReturnCode(LG_CUSTOM, 0, True, txtSearch(1))
    End Select
End Sub

'********************************************************
'* Date : 2001-06-21 (THU)
'*
'* Description: Operate Button의 Index 상수
'*
'********************************************************
Private Sub cmdOperate_Click(Index As Integer)

    On Error GoTo ErrHandler
    
    If (Index <> ID_CANCEL) And (grdOrder.FixedRows = grdOrder.Rows) Then
        MsgBox "수주건을 선택하지 않았습니다", vbInformation, "수주건 선택"
        Exit Sub
    End If
    
    Select Case Index
        '-------------------------------------------------------------------------------------'
        Case ID_ADDNEW
            m_iFlag = ID_ADDNEW

            Call ClearData
            Call ChangeMode(Me, False)
            Call SetModeForEditor
        '-------------------------------------------------------------------------------------'
        Case ID_UPDATE
            If grdOrder.Rows = grdOrder.FixedRows Then
                MsgBox LoadResString(111), vbInformation
                cmdSearch.SetFocus
                Exit Sub
            End If
            m_iFlag = ID_UPDATE
            
            Call ChangeMode(Me, False)
            Call SetModeForEditor
            
        '-------------------------------------------------------------------------------------'
        Case ID_DELETE
        '    If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo) = vbYes Then
                If DeleteData() Then
                    Call ShowData
                    Call ClearData
                End If
            End If
            Call SetModeForEditor
        '-------------------------------------------------------------------------------------'
        Case ID_SAVE
            If SaveData() Then
                Call ChangeMode(Me, True)
                Call ShowData
                m_iFlag = -1
            End If
            Call SetModeForEditor
        
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL
            m_iFlag = -1
            Call ChangeMode(Me, True)
      '      Call SetModeForEditor
            Call ShowData
    End Select

    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmInstCondition.cmdOperate_Click", Err.Description)

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

'------------------------------------------------------------------------------
' 프로그램명: InitGrid
'-----------------------------------------------------------------------------
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
        .ScrollBars = flexScrollBarBoth
            
        .TextArray(0) = " "
        .TextArray(1) = "완료":         .ColWidth(1) = 300
        .TextArray(2) = "관리번호":     .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(3) = "Order No.":    .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignLeftCenter    '[2] 오더 번호(15)
        .TextArray(4) = "거래처명":     .ColWidth(4) = 1920:    .ColAlignment(4) = flexAlignLeftCenter   '[4] 거래처명
        .TextArray(5) = "수주량":       .ColWidth(5) = 1000:    .ColAlignment(5) = flexAlignLeftCenter   '[4] 거래처명
        
        .ColHidden(3) = True
        .ColAlignment(1) = flexAlignCenterCenter
        
        .Redraw = True
    End With
    
    Select Case tabProc.Tab
        Case 0:     Call TitGrdCPBPre             '---- 정련CPB기 데이터 나타내기
        Case 1:     Call TitGrdReFine             '---- 수세기 데이터 나타내기
        Case 2:     Call TitGrdTenter             '---- 텐터기 데이터 나타내기
        Case 3:     Call TitGrdPeach              '---- Peach기 데이터 나타내기
        Case 4:     Call TitGrdDry                '---- Dry기 데이터 나타내기
    End Select
End Sub

Private Sub ClearData()
    Select Case tabProc.Tab
        Case 0                     '정련CPB기
            Call ClearScreen(Me, "pnlCPB")           '---CPB정련기
            txtCPBPersonID.Text = g_sPersonName
            txtCPBPersonID.Tag = g_sUserName
            cboCPBProc.SetFocus
            
        Case 1
            Call ClearScreen(Me, "pnlRefine")        '---수세기
            txtRefinePersonID.Text = g_sPersonName
            txtRefinePersonID.Tag = g_sUserName
            cboRefineProc.SetFocus
            
        Case 2
            Call ClearScreen(Me, "pnlTenter")        '---Tenter
            txtTenterPersonID.Text = g_sPersonName
            txtTenterPersonID.Tag = g_sUserName
            cboTenterProc.SetFocus
            
        Case 3
            Call ClearScreen(Me, "pnlPeach")         '---Peach
            txtPeachPersonID.Text = g_sPersonName
            txtPeachPersonID.Tag = g_sUserName
            cboPeachProc.SetFocus
        
        Case 4
            Call ClearScreen(Me, "pnlDry")           '---건조기
            cboDryPersonID.Text = g_sPersonName
            cboDryPersonID.Tag = g_sUserName
            cboDryProc.SetFocus
    End Select
    
End Sub
Private Sub SetModeForEditor()
'    Dim TabIndex As Integer
    Dim Proc_bol As Boolean, Edit_bol As Boolean
    
'''    If PreTab_int = -1 Then
'''        TabIndex = tabProc.Tab         '현재의 Tab Editmode설정
'''    Else
'''        TabIndex = PreTab_int          '이전 Tab Editmode 설정
'''    End If
    
    Select Case m_iFlag
        Case ID_ADDNEW
            Edit_bol = True
            Proc_bol = True
        '-------------------------------------------------------------------------------------'
        Case ID_UPDATE
            Edit_bol = True
            Proc_bol = False
        '-------------------------------------------------------------------------------------'
        Case ID_DELETE
            Edit_bol = False
            Proc_bol = False
        '-------------------------------------------------------------------------------------'
''        Case ID_SAVE
''            Edit_bol = False
''            Proc_bol = True
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL, -1, ID_SAVE
            Edit_bol = False
            Proc_bol = True
    End Select
    
    Select Case tabProc.Tab
        Case 0
            pnlCPB.Enabled = Edit_bol
            cboCPBProc.Enabled = Proc_bol
        Case 1
            pnlRefine.Enabled = Edit_bol
            cboRefineProc.Enabled = Proc_bol
        Case 2
            pnlTenter.Enabled = Edit_bol
            cboTenterProc.Enabled = Proc_bol
        Case 3
            pnlPeach.Enabled = Edit_bol
            cboPeachProc.Enabled = Proc_bol
        Case 4
            pnlDry.Enabled = Edit_bol
            cboDryProc.Enabled = Proc_bol
    End Select
    
End Sub

Sub FillGrdCondition()

    With grdOrder
'        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
        If .Row < .FixedRows Or .Row >= .Rows Then Exit Sub
        
        mOrderID = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)

    End With
    Call ShowData
End Sub

Private Sub ShowData()
    GrdCondition.Rows = GrdCondition.FixedRows
    
    Select Case tabProc.Tab
        Case 0:            Call FillGrdCPBPre             '---- 정련CPB기 데이터 나타내기
        Case 1:            Call FillGrdRefine             '---- 수세기 데이터 나타내기
        Case 2:            Call FillGrdTenter             '---- 텐터기 데이터 나타내기
        Case 3:            Call FillGrdPeach              '---- Peach기 데이터 나타내기
        Case 4:            Call FillGrdDry                '---- Dry기 데이터 나타내기
    End Select
    
    If GrdCondition.Rows > GrdCondition.FixedRows Then
        GrdCondition.Row = GrdCondition.FixedRows
    End If
    Call ShowEditData
End Sub

Private Sub FillGridOrder()
    Dim rs As ADODB.Recordset
    Dim lNowRow&, lNowSum&, i%
    Dim oOrder As PlusLib2.CInstCondition
    
    On Error GoTo ErrHandler
    
    m_bloading = True
    
    Set oOrder = New PlusLib2.CInstCondition
    oOrder.Connection = g_adoCon
    
    Set rs = oOrder.GetDraftOrder(IIf(chkSearch(0).Value = vbChecked, False, True), _
                MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                IIf(chkSearch(1).Value = vbChecked, txtSearch(1).Tag, ""), _
                IIf(chkSearch(2).Value = vbChecked, txtSearch(2).Text, ""))

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
                lNowSum = lNowSum + (rs!OrderQty * 1.0936)
            End If
            
            .AddItem CStr(.Rows) & vbTab & IIf(Trim(rs!CloseClss) = "", "", "■") & vbTab & _
                    MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & rs!kCustom & vbTab & rs!OrderQty
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
            
            Call ShowData
        Else
            .HighLight = flexHighlightNever
                    
     '       Call ClearData(0)
        End If
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            Call FillGrdCondition
        End If
    End With
    
    m_bloading = False
    
    grdTotal.TextArray(1) = Format(grdOrder.Rows - 1, "#,##0 건")
    grdTotal.TextArray(2) = Format(lNowSum, "#,##0 YDS")
    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing
    
    grdOrder.Redraw = flexRDDirect
    
    Call ErrorBox(Err.Number, "frmInstCondition.FillGridOrder", Err.Description)
End Sub

Private Function SaveData() As Boolean
    Dim Error_bol As Boolean

    On Error GoTo ErrHandler

    Select Case tabProc.Tab
    
        Case 0   '--- 정련CPB기
            Error_bol = AddNewCPBPre_bol
            
        Case 1   '--- 수세기( Refine )
            Error_bol = AddNewRefine_bol
        
        Case 2   '--- 텐터기
            Error_bol = AddNewTenter_bol
        
        Case 3   '--- Peach기
            Error_bol = AddNewPeach_bol
        
        Case 4   '--- Dry기
            Error_bol = AddNewDry_bol
    
    End Select
    SaveData = Error_bol
    Exit Function
ErrHandler:
    SaveData = Error_bol

    Call ErrorBox(Err.Number, "frmInstCondition.SaveData", Err.Description)
    
    Resume Next

End Function

'---- Peach 데이터 등록
Private Function AddNewPeach_bol() As Boolean
    Dim TPeach As PlusLib2.TPeach
    
    Dim oInstCondition As PlusLib2.CInstCondition
    
    Screen.MousePointer = vbHourglass
    
'    On Error GoTo ErrHandler

    Set oInstCondition = New PlusLib2.CInstCondition
    oInstCondition.Connection = g_adoCon

    With TPeach
        If m_iFlag = ID_ADDNEW Then
            .JobFlag = "I"
             Call GetNowDate(mIndiDate, mIndiTime)
            .PersonID = g_sUserName
             
        Else
            .JobFlag = "U"
            .PersonID = txtPeachPersonID.Tag
        End If
        
        .PePaBon1 = GetNumeric(Trim(txtPeachPePaBon1.Text))
        .PePaBon2 = GetNumeric(Trim(txtPeachPePaBon2.Text))
        .PePaBon3 = GetNumeric(Trim(txtPeachPePaBon3.Text))
        .PePaBon4 = GetNumeric(Trim(txtPeachPePaBon4.Text))
        .Pressure1 = GetNumeric(Trim(txtPeachPressure1.Text))
        .Pressure2 = GetNumeric(Trim(txtPeachPressure2.Text))
        .Pressure3 = GetNumeric(Trim(txtPeachPressure3.Text))
        .Density = GetNumeric(txtPeachDensity)
        .Tention = GetNumeric(txtPeachTension)
        .Velocity = GetNumeric(txtPeachVelocity)
        
        .Remark = txtPeachRemark
        .IndiDate = mIndiDate       '[2] 계획일자
        .IndiTime = mIndiTime
        .OrderID = mOrderID
        .Process = cboPeachProc
    
    End With
    '-----------------------------------------------------------------------------------------
    
    AddNewPeach_bol = oInstCondition.AddNewPeach(TPeach)
    
    Set oInstCondition = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function

End Function

' 작업조건 지시 삭제 공용 사용
Private Function DeleteData() As Boolean
    Dim TDelType As PlusLib2.TDelType
    Dim oInstCondition As PlusLib2.CInstCondition
    Dim mDelType  As mDelType
    
    Set oInstCondition = New PlusLib2.CInstCondition
    oInstCondition.Connection = g_adoCon
    
    Select Case tabProc.Tab
        Case 0   '--- 정련CPB기
            With mDelType
                .IndiDate = mIndiDate
                .IndiTime = mIndiTime
                .OrderID = m_SelOrderID
                .Process = mProcID
                .xpProName = "wi_CPBPre"
            End With
            
        Case 1   '--- 수세기( Refine )
            With mDelType
                .IndiDate = mIndiDate
                .IndiTime = mIndiTime
                .OrderID = m_SelOrderID
                .Process = mProcID
                .xpProName = "wi_Refine"
            End With
        
        Case 2   '--- 텐터기
            With mDelType
                .IndiDate = mIndiDate
                .IndiTime = mIndiTime
                .OrderID = m_SelOrderID
                .Process = mProcID
                .xpProName = "wi_Tenter"
            End With
            
        Case 3   '--- Peach기
            With mDelType
                .IndiDate = mIndiDate
                .IndiTime = mIndiTime
                .OrderID = m_SelOrderID
                .Process = mProcID
                .xpProName = "wi_Peach"
            End With
            
        Case 4   '--- Dry기
             With mDelType
                .IndiDate = mIndiDate
                .IndiTime = mIndiTime
                .OrderID = m_SelOrderID
                .Process = mProcID
                .xpProName = "wi_Dry"
            End With
    End Select
    
    With TDelType
        .IndiDate = mDelType.IndiDate
        .IndiTime = mDelType.IndiTime
        .OrderID = mDelType.OrderID
        .Process = mDelType.Process
        .xpProName = mDelType.xpProName
    End With
    
    DeleteData = oInstCondition.DelInstCondition(TDelType)
    
    Set oInstCondition = Nothing
    
    Exit Function
ErrHandler:
    Set oInstCondition = Nothing
    Call ErrorBox(Err.Number, "frmInstCondition.DeleteData", Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdCondition_Click()
    Call ShowEditData
End Sub

Sub ShowEditData()
    Dim dIndiDateTime As String
    Dim dViewName As String
    Dim oInstCondition As PlusLib2.CInstCondition
    
    On Error GoTo ErrHandler
    
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset
    
    Call ClearData
    
    Screen.MousePointer = vbHourglass

    Set adoCmd = New ADODB.Command
    
    With GrdCondition
        If .Row < .FixedRows Or .Row >= .Rows Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        m_SelOrderID = MakeOrderID(.TextMatrix(.Row, 1), OM_REDUCE)    '현재선택한 OrderID
        mProcID = GetProcessID(.TextMatrix(.Row, 5))
        dIndiDateTime = .TextMatrix(.Row, 6)
        mIndiDate = Replace(Left(Trim(dIndiDateTime), 10), "-", "")
        mIndiTime = Replace(Right(Trim(dIndiDateTime), 5), ":", "")
        
        Select Case tabProc.Tab
            Case 0:    dViewName = "vw_wiCPBPre" '----정련 CPB기
            Case 1:    dViewName = "vw_wiRefine" '----수세기
            Case 2:    dViewName = "vw_wiTenter" '----텐터기
            Case 3:    dViewName = "vw_wiPeach"  '----Peach기
            Case 4:    dViewName = "vw_wiDry"    '----Dry기
        End Select

        Set oInstCondition = New PlusLib2.CInstCondition
        oInstCondition.Connection = g_adoCon

        Set rs = oInstCondition.GetInstOneRec(dViewName, m_SelOrderID, mProcID, mIndiDate, mIndiTime)
        
        Set oInstCondition = Nothing
        
        Select Case tabProc.Tab
            Case 0:    Call ShowCPBPre(rs)  '----정련 CPB기
            Case 1:    Call ShowRefine(rs)  '----수세기
            Case 2:    Call ShowTenter(rs)  '----텐터기
            Case 3:    Call ShowPeach(rs)   '----Peach기
            Case 4:    Call ShowDry(rs)     '----Dry기
        End Select
        
        rs.Close
        Set rs = Nothing
        
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:

End Sub

'--- 정련 c.p.b기 에디터에 나타내기
Sub ShowCPBPre(ByVal prs As ADODB.Recordset)
    If prs.RecordCount > 0 Then
        cboCPBProc.ListIndex = FindItem(cboCPBProc, prs!ProcessName)
        txtCPBVelocity = prs!Velocity
        
        cboCPBRefineClss.ListIndex = FindItem(cboCPBRefineClss, prs!RefineClss)
        txtCPBRemark.Text = prs!Remark
        txtCPBPersonID.Text = prs!PersonName
        txtCPBPersonID.Tag = prs!PersonID
        
    End If
End Sub


'--- 수세기 에디터에 나타내기
Sub ShowRefine(ByVal prs As ADODB.Recordset)
    If prs.RecordCount > 0 Then
        cboRefineProc.ListIndex = FindItem(cboRefineProc, prs!ProcessName)
        txtRefineTemper = prs!Temper
        txtRefineVelocity = prs!Velocity
        cboRefineClss.ListIndex = FindItem(cboRefineClss, prs!RefineClss)
        
        txtRefineRemark.Text = prs!Remark
        txtRefinePersonID.Text = prs!PersonName
        txtRefinePersonID.Tag = prs!PersonID
        
    End If
End Sub

'--- 텐터기 에디터에 나타내기
Sub ShowTenter(ByVal prs As ADODB.Recordset)
    If prs.RecordCount > 0 Then
        cboTenterProc.ListIndex = FindItem(cboTenterProc, prs!ProcessName)
        txtTenterTemper = Trim(prs!Temper)
        txtTenterVelocity = Trim(prs!Velocity)
        txtTenterOverFeed = Trim(prs!OverFeed)
        txtTenterDensity = Trim(prs!Density)
        txtTenterSettingClss = Trim(prs!SettingClss)
        cboTenterWorkClss.ListIndex = FindItem(cboTenterWorkClss, prs!WorkCond)
        cboTenterCodeID.ListIndex = FindItem(cboTenterCodeID, prs!CodeIDName)
        txtTenterPersonID.Text = Trim(prs!PersonName)
        txtTenterPersonID.Tag = Trim(prs!PersonID)
        txtTenterRemark.Text = Trim(prs!Remark)
    End If

End Sub

'--- Peach기 에디터에 나타내기
Sub ShowPeach(ByVal prs As ADODB.Recordset)
    If prs.RecordCount > 0 Then
    
        cboPeachProc.ListIndex = FindItem(cboPeachProc, prs!ProcessName)
        txtPeachVelocity = prs!Velocity
        
        txtPeachPePaBon1 = prs!PePaBon1
        txtPeachPePaBon2 = prs!PePaBon2
        txtPeachPePaBon3 = prs!PePaBon3
        txtPeachPePaBon4 = prs!PePaBon4
        
        txtPeachDensity = prs!Density
        txtPeachTension = prs!Tension
        
        txtPeachPressure1 = prs!Pressure1
        txtPeachPressure2 = prs!Pressure2
        txtPeachPressure3 = prs!Pressure3
        
        txtPeachPersonID.Text = prs!PersonName
        txtPeachPersonID.Tag = prs!PersonID
        
        txtPeachRemark = prs!Remark
    End If
End Sub

'--- 건조기 에디터에 나타내기
Sub ShowDry(ByVal prs As ADODB.Recordset)
    If prs.RecordCount > 0 Then
    
        cboDryProc.ListIndex = FindItem(cboDryProc, prs!ProcessName)
        cboDryTemper = prs!Temper
        cboDryVelocity = prs!Velocity
        cboDryOverFeed = prs!OverFeed
        
        cboDryCodeID.ListIndex = FindItem(cboDryCodeID, prs!CodeIDName)
        
        cboDryPersonID.Text = prs!PersonName
        cboDryPersonID.Tag = prs!PersonID
        
        cboDryRemark = prs!Remark
        
    End If
End Sub

Private Sub GrdCondition_RowColChange()
    Call ShowEditData
End Sub

Private Sub grdOrder_AfterSort(ByVal Col As Long, Order As Integer)
    Call FillGrdCondition
End Sub


Private Sub grdOrder_Click()
    Call FillGrdCondition

End Sub

Private Sub grdOrder_RowColChange()
    If m_bloading Then Exit Sub
    
    Call FillGrdCondition
End Sub

'Private Sub mskOrderID_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        Call NextFocus
'    End If
'End Sub

Private Sub opCustomItem_Click()
    Call FillGrdCondition

End Sub

Private Sub opMachine_Click()
    Call FillGrdCondition

End Sub

Private Sub opOrderID_Click()
    Call FillGrdCondition

End Sub

Private Sub opProcess_Click()
    Call FillGrdCondition

End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdOrder
        If optOrder(0).Value Then '[0] 관리번호
            .ColHidden(2) = True
            .ColHidden(3) = False
            chkSearch(2).Caption = "Order No."
        Else '[1] Order No.
            .ColHidden(2) = False
            .ColHidden(3) = True
            chkSearch(2).Caption = "관리번호"
        End If
    End With
End Sub
'
'Private Sub txtBox_LostFocus(Index As Integer)
'    If Not IsNumeric(txtBox(Index)) Then txtBox(Index) = "0"
'
'End Sub

'Private Sub txtCode_Change(Index As Integer)
'    If Index = 1 And m_iFlag >= 0 Then
'        txtName(6) = txtCode(1)         ' 품명 >>>> Tag 품명
'    End If
'End Sub

'Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        If Index = 0 Then               '[1] 거래처 코드
'            Call ReturnCode(LG_CUSTOM, , False, txtCode(0))
'        ElseIf Index = 1 Then           '[2] 품명 코드
'            Call ReturnCode(LG_ARTICLE, , False, txtCode(1))
'        ElseIf Index = 2 Then           '[3] 가공구분 코드
'            Call ReturnCode(LG_WORK, , False, txtCode(2))
'        End If
'    End If
'End Sub
'
'Private Sub txtName_Change(Index As Integer)
'    If Index = 0 And m_iFlag >= 0 Then
'        txtName(7) = txtName(0)    ' Order NO. >>>> Tag 주문번호
'    End If
'End Sub
'
'Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn And Index = 1 Then ' 출고처 선택
'        Call ReturnCode(LG_CUSTOM, , False, txtName(1))
'    End If
'End Sub

Private Sub tabProc_Click(PreviousTab As Integer)
    Call cmdOperate_Click(4)
    m_iFlag = -1
    Call SetModeForEditor
    Call ShowData
End Sub
Sub TitGrdCPBPre()
    '정련 CPB기 GRID에 TITLE 나타내기
    Call SetVSFlexGrid(GrdCondition)
    With GrdCondition
        .Redraw = False
        .Cols = 9
            
        .TextArray(1) = "Order No":    .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(2) = "관리번호":    .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(3) = "거래처":      .ColWidth(3) = 2000:    .ColAlignment(3) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(4) = "품명":        .ColWidth(4) = 2600:    .ColAlignment(4) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(5) = "공정":        .ColWidth(5) = 1400:    .ColAlignment(5) = flexAlignCenterCenter    '[2] 오더 번호(15)
        .TextArray(6) = "지시일시":    .ColWidth(6) = 1600:    .ColAlignment(6) = flexAlignCenterCenter    '[2] 오더 번호(15)
        .TextArray(7) = "속도":        .ColWidth(7) = 1300:    .ColAlignment(7) = flexAlignCenterCenter     '[4] 거래처명
        .TextArray(8) = "정련구분":    .ColWidth(8) = 800:    .ColAlignment(8) = flexAlignCenterCenter    '[4] 거래처명
        
        Call ColHiddenFalse(GrdCondition, .Cols)
        .ColHidden(8) = True
        .Redraw = True
    End With
    GrdCondition.ScrollBars = flexScrollBarBoth
    GrdCondition.ColHidden(2) = True

End Sub
'-- CPB( C0 ) 데이터 Select
Sub FillGrdCPBPre()
    Dim II%
    '-- Grid Title 설정
    Call TitGrdCPBPre
    
    '정련 CPB기 GRID에 TITLE 나타내기
    Dim oInstCPBPre As PlusLib2.CInstCondition
    Dim rs As ADODB.Recordset
    Dim sPattern$, sPatternID$

    Screen.MousePointer = vbHourglass

  '  On Error GoTo ErrHandler

    Set oInstCPBPre = New PlusLib2.CInstCondition
    oInstCPBPre.Connection = g_adoCon
    
    Set rs = oInstCPBPre.GetInstRecord("xp_InstCondi_sDraftCPBPre", mOrderID _
                                    , IIf(opOrderID.Value = True, 1, 2) _
                                    , IIf(opMachine.Value = True, 1, 0), Trim$(cboCPBProc.Text), val(txtSelRecs))
    II = 1
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            GrdCondition.AddItem "" & vbTab & MakeOrderID(rs("OrderID"), OM_EXPAND) & vbTab & rs("OrderNO") & vbTab & rs("KCustom") & vbTab & _
                               rs("Article") & vbTab & rs("ProcessName") & vbTab & MakeDate(DF_LONG, rs("IndiDate")) & " " & Format$(rs("IndiTime"), "0#:##") & vbTab & _
                               rs("Velocity") & vbTab & rs("RefineClss")
        II = II + 1
        If II > val(txtSelRecs) Then
            Exit Do
        End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
    End If
    
    Set oInstCPBPre = Nothing
    
    
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oInstCPBPre = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmInstCondition.FillGrdCPBPre", Err.Description)
End Sub

Sub TitGrdReFine()
    Call SetVSFlexGrid(GrdCondition)
    With GrdCondition
        .Redraw = False
        .Cols = 11
            
        .TextArray(1) = "Order No":      .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(2) = "관리번호":      .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(3) = "거래처":        .ColWidth(3) = 2400:    .ColAlignment(3) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(4) = "품명":          .ColWidth(4) = 2000:    .ColAlignment(4) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(5) = "공정":          .ColWidth(5) = 1400:     .ColAlignment(5) = flexAlignCenterCenter      '[2] 오더 번호(15)
        .TextArray(6) = "지시일시":      .ColWidth(6) = 1600:    .ColAlignment(6) = flexAlignCenterCenter      '[2] 오더 번호(15)
        
        .TextArray(7) = "온도":          .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(8) = "속도":          .ColWidth(8) = 1000:    .ColAlignment(8) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(9) = "정련구분":      .ColWidth(9) = 900:    .ColAlignment(9) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(10) = "Setting구분":  .ColWidth(10) = 900:   .ColAlignment(10) = flexAlignCenterCenter      '[4] 거래처명
        .ScrollBars = flexScrollBarBoth
        
        Call ColHiddenFalse(GrdCondition, .Cols)
        .ColHidden(2) = True
        .ColHidden(9) = True
        .ColHidden(10) = True
        .Redraw = True
    End With
End Sub
'-- 수세( DO / EO  )  데이터 Select
Sub FillGrdRefine()
    Dim II%
    '---
    Call TitGrdReFine
    
    II = 1
    
    '수세 GRID에 TITLE 나타내기
    Dim oInstCPBPre As PlusLib2.CInstCondition
    Dim rs As ADODB.Recordset
    Dim sPattern$, sPatternID$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oInstCPBPre = New PlusLib2.CInstCondition
    oInstCPBPre.Connection = g_adoCon
    
    Set rs = oInstCPBPre.GetInstRecord("xp_InstCondi_sDraftRefine", mOrderID _
                                    , IIf(opOrderID.Value = True, 1, 2) _
                                    , IIf(opMachine.Value = True, 1, 0), Trim$(cboRefineProc.Text), val(txtSelRecs))
                                    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            GrdCondition.AddItem "" & vbTab & MakeOrderID(rs("OrderID"), OM_EXPAND) & vbTab & rs("OrderNO") & vbTab & rs("KCustom") & vbTab & _
                               rs("Article") & vbTab & rs("Process") & vbTab & MakeDate(DF_LONG, rs("IndiDate")) & " " & Format$(rs("IndiTime"), "0#:##") & vbTab & _
                               rs("Temper") & vbTab & rs("Velocity") & vbTab & rs("RefineClss") & vbTab & rs("SettingClss")
                               
            rs.MoveNext
            II = II + 1
            If II > val(txtSelRecs) Then
                Exit Do
            End If
        Loop
        rs.Close
        Set rs = Nothing
        
    End If
    
    Set oInstCPBPre = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:

End Sub
Sub TitGrdTenter()
    Call SetVSFlexGrid(GrdCondition)
    With GrdCondition
        .Redraw = False
        .Cols = 15
            
        .TextArray(1) = "Order No":       .ColWidth(1) = 1300:     .ColAlignment(1) = flexAlignCenterCenter     '[1] 관리 번호(9)
        .TextArray(2) = "관리번호":       .ColWidth(2) = 1300:     .ColAlignment(2) = flexAlignCenterCenter     '[1] 관리 번호(9)
        .TextArray(3) = "거래처":         .ColWidth(3) = 1800:     .ColAlignment(3) = flexAlignCenterCenter     '[1] 관리 번호(9)
        .TextArray(4) = "품명":           .ColWidth(4) = 1700:     .ColAlignment(4) = flexAlignCenterCenter     '[1] 관리 번호(9)
        .TextArray(5) = "공정":           .ColWidth(5) = 1000:     .ColAlignment(5) = flexAlignCenterCenter       '[2] 오더 번호(15)
        .TextArray(6) = "지시일시":       .ColWidth(6) = 1600:     .ColAlignment(6) = flexAlignCenterCenter       '[2] 오더 번호(15)
        
        .TextArray(7) = "온도":           .ColWidth(7) = 550:     .ColAlignment(7) = flexAlignCenterCenter       '[4] 거래처명
        .TextArray(8) = "속도":           .ColWidth(8) = 550:     .ColAlignment(8) = flexAlignCenterCenter       '[4] 거래처명
        .TextArray(9) = "Over" & vbCrLf & "Feed":       .ColWidth(9) = 500:     .ColAlignment(9) = flexAlignCenterCenter       '[4] 거래처명
        .TextArray(10) = "위사" & vbCrLf & "밀도":      .ColWidth(10) = 500:    .ColAlignment(10) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(11) = "Setting" & vbCrLf & "구분":   .ColWidth(11) = 500:    .ColAlignment(11) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(12) = "작업" & vbCrLf & "조건":      .ColWidth(12) = 650:    .ColAlignment(12) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(13) = "불량명":        .ColWidth(13) = 1300:    .ColAlignment(13) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(14) = "건조정도":      .ColWidth(14) = 900:    .ColAlignment(14) = flexAlignCenterCenter      '[4] 거래처명
        .ScrollBars = flexScrollBarBoth
        
        '태을염직의 경우 건조정도 항목이 없음
        
        Call ColHiddenFalse(GrdCondition, .Cols)
        .ColHidden(2) = True
        .ColHidden(13) = True
        .ColHidden(14) = True
        .ColHidden(11) = True
        
        .Redraw = True
    End With
End Sub

'-- 텐터기(  F0 )  데이터 Select
Sub FillGrdTenter()
    
    'Tenter GRID에 TITLE 나타내기
    Dim oInstCPBPre As PlusLib2.CInstCondition
    Dim rs As ADODB.Recordset
    Dim sPattern$, sPatternID$, II%
    
    II = 1
    
    Screen.MousePointer = vbHourglass

  '  On Error GoTo ErrHandler
    
    Call TitGrdTenter

    Set oInstCPBPre = New PlusLib2.CInstCondition
    oInstCPBPre.Connection = g_adoCon
    
    Set rs = oInstCPBPre.GetInstRecord("xp_InstCondi_sDraftTenter" _
                                    , mOrderID _
                                    , IIf(opOrderID.Value = True, 1, 2) _
                                    , IIf(opMachine.Value = True, 1, 0) _
                                    , Trim$(cboTenterProc.Text) _
                                    , val(txtSelRecs))
                                    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            GrdCondition.AddItem "" & vbTab & MakeOrderID(rs("OrderID"), OM_EXPAND) & vbTab & rs("OrderNO") & vbTab & rs("KCustom") & vbTab & _
                               rs("Article") & vbTab & rs("Process") & vbTab & MakeDate(DF_LONG, rs("IndiDate")) & " " & Format$(rs("IndiTime"), "0#:##") & vbTab & _
                               rs("Temper") & vbTab & rs("Velocity") & vbTab & rs("OverFeed") & vbTab & _
                               rs("Density") & vbTab & rs("SettingClss") & vbTab & _
                               rs("WorkCond") & vbTab & rs("DefectName") & vbTab & rs("DryID")
                               
            rs.MoveNext
            II = II + 1
            If II > val(txtSelRecs) Then
                Exit Do
            End If
        Loop
        rs.Close
        Set rs = Nothing
        
    End If
    
    Set oInstCPBPre = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:

End Sub


Sub TitGrdPeach()
    Call SetVSFlexGrid(GrdCondition)
    With GrdCondition
        .Redraw = False
        .Cols = 17
            
        .TextArray(1) = "Order No":       .ColWidth(1) = 1300:     .ColAlignment(1) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(2) = "관리번호":       .ColWidth(2) = 1300:     .ColAlignment(2) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(3) = "거래처":         .ColWidth(3) = 1600:     .ColAlignment(3) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(4) = "품명":           .ColWidth(4) = 1300:     .ColAlignment(4) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(5) = "공정":           .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignCenterCenter      '[2] 오더 번호(15)
        .TextArray(6) = "지시일시":       .ColWidth(6) = 1600:     .ColAlignment(6) = flexAlignCenterCenter      '[2] 오더 번호(15)
        
        .TextArray(7) = "속도":           .ColWidth(7) = 900:     .ColAlignment(7) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(8) = "페파본1":        .ColWidth(8) = 900:     .ColAlignment(8) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(9) = "페파본2":        .ColWidth(9) = 900:     .ColAlignment(9) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(10) = "페파본3":       .ColWidth(10) = 900:    .ColAlignment(10) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(11) = "페파본4":       .ColWidth(11) = 900:    .ColAlignment(11) = flexAlignCenterCenter      '[4] 거래처명
        
        .TextArray(12) = "밀도":          .ColWidth(12) = 900:    .ColAlignment(12) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(13) = "장력":          .ColWidth(13) = 900:    .ColAlignment(13) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(14) = "압력1":         .ColWidth(14) = 900:    .ColAlignment(14) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(15) = "압력2":         .ColWidth(15) = 900:    .ColAlignment(15) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(16) = "압력3":         .ColWidth(16) = 900:    .ColAlignment(16) = flexAlignCenterCenter      '[4] 거래처명
        .ScrollBars = flexScrollBarBoth
        
        Call ColHiddenFalse(GrdCondition, .Cols)
        .ColHidden(2) = True
        .Redraw = True
    End With
End Sub
Sub ColHiddenFalse(ByVal oGrid As VSFlexGrid, ByVal iCols As Integer)
    Dim iCount As Integer
    For iCount = 0 To iCols - 1
        oGrid.ColHidden(iCount) = False
    Next iCount
End Sub
'-- Peach ( G0 ) 데이터 Select
Sub FillGrdPeach()
    '---
    Call TitGrdPeach
    Dim II%
    
    II = 1
    
    '---- Peach  GRID에 TITLE 나타내기
    Dim oInstCPBPre As PlusLib2.CInstCondition
    Dim rs As ADODB.Recordset
    Dim sPattern$, sPatternID$

    Screen.MousePointer = vbHourglass

 '   On Error GoTo ErrHandler

    Set oInstCPBPre = New PlusLib2.CInstCondition
    oInstCPBPre.Connection = g_adoCon
    
    Set rs = oInstCPBPre.GetInstRecord("xp_InstCondi_sDraftPeach" _
                                    , mOrderID _
                                    , IIf(opOrderID.Value = True, 1, 2) _
                                    , IIf(opMachine.Value = True, 1, 0) _
                                    , Trim$(cboPeachProc.Text) _
                                    , val(txtSelRecs))
                                    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            GrdCondition.AddItem "" & vbTab & MakeOrderID(rs("OrderID"), OM_EXPAND) & vbTab & rs("OrderNO") & vbTab & rs("KCustom") & vbTab & _
                               rs("Article") & vbTab & rs("Process") & vbTab & MakeDate(DF_LONG, rs("IndiDate")) & " " & Format$(rs("IndiTime"), "0#:##") & vbTab & _
                               rs("Velocity") & vbTab & rs("PePaBon1") & vbTab & rs("PePaBon2") & vbTab & rs("PePaBon3") & vbTab & rs("PePaBon4") & vbTab & _
                               rs("Density") & vbTab & rs("Tension") & vbTab & _
                               rs("Pressure1") & vbTab & rs("Pressure2") & vbTab & rs("Pressure3")
                               
            rs.MoveNext
            II = II + 1
            If II > val(txtSelRecs) Then
                Exit Do
            End If
        Loop
        rs.Close
        Set rs = Nothing
        
    End If
    
    Set oInstCPBPre = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHandler:
End Sub
Sub TitGrdDry()
    Call SetVSFlexGrid(GrdCondition)
    With GrdCondition
        .Redraw = False
        .Cols = 12
            
        .TextArray(1) = "Order No":       .ColWidth(1) = 1300:      .ColAlignment(1) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(2) = "관리번호":       .ColWidth(2) = 1300:      .ColAlignment(2) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(3) = "거래처":         .ColWidth(3) = 1800:      .ColAlignment(3) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(4) = "품명":           .ColWidth(4) = 2100:      .ColAlignment(4) = flexAlignCenterCenter    '[1] 관리 번호(9)
        .TextArray(5) = "공정":           .ColWidth(5) = 800:      .ColAlignment(5) = flexAlignCenterCenter      '[2] 오더 번호(15)
        .TextArray(6) = "지시일시":       .ColWidth(6) = 1600:      .ColAlignment(6) = flexAlignCenterCenter      '[2] 오더 번호(15)
        
        .TextArray(7) = "온도":           .ColWidth(7) = 900:       .ColAlignment(7) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(8) = "속도":           .ColWidth(8) = 900:       .ColAlignment(8) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(9) = "Over" & vbCrLf & "Feed":       .ColWidth(9) = 450:       .ColAlignment(9) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(10) = "불량명":        .ColWidth(10) = 900:      .ColAlignment(10) = flexAlignCenterCenter      '[4] 거래처명
        .TextArray(11) = "건조정도":      .ColWidth(11) = 900:      .ColAlignment(11) = flexAlignCenterCenter      '[4] 거래처명
        
        Call ColHiddenFalse(GrdCondition, .Cols)
        
        '태을염직의 경우 건조정도 항목이 없음
        .ColHidden(2) = True
        .ColHidden(10) = True
        .ColHidden(11) = True
        
        .Redraw = True
        .ScrollBars = flexScrollBarBoth
        
    End With
End Sub

'-- Dry(  K0 )  데이터 Select
Sub FillGrdDry()
    '--- Grid Dry Title설정
    Call TitGrdDry
    Dim II%
    
    '---- Dry  GRID에 TITLE 나타내기
    Dim oInstCPBPre As PlusLib2.CInstCondition
    Dim rs As ADODB.Recordset
    Dim sPattern$, sPatternID$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oInstCPBPre = New PlusLib2.CInstCondition
    oInstCPBPre.Connection = g_adoCon
    
    Set rs = oInstCPBPre.GetInstRecord("xp_InstCondi_sDraftDry" _
                                    , mOrderID _
                                    , IIf(opOrderID.Value = True, 1, 2) _
                                    , IIf(opMachine.Value = True, 1, 0) _
                                    , Trim$(cboDryProc.Text) _
                                    , val(txtSelRecs))

    II = 1
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            GrdCondition.AddItem "" & vbTab & MakeOrderID(rs("OrderID"), OM_EXPAND) & vbTab & rs("OrderNO") & vbTab & rs("KCustom") & vbTab & _
                               rs("Article") & vbTab & rs("Process") & vbTab & MakeDate(DF_LONG, rs("IndiDate")) & " " & Format$(rs("IndiTime"), "0#:##") & vbTab & _
                               rs("Temper") & vbTab & rs("Velocity") & vbTab & rs("OverFeed") & vbTab & rs("KDefect") & vbTab & rs("DryID")
                               
            rs.MoveNext
            II = II + 1
            If II > val(txtSelRecs) Then
                Exit Do
            End If
        Loop
        rs.Close
        Set rs = Nothing
        
    End If
    
    Set oInstCPBPre = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    
End Sub


Private Sub txtCPBPersonID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub


Private Sub txtCPBRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtCPBVelocity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachDensity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPePaBon1_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPePaBon2_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPePaBon3_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPePaBon4_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPressure1_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPressure2_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachPressure3_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachTension_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtPeachVelocity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtRefinePersonID_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtRefineRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtRefineSettingClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtRefineTemper_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub


Private Sub txtRefineVelocity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Index = 1 Then
        Call ReturnCode(LG_CUSTOM, , False, txtSearch(1))
        
        Call NextFocus
    ElseIf KeyAscii = vbKeyReturn And Index = 2 Then
        Call NextFocus
    End If
    
End Sub

Private Sub txtSelRecs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FillGrdCondition
    End If
End Sub

Private Sub txtSelRecs_LostFocus()
    If val(txtSelRecs) > 0 Then
        Call FillGrdCondition
    End If
End Sub

Private Sub txtTenterDensity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtTenterOverFeed_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub


Private Sub txtTenterRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub


Private Sub txtTenterSettingClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub



Private Sub txtTenterTemper_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtTenterVelocity_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub
