VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffINReturn 
   Caption         =   "생지 반품 관리"
   ClientHeight    =   9255
   ClientLeft      =   2775
   ClientTop       =   3255
   ClientWidth     =   15195
   Icon            =   "frmStuffINReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15195
   Begin VB.TextBox txtOLDStuff 
      Height          =   315
      Left            =   30
      TabIndex        =   87
      Top             =   8700
      Visible         =   0   'False
      Width           =   1995
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   4605
      Left            =   30
      TabIndex        =   69
      Top             =   1050
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   8123
      _Version        =   196609
      Caption         =   "SSPanel6"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optGroup 
         Caption         =   "거래처별"
         Height          =   330
         Index           =   1
         Left            =   1110
         Style           =   1  '그래픽
         TabIndex        =   73
         Top             =   30
         Width           =   990
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "오더별"
         Height          =   330
         Index           =   0
         Left            =   30
         Style           =   1  '그래픽
         TabIndex        =   72
         Top             =   30
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "축소"
         Height          =   345
         Index           =   1
         Left            =   3990
         TabIndex        =   71
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "확장"
         Height          =   345
         Index           =   0
         Left            =   3210
         TabIndex        =   70
         Top             =   30
         Width           =   765
      End
      Begin VSFlex7LCtl.VSFlexGrid grdGroup 
         Height          =   3810
         Left            =   0
         TabIndex        =   74
         Top             =   390
         Width           =   9390
         _cx             =   16563
         _cy             =   6720
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
      Begin VSFlex7LCtl.VSFlexGrid grdTotal 
         Height          =   330
         Left            =   0
         TabIndex        =   75
         Top             =   4200
         Width           =   9390
         _cx             =   16563
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   6300
         TabIndex        =   76
         Top             =   30
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   609
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No"
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   90
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   3
            Left            =   1380
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   90
            Width           =   1140
         End
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   2865
      Left            =   30
      TabIndex        =   66
      Top             =   5670
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   5054
      _Version        =   196609
      Caption         =   "SSPanel3"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton Command1 
         Caption         =   "다른색상입력"
         Height          =   345
         Left            =   13530
         TabIndex        =   67
         Top             =   30
         Visible         =   0   'False
         Width           =   1545
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   8
         Left            =   30
         TabIndex        =   68
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "반품절수"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdStuffINSub 
         Height          =   2430
         Left            =   30
         TabIndex        =   16
         Top             =   390
         Width           =   15060
         _cx             =   26564
         _cy             =   4286
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
   Begin Threed.SSPanel SSPanel5 
      Height          =   1035
      Left            =   8550
      TabIndex        =   54
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1826
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   795
         Index           =   3
         Left            =   2610
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   14
         ToolTipText     =   "자료 저장"
         Top             =   90
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   795
         Index           =   0
         Left            =   4200
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         ToolTipText     =   "자료 추가"
         Top             =   90
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   795
         Index           =   2
         Left            =   5790
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   57
         ToolTipText     =   "자료 삭제"
         Top             =   90
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   795
         Index           =   1
         Left            =   4995
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   56
         ToolTipText     =   "자료 수정"
         Top             =   90
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   795
         Index           =   4
         Left            =   3405
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   55
         ToolTipText     =   "자료 취소"
         Top             =   90
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   450
         Left            =   30
         TabIndex        =   58
         Top             =   405
         Visible         =   0   'False
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   794
         _Version        =   196609
         BackColor       =   12648447
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlChoice 
      Height          =   1035
      Left            =   30
      TabIndex        =   32
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   1826
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox CboStuffClss2 
         Height          =   300
         Left            =   1470
         Style           =   2  '드롭다운 목록
         TabIndex        =   61
         Top             =   750
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.ComboBox cboOrderID 
         Height          =   300
         Left            =   1440
         Style           =   2  '드롭다운 목록
         TabIndex        =   41
         Top             =   1080
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   4860
         TabIndex        =   40
         Top             =   690
         Width           =   1935
      End
      Begin VB.TextBox txtArticle 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4860
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   1
         Left            =   4860
         TabIndex        =   36
         Top             =   30
         Width           =   1935
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   30
         MousePointer    =   99  '사용자 정의
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   30
         Width           =   615
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "전일"
         Height          =   315
         Index           =   0
         Left            =   30
         MousePointer    =   99  '사용자 정의
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   720
         Left            =   7320
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   33
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   810
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   0
         Left            =   6810
         TabIndex        =   37
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
         Left            =   3540
         TabIndex        =   42
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   43
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   11
         Left            =   3540
         TabIndex        =   44
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품     명"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   1065
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   2
         Left            =   6810
         TabIndex        =   39
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
         Left            =   1950
         TabIndex        =   46
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   120389633
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1950
         TabIndex        =   47
         Top             =   390
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   120389633
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   660
         TabIndex        =   48
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "반품일자"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   49
            Top             =   60
            Value           =   1  '확인
            Width           =   1095
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   16
         Left            =   3540
         TabIndex        =   50
         Top             =   690
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
            TabIndex        =   51
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   150
         TabIndex        =   52
         Top             =   1080
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "확정구분"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   150
         TabIndex        =   62
         Top             =   750
         Visible         =   0   'False
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
            TabIndex        =   63
            Top             =   60
            Value           =   1  '확인
            Width           =   1065
         End
      End
   End
   Begin Threed.SSCommand PrnOK 
      Height          =   690
      Left            =   11850
      TabIndex        =   31
      Top             =   8550
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "거래명세서 "
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13530
      TabIndex        =   18
      Top             =   8550
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
      Left            =   4860
      TabIndex        =   19
      Top             =   8640
      Visible         =   0   'False
      Width           =   3690
      _cx             =   6509
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
   Begin Threed.SSPanel pnlData 
      Height          =   4590
      Left            =   9450
      TabIndex        =   20
      Top             =   1050
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8096
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboSubulWidth 
         Height          =   300
         Left            =   4230
         Style           =   2  '드롭다운 목록
         TabIndex        =   83
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   4140
         TabIndex        =   8
         Top             =   1785
         Width           =   1170
      End
      Begin VB.TextBox txtOrderNO 
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   780
         Width           =   2340
      End
      Begin VB.TextBox txtOrderID 
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   450
         Width           =   2340
      End
      Begin VB.TextBox TxtArticleID2 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   1560
         TabIndex        =   6
         Top             =   1440
         Width           =   2340
      End
      Begin VB.TextBox txtCustomID 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   1560
         TabIndex        =   5
         Top             =   1125
         Width           =   2340
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   1785
         Width           =   1170
      End
      Begin VB.TextBox txtRemark 
         Height          =   300
         Left            =   1560
         ScrollBars      =   2  '수직
         TabIndex        =   13
         Top             =   3450
         Width           =   4125
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         IMEMode         =   10  '한글 
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   2130
         Width           =   2340
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1560
         Style           =   2  '드롭다운 목록
         TabIndex        =   10
         Top             =   2475
         Width           =   1395
      End
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   14
         Left            =   1560
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   3105
         Width           =   1395
      End
      Begin VB.ComboBox CboOrderFlag 
         Height          =   300
         Left            =   1560
         Style           =   2  '드롭다운 목록
         TabIndex        =   11
         Top             =   2790
         Width           =   1395
      End
      Begin VB.ComboBox CboStuffClss 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3240
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   60
         Width           =   1455
      End
      Begin VB.TextBox txtStuffSeq 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   60
         Width           =   765
      End
      Begin VB.TextBox txtThreadName 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   4110
         TabIndex        =   15
         Top             =   2460
         Visible         =   0   'False
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   1
         Top             =   45
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyy-MM-dd (ddd)"
         Format          =   120389635
         CurrentDate     =   37068
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   30
         TabIndex        =   22
         Top             =   3450
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "비고 사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   30
         TabIndex        =   23
         Top             =   1785
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "반품절수"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   2790
         TabIndex        =   24
         Top             =   1785
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "반품수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   13
         Left            =   3240
         TabIndex        =   25
         Top             =   2820
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "사종"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   14
         Left            =   30
         TabIndex        =   26
         Top             =   1470
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "품명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   15
         Left            =   30
         TabIndex        =   27
         Top             =   780
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "OrderNO"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   30
         TabIndex        =   28
         Top             =   2130
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "반품처명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   12
         Left            =   30
         TabIndex        =   29
         Top             =   1125
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "거래처"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   30
         TabIndex        =   30
         Top             =   450
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "관리 번호"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   30
         TabIndex        =   59
         Top             =   2475
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "반품단위"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   12
         Left            =   30
         TabIndex        =   60
         Top             =   3105
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
         Caption         =   "가공 구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   720
         Left            =   30
         TabIndex        =   64
         Top             =   3780
         Width           =   5595
         _cx             =   9869
         _cy             =   1270
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
         Height          =   300
         Index           =   17
         Left            =   30
         TabIndex        =   65
         Top             =   2790
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "사용구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   285
         Left            =   4170
         TabIndex        =   79
         Top             =   2970
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkAddClss 
            Caption         =   "추가분"
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1095
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3870
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1125
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   4
         Left            =   3870
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   3
         Left            =   3870
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   780
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   7
         Left            =   30
         TabIndex        =   85
         Top             =   60
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkDateUPD 
            Caption         =   "반품일자수정"
            Height          =   195
            Left            =   60
            TabIndex        =   86
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5670
         Y1              =   390
         Y2              =   390
      End
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   10140
      TabIndex        =   84
      Top             =   8550
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmStuffINReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'변경이력
' 요청 ID : S_201105_태을염직_03
' 요청자 : 김대진 대리
' 요청내용 : 생지반품 명세서 A4용지용 추가
' 변경일자 : 2011.05.31
' 변경내용 : 폼지 대신 엑셀양식으로 대체
'
' 요청 ID : S_201205_태을염직_01
' 요청자 : 김대진 대리
' 요청내용 : 생지반품일자 변경 가능하게
' 변경일자 : 2012.05.21
' 변경내용 : 생지입고처럼  일자수정 추가
'
' 요청 ID : S_201303_태을염직_01
' 요청자 : 김대진 대리
' 요청내용 : 수량이 10만이상시 오류
' 변경일자 : 2013.03.19
' 변경내용 : integer에서 long으로 변경
'
'******************************************************************************************
'
Option Explicit

'Dim m_iFlag As String * 1

Dim m_iFlag As Integer
Dim m_bGroupClss As Boolean     '거래처별, 오더별 Grid 구분

Private Const COMBOLIST = "1.정단|2.난단|"

' S_201105_태을염직_03 에 의한 추가
Private Const REPORTFILE   As String = "\Report\ReturnRoll.xls"
Private Const REPORTFILE1  As String = "\Report\TmpReturnRoll.xls"

Private Const EXCEL_ROLL_ROW As Integer = 41


Private Const LIMIT_WIDTH1 = 1640
Private Const LIMIT_WIDTH2 = 2100
Private Const LIMIT_WIDTH3 = 560
Private Const LIMIT_WIDTH4 = 2000
Private Const LIMIT_ROW1 = 11
Private Const LIMIT_ROW2 = 28
Private Const LIMIT_ROW3 = 9
Private m_bSortForward As Boolean
Private m_StuffDate As String, m_StuffClss As String, m_StuffSeq As Integer


' 수불명세서 Form에서 Call하기
Public Sub LoadStuffIN(ByVal StuffINKey As String)
    
    Me.Show
    chkSearch(3).Value = False
    
    
    Call MakeStuffKey(StuffINKey, m_StuffDate, m_StuffClss, m_StuffSeq)
    Call FillGridStuffSub(m_StuffDate, m_StuffClss, m_StuffSeq)
    
    cmdFind(1).Enabled = True
    cmdFind(3).Enabled = True
    cmdFind(4).Enabled = True

End Sub

Private Sub cboName_KeyPress(Index As Integer, KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub



Private Sub CboStuffClss_Click()
    '--- 반품이면
    If CboStuffClss.ItemData(CboStuffClss.ListIndex) = 3 Then
  '      fraData.Enabled = True
        
        With grdStuffINSub
            .Rows = grdStuffINSub.FixedRows
           .AddItem CStr(.Rows)
        End With
    Else
'        fraData.Enabled = False
    End If
End Sub

Private Sub CboStuffClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cmdShink_Click(Index As Integer)

End Sub

'S_201205_태을염직_01 에 의한 추가
Private Sub chkDateUPD_Click()
    Dim vStuffDate As Variant
    If chkDateUPD.Value = vbChecked Then
        If m_iFlag = ID_UPDATE Then
            dtpDate(2).Enabled = True
      '      CboStuffClss.Enabled = True
            txtStuffSeq.Text = ""
        End If
    Else
    
        If m_iFlag = ID_UPDATE Then
            dtpDate(2).Enabled = False
            CboStuffClss.Enabled = False
            vStuffDate = Split(txtOLDStuff, "-")
            dtpDate(2) = Format(vStuffDate(0), "####-##-##")
         '   CboStuffClss.ListIndex = vStuffDate(1)
            txtStuffSeq = vStuffDate(2)
        End If
    
    End If
End Sub

'S_201105_태을염직_03 에 의한 추가
Private Sub cmdExcel_Click()

    On Error GoTo ErrHandler
    If grdGroup.Rows <= grdGroup.FixedRows Then Exit Sub
    
    '거래처 있는 경우만
    If txtCustomID.Tag <> "" Then
        Call MakeExcelPacking           '생지반품 명세서 엑셀출력
    End If


    Exit Sub

ErrHandler:

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdShrink_Click(Index As Integer)
    If Index = 0 Then
        Call SetGrdShrink(grdGroup, OM_EXPAND)
    Else
        Call SetGrdShrink(grdGroup, OM_REDUCE)
    End If


End Sub

Private Sub Command1_Click()
    With grdStuffINSub
        .Rows = .Rows + 1
        .Select .Rows - 1, 1
    End With
    grdStuffINSub.SetFocus
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
    m_bGroupClss = True
    Call InitGrid
    Call InitGroup
    
    Call SetOperate(Me)
    
    '----- 입고단위 설정
    With cboUnit
        .AddItem "YDS":  .ItemData(0) = 0
        .AddItem "MTS":  .ItemData(1) = 1
        .ListIndex = 0
    End With
    
    '----- 입고구분 설정
    With CboStuffClss
        .AddItem "1.생지":              .ItemData(0) = 1
        .AddItem "3.반품(생지)":        .ItemData(1) = 3
        .ListIndex = 1
        .Enabled = False
    End With
    
    
    '----- 검색용 입고구분 설정
    With CboStuffClss2
        .AddItem "1.생지":        .ItemData(0) = 1
        .AddItem "3.반품 생지":   .ItemData(1) = 3
        .ListIndex = 1
        .Enabled = False
    End With
    
    
    '----- 확정구분
    With cboOrderID
        .AddItem "수주확정"
        .AddItem "수주미확정"
        .ListIndex = 0
        .Visible = False
    End With
    
    '----- OrderFlag
    With CboOrderFlag
        .AddItem "0.비사용":        .ItemData(0) = 0       ' A건
        .AddItem "1.사용":          .ItemData(1) = 1       ' B건
        .ListIndex = 0
    End With
    
    Call MakeCodeCombo(cboName(14), CD_WORK)        ' 가공 구분
    
    Dim nSeq As Integer
    
    If Weekday(DateAdd("D", -1, Now)) = 1 Then
        nSeq = -2
    Else
        nSeq = -1
    End If
    
    '---- 날짜 설정
    For i = 0 To 2
        dtpDate(i) = DateAdd("D", nSeq, Now)
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
    
'    Call ClearText(txtNum, "0")
    
    
    '--- 필수입력 항목에 표시하기  거래처명
    pnlCaption(1).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(2).Picture = LoadResPicture("BASIC", vbResIcon)
'    pnlCaption(7).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(8).Picture = LoadResPicture("BASIC", vbResIcon)
'    pnlCaption(10).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(12).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(14).Picture = LoadResPicture("BASIC", vbResIcon)
    
    
    '---- 오더별 데이터 나타내기
    m_bGroupClss = True
    Call FillGridGroup(m_bGroupClss)
    Call SetToggle
    Call NonEditMode(True)
    
    txtCustom(1).Enabled = chkSearch(1).Value
    cmdFind(0).Enabled = chkSearch(1).Value
    txtArticle.Enabled = chkSearch(2).Value
    cmdFind(2).Enabled = chkSearch(2).Value
    txtSearch(3).Enabled = chkSearch(0).Value
    CboStuffClss2.Enabled = chkSearch(4).Value
    cboOrderID.Enabled = chkSearch(5).Value
    dtpDate(0).Enabled = chkSearch(3).Value
    dtpDate(1).Enabled = chkSearch(3).Value
    
    Call SetStuffWidth(cboSubulWidth)
    
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
    Dim vOLDStuffDate As Variant
    
    On Error GoTo ErrHandler
    
    Select Case m_iFlag
        Case ID_ADDNEW
            sJobFlag = "I"
        Case ID_UPDATE
            sJobFlag = "U"
    End Select
    
    'S_201205_태을염직_01 에 의한 추가
    If chkDateUPD.Value = vbChecked Then
        vOLDStuffDate = Split(txtOLDStuff, "-")
    Else
        ReDim vOLDStuffDate(3)
        vOLDStuffDate(0) = ""
        vOLDStuffDate(1) = ""
        vOLDStuffDate(2) = 0
    End If
    
    StuffClss = CboStuffClss.ItemData(CboStuffClss.ListIndex)
    
    If StuffClss = "3" And (txtNum(0) > 0 Or txtNum(1) > 0) Then
        txtNum(0) = CheckNum(txtNum(0)) * -1
        txtNum(1) = CheckNum(txtNum(1)) * -1
    End If
    
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
        .sOrderNO = txtOrderNO                           'OrderNo
        .ADDClss = chkAddClss.Value
        .sOrderFlag = Left(CboOrderFlag, 1)
        
        'S_201205_태을염직_01 에 의한 추가--------------------------------------
        .nChkDateUpd = IIf(chkDateUPD.Value = vbChecked, 1, 0)
        .sOLDDate = vOLDStuffDate(0)
        .sOLDClss = vOLDStuffDate(1)
        .nOLDSeq = vOLDStuffDate(2)
        '---------------------------------------------------
        .sSubulWidthID = Format(cboSubulWidth.ItemData(cboSubulWidth.ListIndex), "0#")
        
    End With

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Description, "frmStuffIN.SetNewData", Err.Description)

End Sub


Private Sub SetNewDataSub(SetSubData() As PlusLib2.TStuffINReturn, nSeq As Integer)
    Dim nCount%, II%, JJ%, nInt%

''Type TStuffINReturn
''    sStuffDate      As String  '[3] 반품일자
''    sStuffClss      As String   '입고구분
''    nStuffSeq       As Integer  '입고순번
''    nStuffNO        As Integer  '일련번호
''    sStuffPart      As String   '1.정단, 2.난단 구분
''    nRollNo         As Integer  '절번호
''    nQty            As String   '반품수량
''
''End Type
    
    nCount = 0
    With grdStuffINSub
        If .Rows = .FixedRows Then Exit Sub
        
        For II = .FixedRows To .Rows - 1
            nInt% = 0
            For JJ = 2 To .Cols - 3
                If .ValueMatrix(II, JJ) <> 0 Then
                    nInt = nInt + 1                 '절번호
                    ReDim Preserve SetSubData(nCount + 1)
                    
                    SetSubData(nCount).nStuffNO = .ValueMatrix(II, 0)
                    
                    If .TextMatrix(II, 1) = "1.정단" Then
                        SetSubData(nCount).sStuffPart = "1"
                        SetSubData(nCount).nRollNo = nInt
                        SetSubData(nCount).nQty = .TextMatrix(II, 2) & "*" & .TextMatrix(II, 4)
                        nCount = nCount + 1             '배열의 크기
                        Exit For
                    Else
                        SetSubData(nCount).sStuffPart = "2"
                        SetSubData(nCount).nRollNo = nInt
                        SetSubData(nCount).nQty = .TextMatrix(II, JJ)
                        nCount = nCount + 1             '배열의 크기
                    End If
                End If
            Next JJ
            
        Next II
    End With
    nSeq = nCount
    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffINReturn.SetNewDataSub", Err.Description)

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
'    Dim nRoll%, nQty%, nRollQty%       '합계
'    Dim nTotRoll%, nTotQty%
    'S_201303_태을염직_01 에 의한 수정
    Dim nRoll As Integer, nQty As Long, nRollQty As Long
    Dim nTotRoll As Long, nTotQty As Long

    
    ' Grid Text 값 계산
    With grdStuffINSub
        If .Rows = .FixedRows Then Exit Sub
        
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 1) = "1.정단" Then
                Call GetRollQty(.TextMatrix(i, 2) & .TextMatrix(i, 3) & .TextMatrix(i, 4), nRoll, nRollQty, nQty)
                nTotRoll = nTotRoll + nRoll
                nTotQty = nTotQty + nQty
            Else
                For j = 2 To .Cols - 3
                    If Len(.TextMatrix(i, j)) <> 0 Then
                        Call GetRollQty(.TextMatrix(i, j), nRoll, nRollQty, nQty)
                        nTotRoll = nTotRoll + nRoll
                        nTotQty = nTotQty + nQty
                    End If
                Next j
            End If
        Next i
    End With

    txtNum(0) = nTotRoll * -1
    txtNum(1) = nTotQty * -1
End Sub

Private Function DeleteData() As Boolean
    Dim oStuffIn As PlusLib2.cStuffIN

    On Error GoTo ErrHandler
    
'    Call FillGrid
   
    Set oStuffIn = New PlusLib2.cStuffIN
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
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim NewStuffIN   As PlusLib2.TStuffIN
    Dim StuffINReturn() As PlusLib2.TStuffINReturn
    Dim StuffINSub() As PlusLib2.TStuffINSub
'    Dim StuffINDesign() As PlusLib2.TDesign
    Dim nCount  As Integer
    Dim nStuffSeq As Integer, nSeq As Integer
    
    On Error GoTo ErrHandler
    
    nSeq = Abs(val(txtNum(0)))
    
    ' cStuffIn 클래스의 구조체에 값 대입
    Call SetNewData(NewStuffIN)
    
    '--- Stuffinsub 구조체에 값 넣기
    '--- 반품내역 배열에 넣기
    ReDim StuffINSub(1)
    ReDim StuffINDesign(1)
    
    nCount = 0: ReDim StuffINReturn(1)
    
    Call SetNewDataSub(StuffINReturn, nCount)
    
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    SaveData = oStuffIn.AddNewStuffIN("3", nSeq, NewStuffIN, StuffINSub, StuffINReturn, nCount)
    txtStuffSeq = nSeq
    
    Set oStuffIn = Nothing
    
    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffINReturn.SaveData", Err.Description)
    Set oStuffIn = Nothing

End Function

Sub SetKeyEdit(ByVal dEdit As Boolean)
    dtpDate(2).Enabled = dEdit
'    CboStuffClss.Enabled = dEdit
    txtStuffSeq.Enabled = dEdit
End Sub

'-----------------------------------------------------------------------
' StuffIN , StuffINSub Record 나타내기 및 StuffINSub Grid에 나타내기
' 수정 모드에서 사용함
'
'-------------- 입고데이터 나타내기  --- 수정모드
'-----------------------------------------------------------------------
Private Sub FillGridStuffSub(ByVal StuffDate As String, ByVal StuffClss As String, ByVal StuffSeq As Integer)
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim irow%, i%, j%, iCount%
    'Dim nRollvar()
    Dim nRollCnt As Integer, nRollQty As Long, sRollStr As String, sColor As String
    
'    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    'S_201205_태을염직_01 에 의한 추가
    txtOLDStuff = StuffDate & "-" & StuffClss & "-" & StuffSeq
    
    grdStuffINSub.Rows = grdStuffINSub.FixedRows
    
    ''''' StuffIN 1개의 record을 읽어온다.
    Set rs = oStuffIn.GetStuffINOne(StuffDate, StuffClss, StuffSeq)
    
    ''''' 반품절수 데이터 읽어 오기
    Set rsData = oStuffIn.GetStuffInReturn(StuffDate, StuffClss, StuffSeq)
    
    
    dtpDate(2) = MakeDate(DF_LONG, StuffDate)
    
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
        chkAddClss.Value = val(rs!ADDClss)
        CboOrderFlag.ListIndex = val(rs!OrderFlag)
    End With
    
    rs.Close
    Set rs = Nothing
    
    '--- Order 에 대한 내용 grid에 나타내기
    Call FillStuffOrderData(txtOrderID.Text)
    
    '''''' StuffSub 데이터 나타내기
    '''''' StuffINSub 데이터를 Grd에 나타내기 위해 재 정의 한다.   예) 90 * 3의 으로 읽어옴.
        
    Dim nRoll%, nQty As Long
    With grdStuffINSub
        .Rows = .FixedRows
        
        Do Until rsData.EOF
            If rsData!StuffRollNO = 1 Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = rsData!StuffNO
                .TextMatrix(.Rows - 1, 1) = IIf(rsData!StuffPart = "1", "1.정단", "2.난단")
                
            End If
            
            If rsData!StuffPart = "1" Then
                i = InStr(1, rsData!Qty, "*")
                .TextMatrix(.Rows - 1, 2) = Left(rsData!Qty, i - 1)
                .TextMatrix(.Rows - 1, 3) = "*"
                .TextMatrix(.Rows - 1, 4) = Mid(rsData!Qty, i + 1)
                .TextMatrix(.Rows - 1, 5) = "="
                .TextMatrix(.Rows - 1, 6) = .ValueMatrix(.Rows - 1, 2) * .ValueMatrix(.Rows - 1, 4)
                .TextMatrix(.Rows - 1, .Cols - 2) = .ValueMatrix(.Rows - 1, 4)
                .TextMatrix(.Rows - 1, .Cols - 1) = .ValueMatrix(.Rows - 1, 6)
                                    
            Else
            
                .TextMatrix(.Rows - 1, rsData!StuffRollNO + 1) = rsData!Qty
                Call RowCalcQty(.Rows - 1, nRoll, nQty)
                .TextMatrix(.Rows - 1, .Cols - 2) = nRoll
                .TextMatrix(.Rows - 1, .Cols - 1) = nQty
                
            End If
            
            rsData.MoveNext
            
        Loop
    End With
    
    rsData.Close
    Set rsData = Nothing
    
    Set oStuffIn = Nothing

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
    Dim i%, iColWidth%, iCount%, nRows%
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
    
'    ' 절수, 수량 입력 Grid
'    Dim sComBoList$
'
'    '--- 사용가능한 ColorID, Color명 가져와서 Grid의 combobox의 item으로 설정
'    Dim rs As ADODB.Recordset
'
'    Set rs = GetColor
'    sComBoList$ = " " & "|"
'    Do Until rs.EOF
'        sComBoList$ = sComBoList$ & rs(0) & "|"
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
  
    '----  순번 / 정난단구분 / 마수1, .............마수10 / 합계
    '----   0   /  1         / 2 ~ 11                     / 12
    Call SetVSFlexGrid(grdStuffINSub)
    With grdStuffINSub
        .Redraw = flexRDNone
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 14
        .Rows = .FixedRows + 1
        .TextArray(0) = ""
        .ColWidth(0) = 300
        .SelectionMode = flexSelectionFree
        
        
         nRows = 0
        .TextMatrix(nRows, 0) = "순번":    .ColWidth(0) = 800:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(nRows, 1) = "구분":    .ColWidth(1) = 1200:        .ColAlignment(1) = flexAlignCenterCenter
        
        iColWidth = Int((grdStuffINSub.Width - .ColWidth(0) - .ColWidth(1)) / 12)
        iCount = 1: i = 2
        
        ' 2번째 컬럼부터 12개의 컬럼 생성
        Do While iCount < 13
            .TextMatrix(nRows, i) = iCount
            .ColWidth(i) = iColWidth
            .ColAlignment(i) = flexAlignCenterCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            iCount = iCount + 1
            i = i + 1
        Loop
        
        .TextMatrix(nRows, .Cols - 2) = "절수"
        .TextMatrix(nRows, .Cols - 1) = "수량"
        
        .ColComboList(1) = COMBOLIST
        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightAlways
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect

    End With

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
            Call ReturnRef(LG_CUSTOM, , False, txtCustom(1))
        Case 1                '[2] 거래처 코드
            Call ReturnRef(LG_CUSTOM, , False, txtCustomID)
        Case 2                '[3] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Case 3                '[4] 오더 코드
            Call ReturnCode(LG_ORDER, , False, txtOrderNO)
            If Trim(txtOrderNO.Tag) = "" Then
                txtOrderNO.Text = ""
                txtOrderID.Text = ""
        '        txtCustomID.Text = ""
        '        TxtArticleID2.Text = ""
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
            Call ReturnRef(LG_ARTICLE, , False, TxtArticleID2)
    End Select
End Sub


Sub NonEditMode(ByVal pMode As Boolean)
    pnlChoice.Enabled = pMode
    grdGroup.Enabled = pMode
    pnlData.Enabled = Not pMode
    grdStuffINSub.Enabled = Not pMode
    
'    fraData.Enabled = False
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
            Call NonEditMode(False)
            
            cboName(14).ListIndex = 0
            cmdFind(1).Enabled = True
            cmdFind(3).Enabled = True
            cmdFind(4).Enabled = True
            
            txtCustomID.Enabled = True
            TxtArticleID2.Enabled = True
            
            grdData.Rows = grdData.FixedRows
            PrnOK.Enabled = False
            dtpDate(2).SetFocus
        Case ID_UPDATE

            If val(txtStuffSeq) > 0 Then
                Call NonEditMode(False)
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
            '    Call SetClearEdit
                Call cmdSearch_Click
                m_iFlag = -1
                Call ChangeMode(Me, True)
                Call cmdShrink_Click(0)
                Call NonEditMode(True)
                PrnOK.Enabled = True
            End If
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
            Call NonEditMode(True)
            Call grdGroup_RowColChange
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
    Dim nSeq As Integer
    
    Call ClearScreen(Me, "pnlData")
    
    'S_201205_태을염직_01 에 의한 추가
    chkDateUPD.Value = vbUnchecked
    txtOLDStuff.Text = ""
    
    cmdFind(1).Enabled = False
    cmdFind(3).Enabled = False
    cmdFind(4).Enabled = False
    
    txtCustomID.Tag = ""
    TxtArticleID2.Tag = ""
    
    If Weekday(DateAdd("D", -1, Now)) = 1 Then
        nSeq = -2
    Else
        nSeq = -1
    End If
    
    dtpDate(2) = DateAdd("D", nSeq, Now)
    
    grdData.Rows = grdData.FixedRows
    grdData.HighLight = flexHighlightNever
    grdStuffINSub.Rows = grdStuffINSub.FixedRows
    CboStuffClss.ListIndex = 1
    CboOrderFlag.ListIndex = 0
    chkAddClss.Value = 0
    '.Value = 0
    Call SetKeyEdit(True)
End Sub

Private Sub cmdSearch_Click()
    If optGroup(0) Then
        m_bGroupClss = True
        Call FillGridGroup(True)
    Else
        m_bGroupClss = False
    End If
    
    Call FillGridGroup(m_bGroupClss)
End Sub
Sub FillGrdStuffIN()
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim iTop(2) As Integer
    Dim i%

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.cStuffIN
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
                .AddItem "" & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & Trim(rs!OrderNo) & vbTab & _
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
            .TextArray(15) = "OrderID":                         .ColAlignment(15) = flexAlignCenterCenter
            .TextArray(16) = "Custom1":                         .ColAlignment(17) = flexAlignCenterCenter
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
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim lNowRow%, sUnit$
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set oStuffIn = New PlusLib2.cStuffIN
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
        
     '   .AddItem MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                 MakeDate(DF_LONG, rs!AcptDate) & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                 rs!WorkName & vbTab & rs!Width & vbTab & MakeRating(rs!ChunkRate, rs!LossRate) & vbTab & _
                 CheckNum(rs!ColorCnt) & vbTab & SetCurrency(CheckNum(rs!OrderQty)) & vbTab & _
                 SetCurrency(CheckNum(rs!InQty)) & vbTab & rs!UnitClss & vbTab & CheckNum(rs!InRoll)
        
        txtCustomID.Tag = CheckNull(rs!CustomID)
        txtCustomID.Text = CheckNull(rs!kCustom)
        TxtArticleID2.Text = CheckNull(rs!Article)
        TxtArticleID2.Tag = CheckNull(rs!ArticleID)
        txtOrderNO.Text = CheckNull(rs!OrderNo)
        cboSubulWidth.ListIndex = FindComboBox(cboSubulWidth, rs!SubulWidthID)
        
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
    Call ErrorBox(Err.Number, "frmStuffINReturn.FillStuffOrderData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Function

Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, irow As Integer, iLvl As Integer)
    With oFlex
        '----  iRow을 subTotal Group으로 설정
        .IsSubtotal(irow) = True
        
        '----  iRow행을 subTotal Group의 level설정
        .RowOutlineLevel(irow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, irow, 0, irow, .Cols - 1) = &HE0E0E0
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, irow, 0, irow, .Cols - 1) = &HE0E0E0
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
    
    PrnOK.Enabled = True
    cmdFind(1).Enabled = True
    cmdFind(3).Enabled = True
    cmdFind(4).Enabled = True
    
End Sub

Private Sub grdGroup_RowColChange()
    With grdGroup
        If .Rows <= .FixedRows Then Exit Sub
        
        If .IsSubtotal(.Row) Then
            Call SetClearEdit
        Else
            Call GetStuffData
        End If
    End With
End Sub

Private Sub grdStuffINSub_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim nRoll As Integer, nQty As Long
    With grdStuffINSub
        If Row < .FixedRows Then Exit Sub
        
        Select Case .TextMatrix(.Row, 1)
            Case "1.정단"
                .TextMatrix(.Row, 3) = "*"
                .TextMatrix(.Row, 5) = "="
                If .ValueMatrix(.Row, 2) <> 0 And .ValueMatrix(.Row, 4) <> 0 Then
                    .TextMatrix(.Row, 6) = .ValueMatrix(.Row, 2) * .ValueMatrix(.Row, 4)
                    .TextMatrix(Row, .Cols - 2) = .ValueMatrix(Row, 4)
                    .TextMatrix(Row, .Cols - 1) = .ValueMatrix(Row, 6)
                End If
                
            Case "2.난단"
            
               Call RowCalcQty(Row, nRoll, nQty)
               .TextMatrix(Row, .Cols - 2) = nRoll
               .TextMatrix(Row, .Cols - 1) = nQty
            Case Else
                MsgBox (" 원단구분을 반드시 선택하십시오.")
                Exit Sub
        End Select
        
        Select Case Col
            Case 2
                If .TextMatrix(Row, 1) = "1.정단" Then
                    .Select Row, 4
                Else
                    .Select Row, Col + 1
                End If
            Case 4
                If .TextMatrix(Row, 1) = "1.정단" Then
                    If Row = .Rows - 1 Then
                        .AddItem CStr(.Rows)
                        .Select .Rows - 1, 1
                    Else
                        .Select Row + 1, 1
                    End If
                Else
                    .Select Row, Col + 1
                End If
                
            
            Case 11 To 13
                If Row = .Rows - 1 Then
                    .AddItem CStr(.Rows)
                    .Select .Rows - 1, 1
                Else
                    .Select Row + 1, 1
                End If
            Case Else
                .Select Row, Col + 1
        End Select
    End With
    
    Call CalcQty
End Sub

Sub RowCalcQty(ByVal Row As Integer, ByRef nRollCnt As Integer, ByRef nRollQty As Long)
    Dim nRoll As Integer, II As Integer
    Dim nQty As Long
    'S_201303_태을염직_01 에 의한 수정(integer 에서 Long 로 변경)
    nRoll = 0: nQty = 0
    With grdStuffINSub
        For II = 2 To 11
            If .ValueMatrix(Row, II) <> 0 Then
                nRoll = nRoll + 1
                nQty = nQty + .ValueMatrix(Row, II)
            End If
        Next II
    End With
    nRollCnt = nRoll
    nRollQty = nQty
End Sub


'*******************************************************************************************
'--- 생지 입고 오더별, 거래처별 조회
'*******************************************************************************************
Private Sub FillGridGroup(Optional NewValue As Boolean = True)
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim iTop(2) As Integer, nTop%, iSubRow As Integer
    Dim i%, xpName As String
    Dim nCheckNon As Integer
    Dim nTotOrderQty As Long, nTotRoll As Long, nTotQty As Long, nTotColorQty As Long
    Dim StuffClss As String

'    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.cStuffIN
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
                                , IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0), txtArticle.Tag _
                                , IIf(chkSearch(4) = vbChecked, 1, 0), StuffClss _
                                , IIf(chkSearch(0) = vbChecked, 1, 0), txtSearch(3).Text _
                                , nCheckNon, 0)

    Set oStuffIn = Nothing
    
    If rs.RecordCount = 0 Then
        grdGroup.Rows = grdGroup.FixedRows
        Exit Sub
    End If
    
    Call InitGroup(NewValue)
    
    nTotOrderQty = 0: nTotRoll = 0: nTotQty = 0: nTotColorQty = 0: iSubRow = 0
    
    '------- 오더별 집계 조회
    If NewValue Then
        
        With grdGroup
            .Redraw = flexRDNone
            .Rows = .FixedRows
            Do Until rs.EOF
                
                '---- 첫번째 그룹설정 (OrderID)
                If Trim(rs!OrderID) & Trim$(rs!Custom1) & Trim$(rs!Article) <> _
                   Trim(.TextMatrix(iSubRow, 15)) & Trim(.TextMatrix(iSubRow, 16)) & Trim(.TextMatrix(iSubRow, 6)) Then
                    
                    .AddItem " "
                    iSubRow = .Rows - 1
                    .TextMatrix(.Rows - 1, 2) = IIf(Trim(rs!OrderID) = "*", "", MakeOrderID(rs!OrderID, OM_EXPAND))
                    .TextMatrix(.Rows - 1, 3) = IIf(Trim(rs!OrderNo) = "", "", rs!OrderNo)
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
                    .TextMatrix(.Rows - 1, 16) = rs!Custom1
                    .TextMatrix(.Rows - 1, 17) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
'                    Call GridCollapse(grdGroup, nTop)
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
                .TextMatrix(.Rows - 1, 16) = rs!Custom1
                .TextMatrix(.Rows - 1, 17) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                         
                '-------입고절수 , 입고수량 Order별로 합계
                .TextMatrix(iTop(1), 12) = SetCurrency(.TextMatrix(iTop(1), 12) + rs!StuffRoll)
                .TextMatrix(iTop(1), 13) = SetCurrency(.TextMatrix(iTop(1), 13) + rs!StuffQty)
                nTotRoll = nTotRoll + CheckNum(rs!StuffRoll)
                nTotQty = nTotQty + CheckNum(rs!StuffQty)
    
                rs.MoveNext
            Loop
            
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
                    
'                    Call GridCollapse(grdGroup, nTop)
                    nTop = .Rows - 1
                    
                    iTop(1) = .Rows - 1
                End If
                
                '--- 두번째 그룹설절 OrderID 확인
                If Trim(rs!OrderID) & Trim(rs!Article) <> Trim(.TextMatrix(.Rows - 1, 17)) & Trim(.TextMatrix(.Rows - 1, 6)) Then
                    .AddItem ""
                    
                    .TextMatrix(.Rows - 1, 4) = MakeOrderID(rs!OrderID, OM_EXPAND)
                    .TextMatrix(.Rows - 1, 5) = rs!OrderNo
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
    
    Call SetGrdShrink(grdGroup, OM_EXPAND)
    
    
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

Private Sub grdStuffINSub_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdStuffINSub
        If KeyCode = vbKeyDelete Then
            .RemoveItem .Row
            Call CalcQty
        End If
        If KeyCode = vbKeyDown And .Row = .Rows - 1 Then
            .AddItem ""
            .Select .Rows - 1, 1
        End If
    End With
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

Private Sub PrnOK_Click()
    Dim sPrinter As String
    Dim StuffClss As String
    
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim RsHeader As ADODB.Recordset ', RsDetail As ADODB.Recordset
    'Dim rsData As ADODB.Recordset
    'Dim nRollvar(), nCols As Integer

    On Error GoTo ErrHandler
    
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sPrinter = Printer.DeviceName
    StuffClss = CboStuffClss.ItemData(CboStuffClss.ListIndex)
    
    If CboStuffClss.Enabled = False Then
        If frmPrinter.SelectPrinter(sPrinter) Then
            If oStuffIn.GetStuffINReturnGoods(MakeDate(DF_SHORT, dtpDate(2)), StuffClss, val(txtStuffSeq), RsHeader) Then
                Call SetPrint(RsHeader, grdStuffINSub)
                Set oStuffIn = Nothing
            End If
        End If
    Else
        MsgBox ("저장된 반품만 출력 할 수 있습니다.")
        Exit Sub
    End If
    
    Call ReturnPrinter(sPrinter)
    Exit Sub
ErrHandler:
    MsgBox ("반품 명세서 출력 중 오류 발생 ")
End Sub

Private Sub txtArticle_GotFocus()
    txtArticle.IMEMode = 8

End Sub



Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Call MoveFocus(KeyAscii)
    End If

End Sub

Private Sub TxtArticleID2_GotFocus()
    TxtArticleID2.IMEMode = 8
    
End Sub

Private Sub TxtArticleID2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnRef(LG_ARTICLE, , False, TxtArticleID2)
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtCustom_GotFocus(Index As Integer)
    If Index = 1 Then
        txtCustom(Index).IMEMode = 0
    End If
End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnRef(LG_CUSTOM, , False, txtCustom(1))
        End If
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtCustomID_Change()
    txtCustom(0).Text = Trim(txtCustomID.Text)
End Sub

Private Sub txtCustomID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnRef(LG_CUSTOM, , False, txtCustomID)
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtNum_GotFocus(Index As Integer)
    Call GotFocusText(txtNum(Index))
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

Private Sub txtNum_KeyPress(Index As Integer, KeyAscii As Integer)
'    If IsNumeric(txtNum(Index)) Then
'        Select Case Index
'            Case 1
'                txtNum(Index) = Format(txtNum(Index), "#,###,##0")
'
'            Case Else
'                txtNum(Index) = SetCurrency(txtNum(Index))
'        End Select
'    Else
'        txtNum(Index).Text = 0
'    End If

    Call MoveFocus(KeyAscii)
End Sub


Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    Dim dCheck_bol As Boolean
    
    If KeyAscii = vbKeyReturn Then
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
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtOrderNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(3)
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With grdStuffINSub
            .SetFocus
            If .Rows <= .FixedRows Then
                .Rows = .FixedRows + 1
            Else
                .Select .FixedRows, 1
            End If
            
        End With
    End If
End Sub

Private Sub txtThreadName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call MoveFocus(KeyAscii)
    End If
End Sub


Private Sub SetPrint(ByVal RsHeader As Recordset, ByVal oFlex As VSFlexGrid)
    Dim intBlank$, dRoll_str As String, II%, nLinePos As Long, xPos%, JJ%
    Dim PrnDate As String, nRow%, IntFind As Integer, nPage%, nLineCnt%
    Dim vCustom(20) As String
    
    Printer.Orientation = vbPRORPortrait
    Printer.ScaleMode = vbMillimeters
'    Printer.PaperSize = vbPRPSLetter

    Dim yPos(22) As Integer
        
    yPos(0) = 93
    yPos(1) = 99
    yPos(2) = 105
    yPos(3) = 111
    yPos(4) = 118
    yPos(5) = 124
    yPos(6) = 131
    yPos(7) = 137
    yPos(8) = 144
    yPos(9) = 150
    yPos(10) = 157
    yPos(11) = 163
    yPos(12) = 169
    yPos(13) = 176
    yPos(14) = 182
    yPos(15) = 188
    yPos(16) = 194
    yPos(17) = 201
    yPos(18) = 207
    yPos(19) = 213
    yPos(20) = 219
    yPos(21) = 226
    yPos(22) = 232

    If RsHeader.RecordCount > 0 Then
        vCustom(0) = Left(CheckNull(RsHeader!CustomNo), 3) & " - " & Mid(CheckNull(RsHeader!CustomNo), 4, 2) & " - " & Right(CheckNull(RsHeader!CustomNo), 5)
        vCustom(1) = CheckNull(RsHeader!kCustom)
        vCustom(2) = CheckNull(RsHeader!Chief)
        
''        'S_201312_태을염직_99 에 의한 수정-OLD 소스
''        vCustom(3) = CheckNull(RsHeader!Address1) & " " & CheckNull(RsHeader!Address2)

        'S_201312_태을염직_99 에 의한 수정-NEW 소스
        If CheckNull(RsHeader!Address1) <> "" Then                '도로명 주소 있으면
            vCustom(3) = CheckNull(RsHeader!Address1) & " " & CheckNull(RsHeader!Address2)
        Else                                                '도로명 주소 없으면-지번주소
            vCustom(3) = CheckNull(RsHeader!AddressJiBun1) & " " & CheckNull(RsHeader!AddressJiBun2)  '거래처 주소
        End If
        
        
        vCustom(4) = CheckNull(RsHeader!Condition)          '업태
        vCustom(5) = CheckNull(RsHeader!Category)           '종목
        vCustom(6) = CheckNull(RsHeader!Article)
        vCustom(7) = CheckNull(RsHeader!OrderNo)
        vCustom(8) = SetCurrency(ChkNullValue(RsHeader!OrderQty), 0)
        vCustom(9) = CheckNull(RsHeader!Custom)
        vCustom(10) = SetCurrency(RsHeader!TotRoll, 0)
        vCustom(11) = SetCurrency(RsHeader!TotQty, 0)
        vCustom(12) = RsHeader!StuffDate & "-" & RsHeader!StuffClss & "-" & RsHeader!StuffSeq     '일련번호
        vCustom(13) = Trim(RsHeader!WorkName)
        vCustom(14) = RsHeader!StuffDate
        vCustom(15) = RsHeader!StuffWidth
        vCustom(16) = ChkNullValue(RsHeader!UnitClss)
        vCustom(17) = MakeOrderID(RsHeader!OrderID, OM_EXPAND)
        vCustom(18) = Trim(RsHeader!Remark)
    End If
    
    nPage = 1
    nLineCnt = 21
    nRow = 1
    Call PrintHeader(nPage, vCustom)
    With oFlex
        If .Rows > .FixedRows Then
            Call PrintDot(17, yPos(nRow - 1), "I/G 반품")
            Call PrintDot(50, yPos(nRow - 1), MakeStrBySpace(vCustom(10), 3, 0))
            Call PrintDot(58, yPos(nRow - 1), MakeStrBySpace(vCustom(11) & "Y", 9, 0))
            
            For II = .FixedRows To .Rows - 1
            
                If nRow > nLineCnt Then
                    nPage = nPage + 1
                    Printer.NewPage
                    Call PrintHeader(nPage, vCustom)
                    nRow = 1
                End If
            
                For JJ = 2 To 11
                    xPos = 77 + (JJ - 2) * 10
                    Call PrintDot(xPos, yPos(nRow - 1), Trim(.TextMatrix(II, JJ)))
                Next JJ
                nRow = nRow + 1
            Next II
        End If
    End With
    
    If nRow < nLineCnt Then
        Call PrintDot(17, yPos(nRow - 1), "** 이하여백 **")
        Printer.NewPage
        
'        Printer.Line (8, yPos(nRow - 1) + 4)-(174, yPos(nRow - 1) + 4)
        
    End If
    
    Printer.EndDoc
'    Printer.KillDoc
End Sub

Private Function PrintDot(nXPos As Integer, nYPos As Integer, sStr As String, Optional nFont As Integer = 10)
    With Printer
        .CurrentX = nXPos
        .CurrentY = nYPos
        .Font.Size = nFont
    End With
    Printer.Print sStr
    
End Function

Private Function PrintHeader(nPage As Integer, vCustom() As String)
    
    Call PrintDot(160, 17, "PAGE : " & nPage)
    Call PrintDot(125, 23, "담당")
    Call PrintDot(138, 23, "과장")
    Call PrintDot(152, 23, "이사")
    Call PrintDot(165, 23, "사장")

'    Call PrintDot(23, 32, Left(m_sTranNo, 4) & "-" & Right(m_sTranNo, 2) & "-" & m_nTranSeq) '일련번호
    Call PrintDot(23, 32, vCustom(12)) '일련번호
    Call PrintDot(68, 32, Left(vCustom(14), 4))
    Call PrintDot(85, 32, Mid(vCustom(14), 5, 2))
    Call PrintDot(97, 32, Right(vCustom(14), 2))
    
    Call PrintDot(114, 39, vCustom(1))  '거래처
    Call PrintDot(114, 45, vCustom(9))  '출고처
    Call PrintDot(114, 51, vCustom(18))  '비고사항
    Call PrintDot(114, 58, vCustom(7)) 'Order No
    
    
    Call PrintDot(8, 73, vCustom(17))  '관리번호
    Call PrintDot(43, 73, vCustom(6))  '품명
    Call PrintDot(78, 73, vCustom(15))  '규격
    Call PrintDot(92, 73, vCustom(13))   '가공구분
    Call PrintDot(115, 73, vCustom(8)) '오더량
    Call PrintDot(138, 73, vCustom(16))  '단위
    
    Call PrintDot(150, 73, vCustom(10))    '절수
    Call PrintDot(163, 73, vCustom(11))    '출고량
        
End Function
''
''If oStuffIn.GetStuffINReturnGoods(MakeDate(DF_SHORT, dtpDate(2)), StuffClss, val(txtStuffSeq), RsHeader) Then
''''                Call SetPrint(RsHeader, grdStuffINSub)
''''                Set oStuffIn = Nothing
''''            End If

'S_201105_태을염직_03 에 의한 추가
Private Sub MakeExcelPacking()
    Dim StuffClss As String
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim RsHeader As ADODB.Recordset ', RsDetail As ADODB.Recordset
    Dim rsData As ADODB.Recordset

    Dim oCustom                         As PlusLib2.CCustom
    Dim rs                              As ADODB.Recordset
    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oFs                             As FileSystemObject
    Dim i%, j%, nRow%, nLimitRow%, nBaseRow%, nCurRow%, nPage%, sReport$
    Dim nOrderSeq%, sLotNo$
    Dim sUnit$, nColorRoll%, nColorQty#
    Dim vCustom(19)                      As String
    
    Dim sColor As String
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
   
    '*****************************************************************
    ' 공급받는자 정보 Get
    '------------------------------------------------------------------
    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    StuffClss = CboStuffClss.ItemData(CboStuffClss.ListIndex)
    If oStuffIn.GetStuffINReturnGoods(MakeDate(DF_SHORT, dtpDate(2)), StuffClss, val(txtStuffSeq), RsHeader) Then
        '//Call SetPrint(RsHeader, grdStuffINSub)
        Set oStuffIn = Nothing
    End If

    If RsHeader.RecordCount > 0 Then
        '거래처 사업자 번호
        vCustom(0) = Left(CheckNull(RsHeader!CustomNo), 3) & " - " & Mid(CheckNull(RsHeader!CustomNo), 4, 2) & " - " & Right(CheckNull(RsHeader!CustomNo), 5)
        vCustom(1) = CheckNull(RsHeader!kCustom)        '거래처명
        vCustom(2) = CheckNull(RsHeader!Chief)          '거래처-대표자
        
''        'S_201312_태을염직_99 에 의한 수정-OLD 소스
''        vCustom(3) = CheckNull(RsHeader!Address1) & " " & CheckNull(RsHeader!Address2)  '거래처 주소
        
        'S_201312_태을염직_99 에 의한 수정-NEW 소스
        If CheckNull(RsHeader!Address1) <> "" Then                '도로명 주소 있으면
            vCustom(3) = CheckNull(RsHeader!Address1) & " " & CheckNull(RsHeader!Address2)  '거래처 주소
        Else                                                '도로명 주소 없으면-지번주소
            vCustom(3) = CheckNull(RsHeader!AddressJiBun1) & " " & CheckNull(RsHeader!AddressJiBun2)  '거래처 주소
        End If
        
        vCustom(4) = CheckNull(RsHeader!Condition)      '업태
        vCustom(5) = CheckNull(RsHeader!Category)       '업종
        vCustom(6) = CheckNull(RsHeader!Article)        '품명
        vCustom(7) = CheckNull(RsHeader!OrderNo)        'OrderNo
        vCustom(8) = SetCurrency(ChkNullValue(RsHeader!OrderQty), 0)    '오더수량
        vCustom(9) = CheckNull(RsHeader!Custom)                 '반품처명
        vCustom(10) = SetCurrency(RsHeader!TotRoll, 0)          '총 roll 수
        vCustom(11) = SetCurrency(RsHeader!TotQty, 0)           '촐 수량
        vCustom(12) = RsHeader!StuffDate & "-" & RsHeader!StuffClss & "-" & RsHeader!StuffSeq     '일련번호
        vCustom(13) = Trim(RsHeader!WorkName)       '가공구분
        vCustom(14) = RsHeader!StuffDate            '입고일자
        vCustom(15) = RsHeader!StuffWidth           '생지폭
        vCustom(16) = IIf(ChkNullValue(RsHeader!UnitClss) = "", "Y", RsHeader!UnitClss) '입고단위
        vCustom(17) = IIf(CheckNull(RsHeader!OrderID) = "", "", MakeOrderID(CheckNull(RsHeader!OrderID), OM_EXPAND))  '관리번호
        vCustom(18) = Trim(RsHeader!Remark)         '비고
    End If
    '*****************************************************************
    
     Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & REPORTFILE)

    oExcel.WindowState = xlMaximized
    oExcel.Application.Visible = True

    With oExcel
        ' Make Sum
        .Worksheets("Form").Activate
''        .Cells(4, 1) = MakeDate(DF_FULL, dtpOutDate.Value)
        '입고일자-입고구분-SEQ-d입고일자
        
        .Cells(4, 1) = "일련번호:" & vCustom(12) & Space(15) & MakeDate(DF_FULL, vCustom(14))
        '*****************************************************************
        ' 공급자 정보 출력-S_201312_태을염직_99 에 의한 추가-기존 하드 코딩에서 DB에서 가져옴
        '------------------------------------------------------------------
        .Cells(5, 4) = Format(g_companyInfo.Company_No, "###-##-#####")          '사업자번호
        .Cells(6, 4) = g_companyInfo.Company_Name                                  '회사명
        .Cells(6, 9) = g_companyInfo.Chief         '대표자
        If g_companyInfo.Address1 <> "" Then                '도로명 주소 있으면
            .Cells(7, 4) = g_companyInfo.Address1 & " " & g_companyInfo.Address2
        Else                                                '도로명 주소 없으면-지번주소
            .Cells(7, 4) = g_companyInfo.AddressJiBun1 & " " & g_companyInfo.AddressJiBun2
        End If
        .Cells(8, 4) = g_companyInfo.Company_type        '업태
        .Cells(8, 9) = g_companyInfo.Category        '종목

        .Cells(38, 19) = "Tel. " & g_companyInfo.Phone         '전화번호
        .Cells(39, 19) = "Fax. " & g_companyInfo.FaxNO         '팩스번호
        .Cells(38, 25) = g_companyInfo.Company_Name            '회사명
        '*****************************************************************
        
        '*****************************************************************
        ' 공급받는자 정보 출력
        '------------------------------------------------------------------
        .Cells(5, 18) = vCustom(0)         '사업자번호
        .Cells(6, 18) = vCustom(1)         '회사명
        .Cells(6, 26) = vCustom(2)         '대표
        .Cells(7, 18) = vCustom(3)         '주소
        .Cells(8, 18) = vCustom(4)         '업태
        .Cells(8, 26) = vCustom(5)         '종목
        '*****************************************************************

        Dim nChunRateQty As Long
        
        '2011.05.19 김대진 대리 요청- ORderNo대신 관리번호
        .Cells(9, 4) = vCustom(17)                                                                      '관리번호
        .Cells(9, 13) = IIf(vCustom(8) > 0, Format(vCustom(8), "#,###") & " " & vCustom(16), "")        'Order 량

         
        '.Cells(9, 17) = txtChunkRate.Text + " %"                            '축율
        .Cells(9, 21) = vCustom(9)                            '반품처

        .Cells(12, 1) = vCustom(6)                              '품명
        .Cells(12, 7) = vCustom(15)        '규격
        .Cells(12, 11) = vCustom(13)        '가공구분

        .Cells(12, 14) = Format(vCustom(10), "#,###")           '가공마수-절 ***
        If vCustom(16) = "Y" Then
            .Cells(12, 17) = Format(vCustom(11), "#,###") & "Y" '가공마수길이
            
        Else
             .Cells(12, 17) = Format(vCustom(11), "#,###") & "M" & vbLf & _
                              "(" & Format(CLng((vCustom(11) / 0.9144)), "#,###") & "Y)"  ' 가공마수길이-반품입고량
        End If

        .Cells(10, 22) = vCustom(7)                                'OrderNo
        .Cells(12, 22) = vCustom(18)                              '비고 (OLD:12,20)
        
        .Worksheets("Print").Activate
        
        nPage = 1
        nBaseRow = GetExcelRollBaseRow(nPage)
        Call InsertExcelForm(oExcel, nPage)
        nCurRow = nBaseRow + 15
    
        For i = grdStuffINSub.FixedRows To grdStuffINSub.Rows - 1
            If nCurRow + nRow > nBaseRow + 37 Then
                nPage = nPage + 1
                nBaseRow = GetExcelRollBaseRow(nPage)
                Call InsertExcelForm(oExcel, nPage)
                nCurRow = nBaseRow + 15
                nRow = 0
            End If
            
            sColor = "I/G 반품"

            If i = grdStuffINSub.FixedRows Then
                .Cells(nCurRow + nRow, 3) = sColor          '머지 효과를 위해 첫번쨰 값에서만 출력
            End If
            .Cells(nCurRow + nRow, 6) = Format(grdStuffINSub.TextMatrix(i, 12), "#,###")            'PCS(절수)
            .Cells(nCurRow + nRow, 8) = Format(grdStuffINSub.TextMatrix(i, 13), "#,###")            '수량
            
            '각 칼러별 합계 누적
            nColorRoll = nColorRoll + CheckNum(grdStuffINSub.TextMatrix(i, 12))                     'PCS(절수)
            nColorQty = nColorQty + CheckNum(grdStuffINSub.TextMatrix(i, 13))                       '수량

           .Cells(nCurRow + nRow, 11) = grdStuffINSub.TextMatrix(i, 2)
           .Cells(nCurRow + nRow, 13) = grdStuffINSub.TextMatrix(i, 3)
           .Cells(nCurRow + nRow, 14) = grdStuffINSub.TextMatrix(i, 4)
           .Cells(nCurRow + nRow, 16) = grdStuffINSub.TextMatrix(i, 5)
           .Cells(nCurRow + nRow, 19) = grdStuffINSub.TextMatrix(i, 6)
           .Cells(nCurRow + nRow, 20) = grdStuffINSub.TextMatrix(i, 7)
           .Cells(nCurRow + nRow, 22) = grdStuffINSub.TextMatrix(i, 8)
           .Cells(nCurRow + nRow, 25) = grdStuffINSub.TextMatrix(i, 9)
           .Cells(nCurRow + nRow, 27) = grdStuffINSub.TextMatrix(i, 10)
           .Cells(nCurRow + nRow, 29) = grdStuffINSub.TextMatrix(i, 11)

            nRow = nRow + 1
        Next i
        
        If nCurRow + nRow > nBaseRow + 37 Then
            nPage = nPage + 1
            nBaseRow = GetExcelRollBaseRow(nPage)
            Call InsertExcelForm(oExcel, nPage)
            nCurRow = nBaseRow + 15
        End If
        .Cells(nCurRow + nRow, 3) = sColor & " 계 : "
        .Cells(nCurRow + nRow, 6) = Format(nColorRoll, "#,###")
        .Cells(nCurRow + nRow, 8) = Format(nColorQty, "#,###")
        
    End With

    sReport = App.Path & REPORTFILE1

    Set oFs = New FileSystemObject
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

    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing

    Call Err.Raise(Err.Number, "frmOutware.MakeExcelPacking", Err.Description)
    
 
End Sub

'S_201105_태을염직_03 에 의한 추가
Private Function GetExcelRollBaseRow(nPage)
    GetExcelRollBaseRow = (nPage - 1) * EXCEL_ROLL_ROW
End Function

'S_201105_태을염직_03 에 의한 추가
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
        
        .Cells(nBaseRow + 3, 27) = "PAGE : " & nPage
    End With
End Function



