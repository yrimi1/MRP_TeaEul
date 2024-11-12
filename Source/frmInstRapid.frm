VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInstRapid 
   ClientHeight    =   10050
   ClientLeft      =   -510
   ClientTop       =   1920
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   15255
   Begin VB.ComboBox cboColor 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8445
      Style           =   2  '드롭다운 목록
      TabIndex        =   68
      Top             =   90
      Width           =   1725
   End
   Begin VB.Frame fraWorkEnd 
      BorderStyle     =   0  '없음
      Height          =   4365
      Left            =   4920
      TabIndex        =   40
      Top             =   4110
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ComboBox cboHold 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5820
         TabIndex        =   114
         Text            =   "Combo1"
         Top             =   3210
         Width           =   2145
      End
      Begin VB.CommandButton cmdInvisible 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7620
         TabIndex        =   65
         Top             =   120
         Width           =   315
      End
      Begin VB.TextBox txtRoll 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   1320
         TabIndex        =   50
         Top             =   3630
         Width           =   1755
      End
      Begin VB.TextBox txtRemarkResult 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   49
         Top             =   3210
         Width           =   3315
      End
      Begin VB.CommandButton cmdEndCancel 
         Caption         =   "작성 취소"
         Height          =   615
         Left            =   5880
         TabIndex        =   48
         Top             =   3630
         Width           =   1005
      End
      Begin VB.CommandButton cmdEndConfirm 
         Caption         =   "일지 작성"
         Height          =   615
         Left            =   6930
         TabIndex        =   47
         Top             =   3630
         Width           =   1005
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Tag             =   "염색패턴"
         Top             =   780
         Width           =   2955
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   7
         Left            =   3090
         TabIndex        =   45
         Tag             =   "염색구분"
         Top             =   780
         Width           =   1545
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   8
         Left            =   4650
         TabIndex        =   44
         Tag             =   "염색구분"
         Top             =   780
         Width           =   1365
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   9
         Left            =   6030
         TabIndex        =   43
         Tag             =   "작업자"
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   1320
         TabIndex        =   42
         Top             =   3960
         Width           =   1755
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   10
         Left            =   7140
         TabIndex        =   41
         Tag             =   "작업자"
         Top             =   780
         Width           =   795
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   3630
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "절수"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   2
         Left            =   3090
         TabIndex        =   52
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작업구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   3960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "수량"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   5
         Left            =   6030
         TabIndex        =   54
         Top             =   450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작업자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSMask.MaskEdBox txtEndDate 
         Height          =   315
         Left            =   4290
         TabIndex        =   55
         Top             =   3630
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   6
         Left            =   3120
         TabIndex        =   56
         Top             =   3630
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "종료 일자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   7
         Left            =   3120
         TabIndex        =   57
         Top             =   3960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "종료 시간"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   8
         Left            =   4650
         TabIndex        =   58
         Top             =   450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "염색구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   59
         Top             =   450
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "염색패턴"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   60
         Top             =   3210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "비고사항"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   11
         Left            =   7140
         TabIndex        =   61
         Top             =   450
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작업 조"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   503
         _Version        =   196609
         BackColor       =   16761024
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "염색 작업 일지 작성"
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin MSMask.MaskEdBox txtEndTime 
         Height          =   315
         Left            =   4290
         TabIndex        =   63
         Top             =   3960
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "## : ##"
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   315
         Index           =   6
         Left            =   4650
         TabIndex        =   112
         Top             =   3210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작 업 조"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkHold 
            Caption         =   "보류처리"
            Height          =   180
            Left            =   105
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   60
            Width           =   1020
         End
      End
      Begin VB.Label lblDyeClss 
         Height          =   225
         Left            =   270
         TabIndex        =   67
         Top             =   3480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblSchSeq 
         AutoSize        =   -1  'True
         Caption         =   "00000000101"
         Height          =   180
         Left            =   2850
         TabIndex        =   64
         Top             =   2970
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   7920
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   4305
         Left            =   30
         Top             =   30
         Width           =   7995
      End
   End
   Begin TabDlg.SSTab tabRapid 
      Height          =   8235
      Left            =   30
      TabIndex        =   9
      Top             =   525
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   14526
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabMaxWidth     =   4763
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1호기"
      TabPicture(0)   =   "frmInstRapid.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pnlTab(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "grdTab(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdHide"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraPosition"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmInstRapid.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(1)"
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).Control(2)=   "grdTab(1)"
      Tab(1).Control(3)=   "pnlTab(1)"
      Tab(1).ControlCount=   4
      Begin VB.Frame fraPosition 
         Height          =   405
         Left            =   5250
         TabIndex        =   74
         Top             =   -60
         Width           =   9975
         Begin VB.OptionButton optMachine 
            Caption         =   "12호"
            Height          =   225
            Index           =   11
            Left            =   9255
            TabIndex        =   92
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "11호"
            Height          =   225
            Index           =   10
            Left            =   8550
            TabIndex        =   85
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "10호"
            Height          =   225
            Index           =   9
            Left            =   7848
            TabIndex        =   84
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "09호"
            Height          =   225
            Index           =   8
            Left            =   7146
            TabIndex        =   83
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "08호"
            Height          =   225
            Index           =   7
            Left            =   6444
            TabIndex        =   82
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "07호"
            Height          =   225
            Index           =   6
            Left            =   5742
            TabIndex        =   81
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "06호"
            Height          =   225
            Index           =   5
            Left            =   5040
            TabIndex        =   80
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "05호"
            Height          =   225
            Index           =   4
            Left            =   4338
            TabIndex        =   79
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "04호"
            Height          =   225
            Index           =   3
            Left            =   3636
            TabIndex        =   78
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "03호"
            Height          =   225
            Index           =   2
            Left            =   2934
            TabIndex        =   77
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "02호"
            Height          =   225
            Index           =   1
            Left            =   2232
            TabIndex        =   76
            Top             =   150
            Width           =   675
         End
         Begin VB.OptionButton optMachine 
            Caption         =   "01호"
            Height          =   225
            Index           =   0
            Left            =   1530
            TabIndex        =   75
            Top             =   150
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.Label Label3 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "염색호기로 이동"
            Height          =   180
            Left            =   105
            TabIndex        =   86
            Top             =   180
            Width           =   1335
         End
      End
      Begin Threed.SSCommand cmdHide 
         Height          =   315
         Left            =   2790
         TabIndex        =   73
         Top             =   30
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "실적 감추기"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdTab 
         Height          =   7830
         Index           =   0
         Left            =   30
         TabIndex        =   10
         Top             =   375
         Width           =   15150
         _cx             =   26723
         _cy             =   13811
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   20
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
         Begin VB.Shape shpBox 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Height          =   690
            Left            =   1770
            Shape           =   4  '둥근 사각형
            Top             =   3180
            Width           =   2900
         End
      End
      Begin Threed.SSPanel pnlTab 
         Height          =   345
         Index           =   0
         Left            =   45
         TabIndex        =   11
         Top             =   15
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   0
         BackColor       =   16761024
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가나다라"
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlTab 
         Height          =   285
         Index           =   1
         Left            =   -72180
         TabIndex        =   12
         Top             =   60
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   5
         BackColor       =   12259610
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdTab 
         Height          =   7800
         Index           =   1
         Left            =   -74970
         TabIndex        =   28
         Top             =   390
         Width           =   15150
         _cx             =   26723
         _cy             =   13758
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483634
         GridColorFixed  =   -2147483639
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   20
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
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  '단일 고정
         Height          =   225
         Index           =   1
         Left            =   -71040
         TabIndex        =   35
         Top             =   150
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "작업중"
         Height          =   180
         Index           =   1
         Left            =   -70500
         TabIndex        =   34
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "작업중"
         Height          =   180
         Index           =   0
         Left            =   4650
         TabIndex        =   30
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  '단일 고정
         Height          =   225
         Index           =   0
         Left            =   4320
         TabIndex        =   29
         Top             =   90
         Width           =   315
      End
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   495
      Left            =   75
      TabIndex        =   26
      Top             =   0
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   65535
      MarqueeDelay    =   700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "조회 중 입니다...."
      BorderWidth     =   2
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin VB.Frame fraButton 
      BorderStyle     =   0  '없음
      Height          =   615
      Left            =   12390
      TabIndex        =   25
      Top             =   -120
      Visible         =   0   'False
      Width           =   2895
      Begin Threed.SSCommand cmdConfirm 
         Height          =   495
         Left            =   975
         TabIndex        =   108
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "카드선택"
      End
      Begin Threed.SSCommand cmdScreen 
         Height          =   495
         Left            =   1920
         TabIndex        =   109
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "편집취소"
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   495
         Left            =   30
         TabIndex        =   110
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "삭제"
      End
   End
   Begin Threed.SSCommand cmdToggle 
      Height          =   495
      Left            =   2760
      TabIndex        =   24
      Tag             =   "PERM_ADDNEW"
      Top             =   0
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   873
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "염색 스케쥴 작성"
   End
   Begin Threed.SSPanel pnlView 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   5054
      _Version        =   196609
      Font3D          =   5
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin TabDlg.SSTab SSTab1 
         Height          =   2805
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   4948
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   617
         TabMaxWidth     =   4410
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "염색공정 대기"
         TabPicture(0)   =   "frmInstRapid.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "pnlWaitTab(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "grdList(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdRefresh"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cboCustom"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cboArticle"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "염색직전 공정 대기"
         TabPicture(1)   =   "frmInstRapid.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "pnlWaitTab(1)"
         Tab(1).Control(1)=   "grdList(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "그 외 이전 공정 대기"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "pnlWaitTab(2)"
         Tab(2).Control(1)=   "grdList(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "염색완료 카드내역"
         TabPicture(3)   =   "frmInstRapid.frx":0070
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label4"
         Tab(3).Control(1)=   "grdList(3)"
         Tab(3).Control(2)=   "pnlWaitTab(3)"
         Tab(3).ControlCount=   3
         Begin VB.ComboBox cboArticle 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11940
            Style           =   2  '드롭다운 목록
            TabIndex        =   118
            Top             =   15
            Width           =   2280
         End
         Begin VB.ComboBox cboCustom 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10215
            Style           =   2  '드롭다운 목록
            TabIndex        =   117
            Top             =   15
            Width           =   1710
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   315
            Left            =   14235
            TabIndex        =   116
            Top             =   15
            Width           =   945
         End
         Begin VSFlex7LCtl.VSFlexGrid grdList 
            Height          =   2370
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   15090
            _cx             =   26617
            _cy             =   4180
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
            Cols            =   30
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
         Begin VSFlex7LCtl.VSFlexGrid grdList 
            Height          =   2370
            Index           =   1
            Left            =   -74940
            TabIndex        =   15
            Top             =   383
            Width           =   15090
            _cx             =   26617
            _cy             =   4180
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
            Cols            =   30
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
         Begin VSFlex7LCtl.VSFlexGrid grdList 
            Height          =   2370
            Index           =   2
            Left            =   -74940
            TabIndex        =   16
            Top             =   383
            Width           =   15090
            _cx             =   26617
            _cy             =   4180
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
            Cols            =   30
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
         Begin Threed.SSPanel pnlWaitTab 
            Height          =   345
            Index           =   0
            Left            =   75
            TabIndex        =   17
            Top             =   30
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   5
            ForeColor       =   16777215
            BackColor       =   12259610
            PictureMaskColor=   16777215
            MarqueeDelay    =   700
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "염색공정 대기"
            BevelOuter      =   0
            FloodColor      =   12259610
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlWaitTab 
            Height          =   345
            Index           =   1
            Left            =   -72405
            TabIndex        =   18
            Top             =   30
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   5
            ForeColor       =   16777215
            BackColor       =   12539970
            PictureMaskColor=   16777215
            MarqueeDelay    =   700
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "염색직전 공정 대기"
            BevelOuter      =   0
            FloodColor      =   12539970
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlWaitTab 
            Height          =   345
            Index           =   2
            Left            =   -69870
            TabIndex        =   19
            Top             =   30
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   5
            ForeColor       =   16777215
            BackColor       =   14389120
            PictureMaskColor=   16777215
            MarqueeDelay    =   700
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "그 외 이전 공정 대기"
            BevelOuter      =   0
            FloodColor      =   14389120
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlWaitTab 
            Height          =   345
            Index           =   3
            Left            =   -67335
            TabIndex        =   32
            Top             =   30
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   5
            ForeColor       =   65535
            BackColor       =   15715015
            PictureMaskColor=   16777215
            MarqueeDelay    =   700
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "염색완료 카드내역"
            BevelOuter      =   0
            FloodColor      =   15715015
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid grdList 
            Height          =   2370
            Index           =   3
            Left            =   -74940
            TabIndex        =   33
            Top             =   390
            Width           =   15090
            _cx             =   26617
            _cy             =   4180
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
            Cols            =   30
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "■  수정/추가에만 선택하십시요"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   -62730
            TabIndex        =   66
            Top             =   150
            Width           =   2580
         End
      End
   End
   Begin VB.Frame fraFunc 
      Height          =   675
      Left            =   30
      TabIndex        =   27
      Top             =   8640
      Width           =   15225
      Begin VB.Frame fraUpDown 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   555
         Left            =   7200
         TabIndex        =   87
         Top             =   120
         Width           =   2925
         Begin Threed.SSCommand cmdUP 
            Height          =   435
            Left            =   90
            TabIndex        =   88
            Top             =   60
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   767
            _Version        =   196609
            Alignment       =   8
         End
         Begin Threed.SSCommand cmdDown 
            Height          =   435
            Left            =   630
            TabIndex        =   89
            Top             =   60
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   767
            _Version        =   196609
            Alignment       =   8
         End
         Begin Threed.SSCommand cmdSequence 
            Height          =   435
            Left            =   1200
            TabIndex        =   90
            Top             =   60
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   767
            _Version        =   196609
            Caption         =   "스케쥴적용"
         End
         Begin Threed.SSCommand cmdCancelSeq 
            Height          =   435
            Left            =   2280
            TabIndex        =   91
            Top             =   60
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   767
            _Version        =   196609
            Caption         =   "취소"
         End
         Begin VB.Shape shpUpDown 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Height          =   495
            Left            =   30
            Top             =   30
            Width           =   2880
         End
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   13920
         TabIndex        =   39
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "      닫기(&X)"
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdWorkStart 
         Height          =   435
         Left            =   10260
         TabIndex        =   93
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
         _Version        =   196609
         Caption         =   "작업시작"
      End
      Begin Threed.SSCommand cmdWorkEnd 
         Height          =   435
         Left            =   11130
         TabIndex        =   94
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
         _Version        =   196609
         Caption         =   "작업완료"
      End
      Begin Threed.SSCommand cmdCancelStart 
         Height          =   435
         Left            =   12060
         TabIndex        =   95
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
         _Version        =   196609
         Caption         =   "작업취소"
      End
      Begin Threed.SSCommand cmdAddWork 
         Height          =   435
         Left            =   12930
         TabIndex        =   96
         Top             =   180
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "추가작업"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   0
         Left            =   30
         TabIndex        =   97
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "카드분리"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   6
         Left            =   6225
         TabIndex        =   98
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         ForeColor       =   0
         Caption         =   "일지조회"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   5
         Left            =   5340
         TabIndex        =   99
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "카드상세"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   4
         Left            =   4455
         TabIndex        =   100
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "수주상세"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   3
         Left            =   3570
         TabIndex        =   101
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "평량지시"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   2
         Left            =   2685
         TabIndex        =   102
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "처방조회"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   1
         Left            =   915
         TabIndex        =   103
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "카드변경"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   8
         Left            =   1800
         TabIndex        =   104
         Top             =   120
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "패턴변경"
      End
      Begin Threed.SSCommand cmdButton 
         Height          =   525
         Index           =   7
         Left            =   4200
         TabIndex        =   105
         Top             =   120
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "염색패턴"
      End
      Begin VB.Shape shpButton 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   495
         Left            =   10230
         Top             =   150
         Width           =   3570
      End
   End
   Begin Threed.SSPanel pnlCardID 
      Height          =   315
      Left            =   6810
      TabIndex        =   69
      Top             =   90
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "카드번호"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlSplitID 
      Height          =   315
      Left            =   7845
      TabIndex        =   70
      Top             =   90
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "분할"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdButton 
      Height          =   495
      Index           =   9
      Left            =   10170
      TabIndex        =   106
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "색상변경"
   End
   Begin Threed.SSCommand cmdRefesh 
      Height          =   495
      Left            =   11310
      TabIndex        =   107
      Top             =   0
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      _Version        =   196609
      Caption         =   "새로고침"
   End
   Begin Threed.SSCommand cmdExpandWait 
      Height          =   495
      Left            =   4950
      TabIndex        =   111
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      _Version        =   196609
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "공정대기 확장"
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   2865
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   5054
      _Version        =   196609
      Font3D          =   5
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   3
         Left            =   14280
         TabIndex        =   23
         Tag             =   "작업자"
         Top             =   420
         Width           =   915
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   4
         Left            =   14190
         TabIndex        =   71
         Tag             =   "작업자"
         Top             =   420
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   5
         Left            =   11700
         TabIndex        =   36
         Tag             =   "염색구분"
         Top             =   420
         Width           =   1305
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   2
         Left            =   13020
         TabIndex        =   22
         Tag             =   "염색구분"
         Top             =   420
         Width           =   1245
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   1
         Left            =   9720
         TabIndex        =   21
         Tag             =   "염색패턴"
         Top             =   420
         Width           =   1965
      End
      Begin VB.ListBox lstArray 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   0
         Left            =   8880
         TabIndex        =   20
         Tag             =   "염색호기"
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox txtRemark 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2520
         Width           =   7485
      End
      Begin VSFlex7LCtl.VSFlexGrid grdList 
         Height          =   2490
         Index           =   4
         Left            =   30
         TabIndex        =   2
         Top             =   15
         Width           =   8850
         _cx             =   15610
         _cy             =   4392
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
         Cols            =   30
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
         Height          =   360
         Index           =   2
         Left            =   8880
         TabIndex        =   3
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "염색호기"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   360
         Index           =   3
         Left            =   9720
         TabIndex        =   4
         Top             =   30
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "염색작업 패턴"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   360
         Index           =   4
         Left            =   13020
         TabIndex        =   5
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "염색구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   360
         Index           =   5
         Left            =   14280
         TabIndex        =   6
         Top             =   30
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "작업자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   7
         Left            =   30
         TabIndex        =   7
         Top             =   2520
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "비고사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   360
         Index           =   0
         Left            =   11700
         TabIndex        =   37
         Top             =   30
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "작업구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Label lblMachID 
      Caption         =   "Label5"
      Height          =   285
      Left            =   6690
      TabIndex        =   115
      Top             =   270
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblOrderID 
      Caption         =   "Label5"
      Height          =   285
      Left            =   10350
      TabIndex        =   72
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblWork 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   7680
      TabIndex        =   38
      Top             =   150
      Width           =   60
   End
   Begin VB.Label lblSchIDSeq 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "00000000101"
      Height          =   180
      Left            =   3810
      TabIndex        =   31
      Top             =   180
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmInstRapid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bEnableWork As Boolean
Private Const Custom = "삼우"   ' Rapid: 11(1~11), CPB: 1(12)
Private Const nMachNo = 12      ' 12호기 추가
Private nMachineID As Integer



Private Sub FoldOrUnFold(bFlag As Boolean)
' bFlag = True(확장), False(축소)
    If bFlag Then
        cmdExpandWait.Caption = "공정대기 축소"
        pnlView.Height = 8205
        SSTab1.Height = 8175
        grdList(0).Height = 7740
        grdList(1).Height = 7740
        grdList(2).Height = 7740
        grdList(3).Height = 7740
        tabRapid.Visible = False
    Else
        cmdExpandWait.Caption = "공정대기 확장"
        tabRapid.Visible = True
        pnlView.Height = 2865
        SSTab1.Height = 2805
        grdList(0).Height = 2370
        grdList(1).Height = 2370
        grdList(2).Height = 2370
        grdList(3).Height = 2370
    End If
End Sub

Private Sub chkHold_Click()
    If chkHold.Value = 1 Then
        cboHold.Enabled = True
    Else
        cboHold.Enabled = False
    End If
End Sub

Private Sub cmdAddWork_Click()
Dim oRapid As PlusLib2.CRapid
Dim nSchID As Long
Dim nSeq As Integer
    
    If MsgBox("정말로 추가작업을 하시겠습니까?", vbYesNo + vbQuestion, "추가작업") = vbYes Then
        Set oRapid = New PlusLib2.CRapid
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
        
        With grdTab(tabRapid.Tab)
            nSchID = CLng(Left(.TextMatrix(.Row, 20), 9))
            nSeq = CInt(Right(.TextMatrix(.Row, 20), 2))
        End With
        
        Call ToggleShapeBox(False, False)
        
        If oRapid.AddDyeWorkRapid(nSchID, nSeq, Format(Now, "YYYYMMDD"), Format(time, "HHMM")) Then
            Set oRapid = Nothing
            MsgBox "염색 추가작업이 정상적으로 처리되었습니다" & vbCrLf & _
                   "평량지시를 내린후 작업을 진행시켜야 합니다", vbInformation, "추가작업"
            
            Call InitGrid
            Call InitGrdTab
            Call FillGridData
            Call FillSchData
        Else
            Set oRapid = Nothing
        End If
    End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
    Dim oRapid As PlusLib2.CRapid

    Select Case Index
        Case 0: '카드분리
            Dim sRs As Recordset
            
            If Trim(pnlCardID) <> "" Then
                Set oRapid = New PlusLib2.CRapid
                oRapid.Connection = g_adoCon
                oRapid.UserName = g_sUserName
                
                Set sRs = oRapid.GetCheckDyeSch(Trim(pnlCardID), Trim(pnlSplitID))
                Set oRapid = Nothing
                
                If sRs.RecordCount > 0 Then
                    If Trim(sRs!Complitclss) = "" Then
                        MsgBox "염색작업지시가 내려진 카드는 카드분리를 할수 없습니다", vbInformation, "카드분리 불가"
                        Exit Sub
                    End If
                End If
                Set sRs = Nothing
                frmCardDivide.chkSearch(4).Value = vbChecked
                frmCardDivide.txtSearch(4).Text = Select_TabRow_No("카드번호")
                Call frmCardDivide.cmdSearch_Click
            End If
        Case 1: '색상변경
            frmCardChange.chkSearch(4).Value = vbChecked
            frmCardChange.txtSearch(4).Text = Select_TabRow_No("카드번호")
            Call frmCardChange.cmdSearch_Click
        Case 2: '처방조회
            frmRecipeView.optOrder(1).Value = True
            frmRecipeView.chkSearch(3).Value = vbUnchecked
            frmRecipeView.chkSearch(2).Value = vbChecked
            frmRecipeView.tabMain.Tab = 0
            If shpBox.Visible = True Then   ' 스케쥴에 근거한 관리번호
                frmRecipeView.txtSearch(2).Text = Select_TabRow_No("관리번호")
            Else
                If grdTab(0).TextMatrix(grdTab(0).Row, 1) = "실적" Then
                    frmRecipeView.txtSearch(2).Text = Select_TabRow_No("관리번호")
                Else            ' 카드에 근거한 관리번호
                    frmRecipeView.txtSearch(2).Text = lblOrderID
                End If
            End If
            Call frmRecipeView.FillGridRecipe
            
        Case 3: '평량지시
            Dim sSchIDSeq As String
            Dim rs As Recordset
            
            If shpBox.Visible = False Then
                MsgBox "염색지시건을 선택해야 합니다", vbInformation, "선택 요구"
                Exit Sub
            End If
            sSchIDSeq = Select_TabRow_No("스케쥴")
            
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
            
            Set rs = oRapid.GetCheckDyeWorking(CLng(Left(sSchIDSeq, 9)), CInt(Right(sSchIDSeq, 2)))
            Set oRapid = Nothing
            
            If rs.RecordCount > 0 Then
                If (Trim(rs!UseClss) = "작업" Or Len(Trim(rs!UseClss)) = 8) And Left(rs!procid, 2) = "43" Then
                    Set rs = Nothing
                    MsgBox "선택되어진 건은 현재 작업중입니다" & vbCrLf & vbCrLf & "평량지시를 내릴수 없습니다", vbCritical, "편집 불가"
                    Exit Sub
                End If
            End If
            Set rs = Nothing
            Call frmRecipeCalc.SetInstruction(CLng(Left(sSchIDSeq, 9)), CInt(Right(sSchIDSeq, 2)))
        Case 4: '수주상세
            frmOrderHistory.optOrder(0).Value = True
            
            If shpBox.Visible = True Then   ' 스케쥴에 근거한 관리번호
                frmOrderHistory.txtSearch.Text = Select_TabRow_No("관리번호")
            Else                            ' 카드에 근거한 관리번호
                frmOrderHistory.txtSearch.Text = lblOrderID
            End If
            
            frmOrderHistory.txtSearch_KeyPress (vbKeyReturn)
        Case 5: '카드상세
            frmCardHistory.txtCard.Text = Select_TabRow_No("카드번호")
            frmCardHistory.txtCard_KeyPress (vbKeyReturn)
        Case 6: '염색일지 조회
            frmDyeResultView.dtpDate(0) = Now:   frmDyeResultView.dtpDate(1) = Now
            Call frmDyeResultView.cmdSearch_Click
        Case 7: '염색패턴
            frmDyePattern.Show 1
        Case 8: '패턴변경
            If Trim(pnlCardID) <> "" And pnlCardID <> "카드번호" Then
                frmCardPattern.chkSearch(4).Value = vbChecked
                frmCardPattern.txtSearch(4).Text = pnlCardID
                frmCardPattern.txtSearch(5).Text = pnlSplitID
                frmCardPattern.cmdSearch_Click
            Else
                MsgBox "카드 선택하고 버튼을 눌러주십시요", vbInformation, "카드 선택 요망"
                Exit Sub
            End If
        Case 9: '색상변경
            If pnlCardID = "카드번호" Or Trim(pnlCardID) = "" Then
                MsgBox "카드를 선택해야 합니다", vbInformation, "카드선택 요망"
                Exit Sub
            End If
            If cboColor.ListIndex < 0 Then
                MsgBox "색상을 선택해야합니다", vbInformation, "색상선택 요망"
                Exit Sub
            End If
            
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
        
            If oRapid.UpdateCardColor(pnlCardID, pnlSplitID, cboColor.ItemData(cboColor.ListIndex), g_sUserName) Then
                MsgBox "카드의 칼라를 변경했습니다", vbOKOnly, "칼라 변경"
            End If
            Set oRapid = Nothing
            Call FillGridData
            Call FillSchData
    End Select
End Sub

Private Function Select_TabRow_No(pCheck As String, Optional sOrderID As String) As String
Dim iCol%

    If pCheck = "카드번호" Then
        Select_TabRow_No = pnlCardID
    ElseIf pCheck = "관리번호" Then
        If sOrderID <> "" Then ' 카드선택에 의한 관리번호
            Select_TabRow_No = sOrderID
        Else                ' 염색스케줄번호에 의한 관리번호
            With grdTab(tabRapid.Tab)
                Select_TabRow_No = .TextMatrix(.Row, 24)
            End With
        End If
    Else    ' 스케쥴번호(9) + 스케쥴차수(2)
        
        With grdTab(tabRapid.Tab)
            Select_TabRow_No = .TextMatrix(.Row, 20)
        End With
    End If
End Function

Private Sub cmdCancelSeq_Click()
    If MsgBox("스케쥴 변경을 정말로 취소하시겠습니까?" & vbCrLf & vbCrLf & "취소시 원상태로 복귀됩니다", vbQuestion + vbYesNo, "취소 여부") = vbYes Then
        Call ToggleShapeBox(False, False)
        bEnableWork = True
        
        Call InitGrdTab
        Call FillSchData
    End If
End Sub

Private Sub cmdCancelStart_Click()
    Dim oRapid As PlusLib2.CRapid
    Dim iCol As Integer
    Dim nSchID As Long
    Dim nSeq As Integer
    
    If MsgBox("작업시작중인 염색지시건을 취소하시겠습니까?", vbQuestion + vbYesNo, "취소 여부") = vbYes Then
    
        Set oRapid = New PlusLib2.CRapid
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
    
        Screen.MousePointer = vbHourglass
        
        With grdTab(tabRapid.Tab)
            nSchID = CLng(Left(.TextMatrix(.Row, 20), 9))
            nSeq = CInt(Right(.TextMatrix(.Row, 20), 2))
        End With
        
        If oRapid.DeletewkRapid(nSchID, nSeq) Then
            MsgBox "작업시작이 취소되었습니다", vbOKOnly, "취소 성공"
        End If
        Set oRapid = Nothing
        Call InitGrid
        Call InitGrdTab
        Call FillGridData
        Call FillSchData
        Call ToggleShapeBox(False, False)
        
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub cmdDelete_Click()
    Dim oRapid As PlusLib2.CRapid
    
    If MsgBox("선택된 염색지시건을 삭제하시겠습니까?" & vbCrLf & vbCrLf & _
                "삭제하게 되면 평량지시 내역도 삭제됩니다" & vbCrLf & vbCrLf & _
                "그래도 삭제하시겠습니까?", vbQuestion + vbYesNo, "삭제 여부") = vbYes Then
    
        Set oRapid = New PlusLib2.CRapid
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
    
        Screen.MousePointer = vbHourglass
        
        If oRapid.DeletewiRapid(CLng(Left(lblSchIDSeq, 9)), CInt(Right(lblSchIDSeq, 2))) Then
            MsgBox "해당 염색지시가 삭제되었습니다", vbOKOnly, "삭제 성공"
        End If
        Set oRapid = Nothing
        Call InitGrid
        Call InitGrdTab
        Call FillSchData
        Call cmdScreen_Click
        
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub cmdDown_Click()
Dim iCol%, iBaseRow%, iBaseCol%
Dim oRapid As PlusLib2.CRapid
Dim tRs As ADODB.Recordset
Dim sRs As ADODB.Recordset

    With grdTab(tabRapid.Tab)
        
        If .Row < .Rows - 1 Then
            bEnableWork = False
            
            Call VisibleWorkFrame(False)
        
            iBaseRow = .Row
            
            If .Cell(flexcpForeColor, iBaseRow + 1, 3) = vbBlue Or _
                .RowHeight(iBaseRow + 1) < 50 Or _
                .TextMatrix(iBaseRow + 1, 1) = "실적" Then
                Exit Sub
            End If
            
            
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
            
            Set tRs = oRapid.GetCheckDyeWorking(CLng(Left(.TextMatrix(iBaseRow, 20), 9)), CInt(Right(.TextMatrix(iBaseRow, 20), 2)))
            
            If tRs.RecordCount > 0 Then
                If (Trim(tRs!UseClss) = "작업" Or Len(Trim(tRs!UseClss)) = 8) And Left(tRs!procid, 2) = "43" Then
                    Set tRs = Nothing
                    Set oRapid = Nothing
                    MsgBox "선택되어진 건은 현재 작업중입니다" & vbCrLf & "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation, "편집 불가"
                    Call cmdRefesh_Click
                    Exit Sub
                Else
                    Set tRs = Nothing
                    If Trim(.TextMatrix(iBaseRow + 1, 20)) <> "" Then
                        Set sRs = oRapid.GetCheckDyeWorking(CLng(Left(.TextMatrix(iBaseRow + 1, 20), 9)), CInt(Right(.TextMatrix(iBaseRow + 1, 20), 2)))
                        If sRs.RecordCount > 0 Then
                            If (Trim(sRs!UseClss) = "작업" Or Len(Trim(sRs!UseClss)) = 8) And Left(sRs!procid, 2) = "43" Then
                                Set sRs = Nothing
                                Set oRapid = Nothing
                                MsgBox "옮기고자 하는 위치의 스케쥴에 변동이 생겼습니다" & vbCrLf & vbCrLf & _
                                        "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation, "이동 불가"
                                Call cmdRefesh_Click
                                Exit Sub
                            End If
                        End If
                        Set sRs = Nothing
                    End If
                End If
            Else
                Set tRs = Nothing
                Set oRapid = Nothing
                
                MsgBox "해당 건은 스케쥴 변동에 의해서 옮겨졌거나 존재하지 않는 건입니다" & vbCrLf & vbCrLf & _
                        "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation + vbOKOnly, "해당 건 없음"
                Call cmdRefesh_Click
                Exit Sub
            End If
            Set oRapid = Nothing
            
            .Rows = .Rows + 1
            
            For iCol = 2 To .Cols - 1
                .TextMatrix(.Rows - 1, iCol) = .TextMatrix(iBaseRow + 1, iCol)
            Next iCol
            For iCol = 2 To .Cols - 1
                .TextMatrix(iBaseRow + 1, iCol) = .TextMatrix(iBaseRow, iCol)
            Next iCol
            For iCol = 2 To .Cols - 1
                .TextMatrix(iBaseRow, iCol) = .TextMatrix(.Rows - 1, iCol)
            Next iCol
            
            .Rows = .Rows - 1
            
            .Col = 1
            .Row = iBaseRow + 1
            shpBox.Left = .CellLeft
            shpBox.Top = .CellTop
        End If
    End With

End Sub

Private Sub cmdEndCancel_Click()
    fraWorkEnd.Visible = False
    chkHold.Value = 0
End Sub

Private Sub cmdEndConfirm_Click()
    Dim oRapid As PlusLib2.CRapid
    Dim iChkHold As Integer
    Dim sHoldReason As String
    
    If Not CheckWorkEnd() Then Exit Sub
    
    If MsgBox("정말로 일지작성을 하시겠습니까?", vbYesNo + vbQuestion, "일지 작성 여부") = vbNo Then
        Exit Sub
    End If
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
        
    If chkHold.Value = 0 Then
        iChkHold = 0
        sHoldReason = ""
    Else
        iChkHold = 1
        sHoldReason = Trim(cboHold.Text)
    End If
    
    If oRapid.UpdatewkRapid(CLng(Left(lblSchSeq, 9)), CInt(Right(lblSchSeq, 2)), Format(CInt("0" & Left(lstArray(6).Text, 2)), "000"), _
                            lstArray(7).Text, lstArray(8).Text, Right(lstArray(9).Text, 8), Format(CInt(Left(lstArray(10).Text, 1)), "00"), _
                            txtRemarkResult, txtEndDate, txtEndTime, lblDyeClss, iChkHold, sHoldReason) Then
        Set oRapid = Nothing
        MsgBox "일지가 작성되었습니다", vbOKOnly, "작성 성공"
        fraWorkEnd.Visible = False
        chkHold.Value = 0
                            
        Call InitGrid
        Call InitGrdTab
        Call FillGridData
        Call FillSchData
    Else
        Set oRapid = Nothing
    End If
End Sub

Private Function CheckWorkEnd() As Boolean
Dim iCount%
    
    
    If lstArray(6).ListCount > 0 Then
        If lstArray(6).SelCount = 0 Then
            MsgBox "염색패턴이 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    End If
    If lstArray(7).SelCount = 0 Then
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    If lstArray(7).ListIndex > 0 Then
        If lstArray(8).SelCount > 0 Then
            MsgBox "염색구분이 선택되면 안됩니다", vbCritical, "작성 오류"
            Exit Function
        End If
    ElseIf lstArray(7).ListIndex = 0 Then
        If lstArray(8).SelCount = 0 Then
            MsgBox "염색구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    Else
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(9).SelCount = 0 Then
        MsgBox "작업자가 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(10).SelCount = 0 Then
        MsgBox "작업조가 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    
    If Trim(txtEndDate) = "" Or Len(Trim(txtEndDate)) < 8 Then
        MsgBox "종료일자가 올바르지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If Trim(txtEndDate) = "" Or Len(Trim(txtEndDate)) < 4 Then
        MsgBox "종료시간이 올바르지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    CheckWorkEnd = True
End Function

Private Sub cmdExpandWait_Click()
    If cmdExpandWait.Caption = "공정대기 확장" Then
        Call ToggleShapeBox(False, False)
        Call FoldOrUnFold(True)
    Else
        Call FoldOrUnFold(False)
    End If
End Sub

Private Sub cmdHide_Click()
Dim iRow%

    With grdTab(0)
        Call ToggleShapeBox(False, False)
        If cmdHide.Caption = "실적 감추기" Then
            For iRow = 1 To .Rows - 1
                If .TextMatrix(iRow, 1) = "실적" Then
                    .RowHeight(iRow) = 0
                End If
            Next iRow
            cmdHide.Caption = "실적 보이기"
        Else
            For iRow = 1 To .Rows - 1
                If .TextMatrix(iRow, 1) = "실적" Then
                    .RowHeight(iRow) = 600
                End If
            Next iRow
            cmdHide.Caption = "실적 감추기"
        End If
    End With
End Sub

Private Sub cmdInvisible_Click()
    fraWorkEnd.Visible = False
End Sub


Private Sub cmdRefesh_Click()
Dim i%

    Call ToggleShapeBox(False, False)
    Call InitGrid
    
    For i = 0 To lstArray.Count - 1
        lstArray(i).ListIndex = -1
    Next i
    bEnableWork = True
    pnlView.Visible = True
    pnlEdit.Visible = False
    cmdScreen.Caption = "편집화면"
    cmdConfirm.Caption = "카드선택"
    grdList(4).Rows = grdList(4).FixedRows
    cmdDelete.Visible = False
    cmdDelete.Enabled = False
    Call FillGridData
    Call FillSchData
End Sub


Private Sub cmdRefresh_Click()
    Call GridReSetting(cboCustom.Text, cboArticle.Text)
End Sub

Private Sub cmdSequence_Click()
    Dim oRapid As PlusLib2.CRapid
    Dim i%, j%, iCol%, iRow%, iSeq%, iMachID%
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
    
    bEnableWork = True
    g_adoCon.BeginTrans
    
    With grdTab(tabRapid.Tab)
        For i = .FixedRows To .Rows - 1
'            If .TextMatrix(i, 1) <> "실적" And .RowHeight(i) > 40 And .Cell(flexcpForeColor, i, 3) <> vbBlue Then
            If .TextMatrix(i, 1) <> "실적" And .RowHeight(i) > 40 Then
                If Trim(.TextMatrix(i, 20)) <> "" Then
                    If Not oRapid.UpdateRapidSeq(CLng(Left(.TextMatrix(i, 20), 9)), CInt(Right(.TextMatrix(i, 20), 2)), _
                                             "4300", Left(.TextMatrix(i, 0), 2), CInt(.TextMatrix(i, 1)), 2) Then
                        Set oRapid = Nothing
                        Exit Sub
                    End If
                End If
            End If
        Next i
    End With
    
    g_adoCon.CommitTrans
    
    Set oRapid = Nothing
    Call ToggleShapeBox(False, False)
    
    
    MsgBox "염색 스케쥴이 적용되었습니다", vbOKOnly, "저장 성공"
    
    Call InitGrid
    Call InitGrdTab
    Call FillGridData
    Call FillSchData

    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    g_adoCon.RollbackTrans

    Screen.MousePointer = vbDefault

    Set oRapid = Nothing
    Call ErrorBox(Err.Number, "frminstRapid.cmdSequence_Click", Err.Description)
End Sub

Private Sub cmdUP_Click()
Dim iCol%, iBaseRow%, iBaseCol%
Dim oRapid As PlusLib2.CRapid
Dim tRs As ADODB.Recordset
Dim sRs As ADODB.Recordset

    With grdTab(tabRapid.Tab)
        If .Row > .FixedRows Then
            bEnableWork = False
            Call VisibleWorkFrame(False)
            
            iBaseRow = .Row
            
            If .Cell(flexcpForeColor, iBaseRow - 1, 3) = vbBlue Or _
                .RowHeight(iBaseRow - 1) < 50 Or _
                .TextMatrix(iBaseRow - 1, 1) = "실적" Then
                Exit Sub
            End If
            
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
            
            Set tRs = oRapid.GetCheckDyeWorking(CLng(Left(.TextMatrix(iBaseRow, 20), 9)), CInt(Right(.TextMatrix(iBaseRow, 20), 2)))
            
            If tRs.RecordCount > 0 Then
                If (Trim(tRs!UseClss) = "작업" Or Len(Trim(tRs!UseClss)) = 8) And Left(tRs!procid, 2) = "43" Then
                    Set tRs = Nothing
                    Set oRapid = Nothing
                    MsgBox "선택되어진 건은 현재 작업중입니다" & vbCrLf & "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation, "편집 불가"
                    Call cmdRefesh_Click
                    Exit Sub
                Else
                    Set tRs = Nothing
                    If Trim(.TextMatrix(iBaseRow - 1, 20)) <> "" Then
                        Set sRs = oRapid.GetCheckDyeWorking(CLng(Left(.TextMatrix(iBaseRow - 1, 20), 9)), CInt(Right(.TextMatrix(iBaseRow - 1, 20), 2)))
                        If sRs.RecordCount > 0 Then
                            If (Trim(sRs!UseClss) = "작업" Or Len(Trim(sRs!UseClss)) = 8) And Left(sRs!procid, 2) = "43" Then
                                Set sRs = Nothing
                                Set oRapid = Nothing
                                MsgBox "옮기고자 하는 위치의 스케쥴에 변동이 생겼습니다" & vbCrLf & vbCrLf & _
                                        "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation, "이동 불가"
                                Call cmdRefesh_Click
                                Exit Sub
                            End If
                        End If
                        Set sRs = Nothing
                    End If
                End If
            Else
                Set tRs = Nothing
                Set oRapid = Nothing
                
                MsgBox "해당 건은 스케쥴 변동에 의해서 옮겨졌거나 존재하지 않는 건입니다" & vbCrLf & vbCrLf & _
                        "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation + vbOKOnly, "해당 건 없음"
                Call cmdRefesh_Click
                Exit Sub
            End If
            Set oRapid = Nothing
            
            .Rows = .Rows + 1
            
            For iCol = 2 To .Cols - 1
                .TextMatrix(.Rows - 1, iCol) = .TextMatrix(iBaseRow - 1, iCol)
            Next iCol
            For iCol = 2 To .Cols - 1
                .TextMatrix(iBaseRow - 1, iCol) = .TextMatrix(iBaseRow, iCol)
            Next iCol
            For iCol = 2 To .Cols - 1
                .TextMatrix(iBaseRow, iCol) = .TextMatrix(.Rows - 1, iCol)
            Next iCol
            
            .Rows = .Rows - 1
            
            .Col = 1
            .Row = iBaseRow - 1
            shpBox.Left = .CellLeft
            shpBox.Top = .CellTop
        End If
    End With

End Sub

Private Sub cmdWorkEnd_Click()
Dim idx%, iCol%, iCntRec%
Dim nSchID As Long
Dim nSeq As Integer
Dim oRapid As PlusLib2.CRapid
Dim rs As Recordset
Dim sWorkJo$, sDyeClss$
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

    chkHold.Value = 0
    idx = tabRapid.Tab
    With grdTab(idx)
        nSchID = CLng(Left(.TextMatrix(.Row, 20), 9))
        nSeq = CInt(Right(.TextMatrix(.Row, 20), 2))
        
        If (.TextMatrix(.Row, 8) = "0" And .TextMatrix(.Row, 9) = "0") Or (Trim(.TextMatrix(.Row, 8)) = "" And Trim(.TextMatrix(.Row, 9)) = "") Then
            sDyeClss = "비염색"
            lblDyeClss = "비염색"
        Else
            sDyeClss = "염색"
            lblDyeClss = "염색"
        End If
        
        If sDyeClss = "비염색" Then
            iCntRec = 1
        Else
            If nSeq > 1 Then        ' 염색 추가작업일때는 전산상의 대기공정에 상관없이 진행
                Set rs = oRapid.GetWaitWorkDyeProcCard(nSchID, nSeq, "")
                iCntRec = rs.RecordCount
                rs.Close
                Set rs = Nothing
            Else
                Set rs = oRapid.GetWaitWorkDyeProcCard(nSchID, nSeq, "작업")
                iCntRec = rs.RecordCount
                rs.Close
                Set rs = Nothing
            End If
        End If
        If iCntRec = 0 Then
            Set oRapid = Nothing
            MsgBox "현재 염색공정에서 작업중이 아닙니다." & vbCrLf & vbCrLf & _
                    "작업 완료가 안됩니다", vbCritical, "작업완료 불가"
            Exit Sub
        End If
        Set oRapid = Nothing
        
        Call ToggleShapeBox(False, False)
        Call InitFraWorkEnd
        lblSchSeq = .TextMatrix(.Row, 20)
        Call LoadRapidWorkData(nSchID, nSeq)
        fraWorkEnd.Visible = True
    End With
End Sub

Private Sub InitFraWorkEnd()
Dim idx%

    For idx = 7 To 10
        lstArray(idx).ListIndex = -1
    Next idx
End Sub

Private Sub LoadRapidWorkData(SchID As Long, Seq As Integer)
Dim oRapid As PlusLib2.CRapid
Dim rs As Recordset
Dim tRs As Recordset
Dim i%, iCount%
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName


    Set rs = oRapid.GetwiRapidData(SchID, Seq)

    If rs.RecordCount > 0 Then
        txtRoll = rs!wiroll
        txtQty = Format(rs!wiqty, "###,##0")
        ' 염색패턴
        Set tRs = oRapid.GetDyePatternList(1, CInt(rs!wimachid), 0)
        
        lstArray(6).Clear
        For iCount = 1 To tRs.RecordCount
            lstArray(6).AddItem Format(tRs!PtNo, "00") & ". " & tRs!PtName
            tRs.MoveNext
        Next iCount
        tRs.Close
        Set tRs = Nothing
        
        For i = 0 To lstArray(6).ListCount - 1
            If Left(lstArray(6).List(i), 2) = Format(CInt(rs!PatternID), "00") Then
                lstArray(6).Selected(i) = True
                Exit For
            End If
        Next i
        
        ' 작업구분
        For i = 0 To lstArray(7).ListCount - 1
            If lstArray(7).List(i) = rs!workclss Then
                lstArray(7).Selected(i) = True
                Exit For
            End If
        Next i
        ' 염색구분
        For i = 0 To lstArray(8).ListCount - 1
            If lstArray(8).List(i) = rs!RapidClss Then
                lstArray(8).Selected(i) = True
                Exit For
            End If
        Next i
        ' 작업자
        For i = 0 To lstArray(9).ListCount - 1
            If Right(lstArray(9).List(i), 8) = Format(rs!PersonID, "00000000") Then
                lstArray(9).Selected(i) = True
                Exit For
            End If
        Next i
        ' 작업조
        For i = 0 To lstArray(10).ListCount - 1
            If Left(lstArray(10).List(i), 1) = CStr(CInt(rs!TeamID)) Then
                lstArray(10).Selected(i) = True
                Exit For
            End If
        Next i
        
    End If
    rs.Close
    Set rs = Nothing
    Set oRapid = Nothing
    
    txtEndDate = Format(Now, "YYYYMMDD")
    txtEndTime = Format(time, "HHMM")
End Sub

Private Sub cmdWorkStart_Click()
Dim idx%, iCol%, iCntRec%, iCount%
Dim nSchID As Long
Dim nSeq As Integer
Dim nMachID As Integer
Dim oRapid As PlusLib2.CRapid
Dim rs As Recordset
Dim sWorkJo$
Dim sTeamMsg$
Dim sDyeClss$       ' 염색, 비염색 구분
Dim sInstClss$
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

    idx = tabRapid.Tab
    With grdTab(idx)
        nSchID = CLng(Left(.TextMatrix(.Row, 20), 9))
        nSeq = CInt(Right(.TextMatrix(.Row, 20), 2))
        nMachID = CInt(Left(.TextMatrix(.Row, 0), 2))
        
        Set rs = oRapid.GetCheckFirstWork("4300", nMachID)
        If rs.RecordCount > 0 Then
            If Not (rs!dyeschid = nSchID And rs!DyeSeq = nSeq) Then
                Set rs = Nothing
                Set oRapid = Nothing
                
                MsgBox "해당 건은 스케쥴 변동에 의해서 옮겨졌거나 존재하지 않는 건입니다" & vbCrLf & vbCrLf & _
                        "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation + vbOKOnly, "해당 건 없음"
                Call cmdRefesh_Click
                Exit Sub
            End If
        Else
            Set rs = Nothing
            Set oRapid = Nothing
            
            MsgBox "해당 건은 스케쥴 변동에 의해서 옮겨졌거나 존재하지 않는 건입니다" & vbCrLf & vbCrLf & _
                    "화면을 Refresh하오니 확인후 재작업해주시기 바랍니다", vbExclamation + vbOKOnly, "해당 건 없음"
            Call cmdRefesh_Click
            Exit Sub
        End If
        Set rs = Nothing
        
        
        For iCount = 0 To lstArray(10).ListCount - 1
            sTeamMsg = sTeamMsg & lstArray(10).List(iCount) & ",  "
        Next iCount
        If (.TextMatrix(.Row, 8) = "0" And .TextMatrix(.Row, 9) = "0") Or (Trim(.TextMatrix(.Row, 8)) = "" And Trim(.TextMatrix(.Row, 9)) = "") Then
            sDyeClss = "비염색"
        Else
            sDyeClss = "염색"
        End If
        
        If sDyeClss = "비염색" Then
            iCntRec = 1
        Else
            If nSeq > 1 Then        ' 염색 추가작업일때는 전산상의 대기공정에 상관없이 진행
                Set rs = oRapid.GetWaitWorkDyeProcCard(nSchID, nSeq, "")
                If rs.RecordCount > 0 Then
                    iCntRec = rs.RecordCount
                    sInstClss = Trim(rs!instclss)
                End If
                rs.Close
                Set rs = Nothing
            Else
                Set rs = oRapid.GetWaitWorkDyeProcCard(nSchID, nSeq, "대기")
                If rs.RecordCount > 0 Then
                    iCntRec = rs.RecordCount
                    sInstClss = Trim(rs!instclss)
                End If
                rs.Close
                Set rs = Nothing
            End If
        End If
        If iCntRec = 0 Then
            Set oRapid = Nothing
            MsgBox "현재 염색공정이 아닌 다른 공정에 대기하고 있어" & vbCrLf & vbCrLf & _
                    "작업 시작이 안됩니다", vbCritical, "작업시작 불가"
            Exit Sub
        Else
            If sDyeClss = "염색" And sInstClss = "" Then
                Set oRapid = Nothing
                MsgBox "평량지시가 내려지지 않은 건은 시작이 불가합니다", vbCritical, "시작 불가"
                Exit Sub
            End If
            Do
                sWorkJo = InputBox("작업조를 입력하여 주십시요(1 ~ 3)" & vbCrLf & vbCrLf & _
                                sTeamMsg, "작업조 입력")
                If Trim(sWorkJo) = "" Then
                    Set oRapid = Nothing
                    Exit Sub
                Else
                    If CInt(sWorkJo) >= 1 And CInt(sWorkJo) <= 3 Then
                        Exit Do
                    End If
                End If
            Loop
            Call ToggleShapeBox(False, False)
            
            If oRapid.AddwkRapid(nSchID, nSeq, Format(CInt(sWorkJo), "00"), sDyeClss) Then
                Set oRapid = Nothing
                MsgBox "작업이 시작되었습니다", vbOKOnly, "작업 시작"
                Call InitGrid
                Call InitGrdTab
                Call FillGridData
                Call FillSchData
'                grdTab(0).Cell(flexcpFontBold, 2, 1, grdTab(0).Rows - 1, grdTab(0).Cols - 1) = False
'                grdTab(1).Cell(flexcpFontBold, 2, 1, grdTab(1).Rows - 1, grdTab(1).Cols - 1) = False
            End If
        End If
    End With
End Sub




Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    Call cmdRefesh_Click
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub

Private Sub Form_Load()
    Dim i%
    
    bEnableWork = True
    
    cmdUP.Picture = LoadResPicture("UP", vbResIcon)
    cmdDown.Picture = LoadResPicture("DOWN", vbResIcon)
'    cmdLeft.Picture = LoadResPicture("LEFT", vbResIcon)
'    cmdRight.Picture = LoadResPicture("RIGHT", vbResIcon)


    Me.Move 0, 0, 15360, 9840
    
'    Call ToggleShapeBox(False, False)
'    Call InitGrid
    Call InitTab
    Call AddLstBox
'    Call FillSchData
    
    nMachineID = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub

Private Sub cmdConfirm_Click()
Dim nRoll As Long
Dim nQty As Long
Dim nChange As Integer

    Call ToggleShapeBox(False, False)

    If cmdConfirm.Caption = "카드선택" Then
        Dim i%, iRow%, iCol%, iSeq%
        Dim iCntA%, iCntB%
        Dim lTotRoll As Long, lTotQty As Long
        Dim sRec
        
        lblWork = ""
        grdList(4).Rows = grdList(4).FixedRows
        For i = 0 To 3
            With grdList(i)
                If .Rows > .FixedRows Then
                    For iRow = 1 To .Rows - 1
                        If .Cell(flexcpChecked, iRow, 0, iRow, 0) = flexChecked Then
                            If i < 3 Then
                                iCntA = iCntA + 1
                            Else
                                iCntB = iCntB + 1
                            End If
                            iSeq = iSeq + 1
                            grdList(4).Rows = grdList(4).Rows + 1
                            grdList(4).RowHeight(grdList(4).Rows - 1) = 300
                            grdList(4).Cell(flexcpChecked, grdList(4).Rows - 1, 0, grdList(4).Rows - 1, 0) = flexChecked
                            lTotRoll = lTotRoll + CLng(.TextMatrix(iRow, 12))
                            lTotQty = lTotQty + CLng(.TextMatrix(iRow, 13))
                            
                            For iCol = 1 To .Cols - 1
                                If iCol = 4 Then
                                    grdList(4).TextMatrix(grdList(4).Rows - 1, iCol) = CStr(iSeq)
                                Else
                                    grdList(4).TextMatrix(grdList(4).Rows - 1, iCol) = .TextMatrix(iRow, iCol)
                                End If
                            Next iCol
                        End If
                    Next iRow
                End If
            End With
        Next i
        If iCntB > 0 And iCntA > 0 Then
            MsgBox "염색대기 카드와 염색완료 카드를 혼용할 수 없습니다", vbCritical, "작성 오류"
            Exit Sub
        End If
        If iCntB > 0 Then
            lblWork = "추가작업"
        Else
            lblWork = ""
        End If
        
        grdList(4).Rows = grdList(4).Rows + 1
        grdList(4).RowHeight(grdList(4).Rows - 1) = 300
        grdList(4).Cell(flexcpText, grdList(4).Rows - 1, 0, grdList(4).Rows - 1, 11) = "선택되어진 카드 총 합계"
        grdList(4).Cell(flexcpFontBold, grdList(4).Rows - 1, 0, grdList(4).Rows - 1, grdList(4).Cols - 1) = True
        grdList(4).TextMatrix(grdList(4).Rows - 1, 12) = Format(lTotRoll, "#,##0")
        grdList(4).TextMatrix(grdList(4).Rows - 1, 13) = Format(lTotQty, "#,###,##0")
        grdList(4).MergeCells = flexMergeRestrictRows
        grdList(4).MergeRow(grdList(4).Rows - 1) = True
        
        For i = 0 To lstArray.Count - 1
            lstArray(i).ListIndex = -1
        Next i
        pnlView.Visible = False
        pnlEdit.Visible = True
        cmdConfirm.Caption = "염색지시"
        cmdScreen.Caption = "편집취소"
        Call FoldOrUnFold(False)
        cmdExpandWait.Visible = False
    ElseIf cmdConfirm.Caption = "염색지시" Then
        If Not CheckData Then Exit Sub
        
        If MsgBox("염색스케쥴에 적용시키겠습니까?", vbYesNo + vbQuestion, "최종 확인") = vbYes Then
            If lstArray(5).ListIndex > 0 Then
                nRoll = 0
                nQty = 0
            Else
                nRoll = CLng(grdList(4).TextMatrix(grdList(4).Rows - 1, 12))
                nQty = CLng(grdList(4).TextMatrix(grdList(4).Rows - 1, 13))
            End If
            If AddData(nRoll, nQty) Then
                Screen.MousePointer = vbHourglass
                Call InitGrid
                Call InitGrdTab
                Call FillSchData
                Call cmdScreen_Click
                Screen.MousePointer = vbDefault
            End If
        End If
    ElseIf cmdConfirm.Caption = "저장" Then
        If Not CheckData Then Exit Sub
        
        If lblMachID <> Left(lstArray(0).Text, 2) Then
            nChange = 1     ' 호기를 변경하여 저장할때
        Else
            nChange = 0
        End If
        If MsgBox("염색스케쥴에 적용시키겠습니까?", vbYesNo + vbQuestion, "최종 확인") = vbYes Then
            If UpdateData(CLng(Left(lblSchIDSeq, 9)), CInt(Right(lblSchIDSeq, 2)), _
                CLng(grdList(4).TextMatrix(grdList(4).Rows - 1, 12)), CLng(grdList(4).TextMatrix(grdList(4).Rows - 1, 13)), nChange) Then
                Screen.MousePointer = vbHourglass
                Call InitGrid
                Call InitGrdTab
                Call FillSchData
                Call cmdScreen_Click
                Screen.MousePointer = vbDefault
            Else
                MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "저장도중 에러"
                Screen.MousePointer = vbDefault
            End If
            Call FoldOrUnFold(False)
        End If
    End If
End Sub

Private Function CheckData() As Boolean
    Dim iRow%, iCol%, iCount%, iChkCnt%
    
    If lstArray(0).SelCount = 0 Then
        MsgBox "염색호기가 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If CInt(Left(lstArray(0).Text, 2)) < 12 Then
        If lstArray(1).SelCount = 0 Then
            MsgBox "염색패턴이 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    End If
    If lstArray(5).SelCount = 0 Then
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(5).ListIndex > 0 Then
        If lstArray(2).SelCount > 0 Then
            MsgBox "염색구분이 선택되면 안됩니다", vbCritical, "작성 오류"
            Exit Function
        End If
    ElseIf lstArray(5).ListIndex = 0 Then
        If lstArray(2).SelCount = 0 Then
            MsgBox "염색구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    Else
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(3).SelCount = 0 Then
        MsgBox "작업자가 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(5).ListIndex = 0 Then
        If grdList(4).Rows = grdList(4).FixedRows + 1 Then
            MsgBox "카드가 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    End If
        
    With grdList(4)
        For iRow = 1 To .Rows - 2
            If .Cell(flexcpChecked, iRow, 0, iRow, 0) = flexChecked Then
                iCount = iCount + 1
            End If
            If .TextMatrix(iRow, 7) = "미확정" Then
                iChkCnt = iChkCnt + 1
            End If
        Next iRow
    End With
'    If iCount = 0 Then
'        MsgBox "카드가 선택되어 있지 않습니다", vbCritical, "작성 오류"
'        Exit Function
'    End If
    If iChkCnt > 0 Then
        MsgBox "색상이 미확정인 카드는 염색지시를 내릴수 없습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
'    With grdTab(0)
'        iRow = CInt(Left(lstArray(0).Text, 2)) * 2
'        iCol = 1 + ((CInt(lstArray(4).Text) - 1) * 5)
'        If .Cell(flexcpForeColor, iRow, iCol) = vbBlue Then
'            MsgBox "현재 작업중인 위치에 스케쥴을 등록할 수 없습니다", vbCritical, "작성 오류"
'            Exit Function
'        End If
'    End With
    
    CheckData = True
End Function

Private Function AddData(TotRoll As Long, TotQty As Long) As Boolean
    Dim oRapid As PlusLib2.CRapid
    Dim tCardList() As PlusLib2.tRapidCard
    Dim i%, iCount%, iCntChk%, iCol%, iRow%, iSeq%
    
    Screen.MousePointer = vbHourglass
    AddData = False

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

    With grdList(4)
        For i = .FixedRows To .Rows - 2
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                iCntChk = iCntChk + 1
            End If
        Next i
        If lstArray(5).ListIndex > 0 Then
            ReDim tCardList(1)
            tCardList(iCount).sCardID = ""
            tCardList(iCount).sSplitID = ""
            tCardList(iCount).lDyeSchID = 0
        Else
            ReDim tCardList(iCntChk)
            iCount = 0
            For i = .FixedRows To .Rows - 2
                If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    tCardList(iCount).sCardID = Trim(.TextMatrix(i, 17))
                    tCardList(iCount).sSplitID = IIf(Trim(.TextMatrix(i, 18)) = "", " ", Trim(.TextMatrix(i, 18)))
                    If lstArray(2).Text = "추가" Then
                        tCardList(iCount).lDyeSchID = CLng(.TextMatrix(i, 23))
                    Else
                        tCardList(iCount).lDyeSchID = 0
                    End If
                    iCount = iCount + 1
                End If
            Next i
        End If
        
    End With
    
    g_adoCon.BeginTrans
    
    If Not oRapid.AddNewwiRapidItem(tCardList(), CLng(tCardList(0).lDyeSchID), "4300", Left(lstArray(0).Text, 2), _
        0, lstArray(5).Text, lstArray(2).Text, Format(CInt("0" & Left(lstArray(1).Text, 2)), "000"), 0, TotRoll, _
        TotQty, " ", Right(lstArray(3).Text, 8), CheckNull(txtRemark)) Then
        Set oRapid = Nothing
        AddData = False
        Exit Function
    End If
    
    AddData = True
    g_adoCon.CommitTrans
    
    Set oRapid = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    AddData = False

    Set oRapid = Nothing
    Call ErrorBox(Err.Number, "frminstRapid.AddData", Err.Description)
End Function

Private Function UpdateData(lDyeSchID As Long, iDyeSeq As Integer, TotRoll As Long, TotQty As Long, iChkChange As Integer) As Boolean
    Dim oRapid As PlusLib2.CRapid
    Dim i%, iCol%, iRow%, iSeq%
    
    Screen.MousePointer = vbHourglass
    UpdateData = False

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
    
    g_adoCon.BeginTrans
    
    If Not oRapid.UpdatewiRapid(lDyeSchID, iDyeSeq, "4300", Left(lstArray(0).Text, 2), 0, _
        lstArray(5).Text, lstArray(2).Text, Format(CInt("0" & Left(lstArray(1).Text, 2)), "000"), 0, _
        TotRoll, TotQty, Right(lstArray(3).Text, 8), IIf(Trim(txtRemark) = "", " ", Trim(txtRemark)), iChkChange) Then
        Set oRapid = Nothing
        UpdateData = False
        Exit Function
    End If
    
    UpdateData = True
    g_adoCon.CommitTrans
    
    Set oRapid = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    UpdateData = False

    Set oRapid = Nothing
    Call ErrorBox(Err.Number, "frminstRapid.UpdateData", Err.Description)
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdScreen_Click()
Dim i%
    
    tabRapid.Top = 3390
    tabRapid.Height = 5325
    grdTab(0).Height = 4920
    grdTab(1).Height = 4920
    Call ToggleShapeBox(False, False)
    
    If cmdScreen.Caption = "편집취소" Then
        pnlView.Visible = True
        pnlEdit.Visible = False
        cmdScreen.Caption = "편집화면"
        cmdConfirm.Caption = "카드선택"
        For i = 0 To lstArray.Count - 1
            lstArray(i).ListIndex = -1
        Next i
        grdList(4).Rows = grdList(4).FixedRows
        cmdDelete.Visible = False
        cmdDelete.Enabled = False
        cmdDelete.Enabled = True
        cmdConfirm.Enabled = True
        
        Call FillGridData
        grdTab(0).Cell(flexcpFontBold, 2, 1, grdTab(0).Rows - 1, grdTab(0).Cols - 1) = False
        cmdExpandWait.Visible = True
        
'        grdTab(1).Cell(flexcpFontBold, 2, 1, grdTab(1).Rows - 1, grdTab(1).Cols - 1) = False
    ElseIf cmdScreen.Caption = "편집화면" Then
        pnlView.Visible = False
        pnlEdit.Visible = True
        cmdScreen.Caption = "편집취소"
        cmdConfirm.Caption = "염색지시"
    ElseIf cmdScreen.Caption = "취소" Then
        For i = 0 To lstArray.Count - 1
            lstArray(i).ListIndex = -1
        Next i
        grdList(4).Rows = grdList(4).FixedRows
        cmdDelete.Enabled = False
        cmdConfirm.Enabled = False
    End If
End Sub

Private Sub cmdToggle_Click()
    Call ToggleShapeBox(False, False)
    If cmdToggle.Caption = "염색 스케쥴 작성" Then
        pnlMsg.Caption = "입력 중 입니다...."
        Call MoveScreen(True)
        cmdScreen.Caption = "편집취소"
        cmdConfirm.Caption = "염색지시"
        cmdConfirm.Enabled = True
        cmdExpandWait.Visible = True
        cmdExpandWait.Caption = "공정대기 확장"
        If cmdExpandWait.Caption = "공정대기 확장" Then
            Call FoldOrUnFold(True)
        Else
            Call FoldOrUnFold(False)
        End If
        Call cmdScreen_Click
    Else
        Call MoveScreen(False)
'        grdTab(0).Cell(flexcpFontBold, 2, 1, grdTab(0).Rows - 1, grdTab(0).Cols - 1) = False
'        grdTab(1).Cell(flexcpFontBold, 2, 1, grdTab(1).Rows - 1, grdTab(1).Cols - 1) = False
        grdTab(0).Cell(flexcpForeColor, 2, 1, grdTab(0).Rows - 1, grdTab(0).Cols - 1) = vbBlack
        grdTab(1).Cell(flexcpForeColor, 2, 1, grdTab(1).Rows - 1, grdTab(1).Cols - 1) = vbBlack
        
        Call InitGrdTab
        Call FillSchData
        bEnableWork = True
        cmdExpandWait.Visible = False
        tabRapid.Visible = True
    End If
    cmdDelete.Visible = False
End Sub

Private Sub MoveScreen(bFlag As Boolean)
    If bFlag = True Then    ' 화면 분할
        tabRapid.Height = 5325
        tabRapid.Top = 3390
        grdTab(0).Height = 4920
        grdTab(1).Height = 4920
        fraButton.Visible = bFlag
        cmdToggle.Caption = "염색 스케쥴 조회"
    Else
        tabRapid.Top = 510
        tabRapid.Height = 8235
        grdTab(0).Height = 7830
        grdTab(1).Height = 7830
        fraButton.Visible = bFlag
        pnlMsg.Caption = "조회 중 입니다...."
        cmdToggle.Caption = "염색 스케쥴 작성"
    End If
End Sub
Private Sub AddLstBox()
    Dim oRapid As PlusLib2.CRapid
    Dim oPerson  As PlusLib2.CPerson
    Dim rs As Recordset
    Dim iCount%, i%, iSeq%
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    txtRemark = ""
    For i = 0 To lstArray.Count - 1
        lstArray(i).Clear
    Next i
    cboHold.Clear
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
                
    Select Case Custom
        Case "삼우":
            Set rs = oRapid.GetMachineNoList("Rapid염색기")
            For iCount = 1 To rs.RecordCount
                lstArray(0).AddItem Format(rs!MachineNO, "00") & " 호기"
                rs.MoveNext
            Next iCount
            rs.Close
            Set rs = Nothing
            lstArray(0).AddItem "12 호기"
            
    End Select
    
    Set oRapid = Nothing
    
' 삼우DFC의 염색구분 목록
    lstArray(2).AddItem "본염"
    lstArray(2).AddItem "얼룩수정"
    lstArray(2).AddItem "주름수정"
    lstArray(2).AddItem "오염수정"
    lstArray(2).AddItem "색수정"
    lstArray(2).AddItem "탈발후 색수정"
    lstArray(2).AddItem "탈색후 재염"
    lstArray(2).AddItem "탈색"
    lstArray(2).AddItem "감색"
    lstArray(2).AddItem "추가"
    
    lstArray(8).AddItem "본염"
    lstArray(8).AddItem "얼룩수정"
    lstArray(8).AddItem "주름수정"
    lstArray(8).AddItem "오염수정"
    lstArray(8).AddItem "색수정"
    lstArray(8).AddItem "탈발후 색수정"
    lstArray(8).AddItem "탈색후 재염"
    lstArray(8).AddItem "탈색"
    lstArray(8).AddItem "감색"
    lstArray(8).AddItem "추가"
   
    cboHold.AddItem "얼룩수정"
    cboHold.AddItem "주름수정"
    cboHold.AddItem "오염수정"
    cboHold.AddItem "색수정"
    cboHold.AddItem "탈발후 색수정"
    cboHold.AddItem "탈색후 재염"
    cboHold.AddItem "탈색"
    cboHold.AddItem "감색"
    cboHold.AddItem "추가"
    
    
' 삼우DFC의 작업구분
    lstArray(5).AddItem "염색"
    lstArray(5).AddItem "BOX 탈색"
    lstArray(5).AddItem "BOX R/C"
    lstArray(5).AddItem "도포 Washing"
    lstArray(5).AddItem "Soaping"
    lstArray(5).AddItem "기계수리"
    
    lstArray(7).AddItem "염색"
    lstArray(7).AddItem "BOX 탈색"
    lstArray(7).AddItem "BOX R/C"
    lstArray(7).AddItem "도포 Washing"
    lstArray(7).AddItem "Soaping"
    lstArray(7).AddItem "기계수리"
    
    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    oPerson.UserName = g_sUserName
    Set rs = oPerson.GetWorkerList("13")     '염색 부서
'    Set rs = oPerson.GetWorkerList("")     '염색 부서
    For iCount = 1 To rs.RecordCount
        lstArray(3).AddItem rs!Name & "             " & Format(rs!PersonID, "00000000")
        lstArray(9).AddItem rs!Name & "             " & Format(rs!PersonID, "00000000")
        rs.MoveNext
    Next iCount
    rs.Close
    Set rs = Nothing
    
    Set rs = oPerson.GetWorkTeam()     '작업 조
    For iCount = 1 To rs.RecordCount
        lstArray(10).AddItem CStr(CInt(rs!TeamID)) & ". " & rs!Team
        rs.MoveNext
    Next iCount
    rs.Close
    Set rs = Nothing
    
    Set oPerson = Nothing
    
'    For iCount = 1 To 10
'        lstArray(4).AddItem Format(iCount, "00")
'    Next iCount

    Screen.MousePointer = vbDefault


    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    Set oPerson = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.AddLstBox", Err.Description)
End Sub

Private Sub InitTab()
    With tabRapid
        .TabCaption(0) = "1 ~ 11호기(Rapid)"
        pnlTab(0) = "1 ~ 11호기(Rapid)"
        .TabCaption(1) = "12호기(C.P.B)"
        pnlTab(1) = "12호기(C.P.B)"
        .TabVisible(1) = False
    End With
End Sub

Private Sub InitGrid()
    Dim i%

    For i = 0 To 4
        Call SetVSFlexGrid(grdList(i))

        With grdList(i)
            .WordWrap = False
            .Redraw = flexRDNone
    
            .Rows = 1:      .Cols = 25
            .RowHeight(0) = 300
            .FixedRows = 1:     .FixedCols = 0
            .Editable = flexEDKbdMouse
            .SelectionMode = flexSelectionFree
            .HighLight = flexHighlightNever
            .ExplorerBar = flexExNone
            .FocusRect = flexFocusSolid
            
            .TextArray(0) = "":                     .ColWidth(0) = 250:         .ColAlignment(0) = flexAlignCenterCenter
            .TextArray(1) = "밧자번호":             .ColWidth(1) = 0:           .ColAlignment(1) = flexAlignLeftCenter
            .TextArray(2) = "밧자순위":             .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
            .TextArray(3) = "밧자":                 .ColWidth(3) = 500:         .ColAlignment(3) = flexAlignLeftCenter
            .TextArray(4) = "No":                   .ColWidth(4) = 300:         .ColAlignment(4) = flexAlignCenterCenter
            .TextArray(5) = "거래처":               .ColWidth(5) = 1100:        .ColAlignment(5) = flexAlignLeftCenter
            .TextArray(6) = "품명":                 .ColWidth(6) = 2500:        .ColAlignment(6) = flexAlignLeftCenter
            .TextArray(7) = "색상":                 .ColWidth(7) = 2000:        .ColAlignment(7) = flexAlignLeftCenter
            .TextArray(8) = "관리번호":             .ColWidth(8) = 1200:           .ColAlignment(8) = flexAlignLeftCenter
            .TextArray(9) = "카드번호":             .ColWidth(9) = 1000:        .ColAlignment(9) = flexAlignLeftCenter
            .TextArray(10) = "분할":                .ColWidth(10) = 500:        .ColAlignment(10) = flexAlignLeftCenter
            .TextArray(11) = "대기":                .ColWidth(11) = 800:        .ColAlignment(11) = flexAlignLeftCenter
            .TextArray(12) = "절수":                .ColWidth(12) = 600:        .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "수량":                .ColWidth(13) = 800:        .ColAlignment(13) = flexAlignRightCenter
            .TextArray(14) = "거래처코드":          .ColWidth(14) = 0:          .ColAlignment(14) = flexAlignLeftCenter
            .TextArray(15) = "품명코드":            .ColWidth(15) = 0:          .ColAlignment(15) = flexAlignLeftCenter
            .TextArray(16) = "색상코드":            .ColWidth(16) = 0:          .ColAlignment(16) = flexAlignLeftCenter
            .TextArray(17) = "카드번호":            .ColWidth(17) = 0:          .ColAlignment(17) = flexAlignCenterCenter
            .TextArray(18) = "분할":                .ColWidth(18) = 0:          .ColAlignment(18) = flexAlignLeftCenter
'            .TextArray(19) = "대기공정코드":        .ColWidth(19) = 0:          .ColAlignment(19) = flexAlignLeftCenter
            .TextArray(19) = "제직처":              .ColWidth(19) = 900:          .ColAlignment(19) = flexAlignLeftCenter
            .TextArray(20) = "관리번호":            .ColWidth(20) = 0:          .ColAlignment(20) = flexAlignLeftCenter
            .TextArray(21) = "OrderSeq":            .ColWidth(21) = 0:        .ColAlignment(21) = flexAlignLeftCenter
            .TextArray(22) = "계획 후공정":         .ColWidth(22) = 2000:       .ColAlignment(22) = flexAlignLeftCenter
            .TextArray(23) = "스케쥴번호":          .ColWidth(23) = 0:          .ColAlignment(23) = flexAlignLeftCenter
            .TextArray(24) = "차수":                .ColWidth(24) = 0:          .ColAlignment(24) = flexAlignLeftCenter
            If i = 4 Then
                .ColWidth(5) = 800
                .ColWidth(6) = 2300
                .ColWidth(7) = 1800
                .ColWidth(8) = 0
                .ColWidth(11) = 0
                .ColWidth(22) = 0
            End If
            .Redraw = flexRDDirect
        End With
    Next i
End Sub

Private Sub InitGrdTab()
    Dim i%, iCol%, iRow%, iNo%, iTabCnt%

    txtRemark = ""
    For i = 1 To tabRapid.Tabs
        If tabRapid.TabVisible(i - 1) = True Then
            iTabCnt = iTabCnt + 1
        End If
    Next i
        
    For i = 0 To iTabCnt - 1
        cmdHide.Caption = "실적 감추기"
        Call SetVSFlexGrid(grdTab(i))

        With grdTab(i)
            .Rows = 0
            .Redraw = flexRDDirect
            .SelectionMode = flexSelectionFree
            .ExplorerBar = flexExNone
            .ExtendLastCol = True
            .RowHeightMin = 0
            
            .Rows = nMachNo * 10 + 1
            .Cols = 30
            .FixedRows = 1:     .FixedCols = 1
            .RowHeight(0) = 350
            .RowHeight(1) = 0
            
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .HighLight = flexHighlightNever
            
            For iCol = 0 To .Cols - 1
                .ColWidth(iCol) = 0
            Next iCol
            
            .TextMatrix(0, 0) = "염색기":           .ColWidth(0) = 700:     .ColAlignment(0) = flexAlignCenterCenter
            .TextMatrix(0, 1) = "번호":             .ColWidth(1) = 500:     .ColAlignment(1) = flexAlignCenterCenter
            .TextMatrix(0, 2) = "일자":             .ColWidth(2) = 0:       .ColAlignment(2) = flexAlignCenterCenter
            .TextMatrix(0, 3) = "거래처":           .ColWidth(3) = 1000:    .ColAlignment(3) = flexAlignLeftCenter
            .TextMatrix(0, 4) = "관리번호":         .ColWidth(4) = 900:    .ColAlignment(4) = flexAlignCenterCenter
            .TextMatrix(0, 5) = "품명":             .ColWidth(5) = 2600:    .ColAlignment(5) = flexAlignLeftCenter
            .TextMatrix(0, 6) = "색상":             .ColWidth(6) = 1800:    .ColAlignment(6) = flexAlignLeftCenter
            .TextMatrix(0, 7) = "카드번호":         .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignLeftCenter
            .TextMatrix(0, 8) = "절수":             .ColWidth(8) = 600:     .ColAlignment(8) = flexAlignRightCenter
            .TextMatrix(0, 9) = "수량":             .ColWidth(9) = 700:     .ColAlignment(9) = flexAlignRightCenter
            .TextMatrix(0, 10) = "밧쟈":            .ColWidth(10) = 700:    .ColAlignment(10) = flexAlignCenterCenter
            .TextMatrix(0, 11) = "처방":            .ColWidth(11) = 500:    .ColAlignment(11) = flexAlignCenterCenter
            .TextMatrix(0, 12) = "평량":            .ColWidth(12) = 500:    .ColAlignment(12) = flexAlignCenterCenter
            .TextMatrix(0, 13) = "작업구분":        .ColWidth(13) = 1000:   .ColAlignment(13) = flexAlignCenterCenter
            .TextMatrix(0, 14) = "염색구분":        .ColWidth(14) = 1000:   .ColAlignment(14) = flexAlignCenterCenter
            .TextMatrix(0, 15) = "대기공정":        .ColWidth(15) = 0:      .ColAlignment(15) = flexAlignCenterCenter
            .TextMatrix(0, 16) = "비고사항":        .ColWidth(16) = 1000:   .ColAlignment(16) = flexAlignLeftCenter
            .TextMatrix(0, 17) = "완료공정":        .ColWidth(17) = 0:      .ColAlignment(17) = flexAlignCenterCenter
            .TextMatrix(0, 18) = "총절수":          .ColWidth(18) = 0:      .ColAlignment(18) = flexAlignCenterCenter
            .TextMatrix(0, 19) = "총수량":          .ColWidth(19) = 0:      .ColAlignment(19) = flexAlignCenterCenter
            
            .TextMatrix(0, 20) = "DyeSchID"
            .TextMatrix(0, 21) = "DyeSchSeq"
            .TextMatrix(0, 22) = "CustomID"
            .TextMatrix(0, 23) = "ArticleID"
            .TextMatrix(0, 24) = "OrderID"
            .TextMatrix(0, 25) = "OrderSeq"
            .TextMatrix(0, 26) = "CardID"
            .TextMatrix(0, 27) = "SplitID"
            .TextMatrix(0, 28) = "WaitProcID"
            .TextMatrix(0, 29) = "CompProcID"
            
            For iRow = 1 To .Rows - 1
                .RowHeight(iRow) = 600
                .TextMatrix(iRow, 0) = Format((iRow \ 10) + 1, "00") & "호기"
                If (iRow Mod 10) < 3 Then
                    If (iRow Mod 10) = 0 Then
                        .RowHeight(iRow) = 40
                        .Cell(flexcpText, iRow, 0, iRow, .Cols - 1) = ""
                        .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &H50505
                    Else
                        .TextMatrix(iRow, 1) = "실적"
                        .Cell(flexcpBackColor, iRow, 1, iRow, .Cols - 1) = &HC0FFFF
                    End If
                Else
                    .TextMatrix(iRow, 1) = (iRow Mod 10) - 2
                End If
            Next iRow
            .MergeCells = flexMergeFixedOnly
            .MergeCol(0) = True
            
            
            .Redraw = flexRDDirect
        End With
    Next i
End Sub

Private Sub FillGridData()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
    Dim iCount%, i%, k%, iSeq%, iNowRow%
    Dim sWorkUnitID$
    Dim bToggle As Boolean
    Dim sCustom$, sArticle$
    
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
        
    For i = 0 To 3
        Set oRapid = New PlusLib2.CRapid
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
    
        Set rs = oRapid.GetRapidScheduling(i, 0)
        Set oRapid = Nothing

        If i = 0 Then
            sCustom = cboCustom.Text
            sArticle = cboArticle.Text
            
            Call DataToCombo(rs)
        End If

        bToggle = False
        With grdList(i)
            .Redraw = flexRDNone
            
            iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
            
            .Rows = .FixedRows
            For iCount = 1 To rs.RecordCount
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 300
                .Row = .Rows - 1
                .Col = 0
                If i < 3 Then
                    If rs!SchID > 0 Then
                        .CellChecked = flexNoCheckbox
                    Else
                        .CellChecked = flexUnchecked
                    End If
                Else
                    .CellChecked = flexUnchecked
                End If
                If iCount = 1 Then
                    sWorkUnitID = rs!WorkUnitId
                    iSeq = 0
                End If
                If sWorkUnitID <> rs!WorkUnitId Then
                    bToggle = Not (bToggle)
                    iSeq = 0
                End If
                .TextMatrix(.Rows - 1, 1) = rs!WorkUnitId
                .TextMatrix(.Rows - 1, 2) = rs!WorkUnitSeq
                .TextMatrix(.Rows - 1, 3) = "" & rs!BatJaNO
                .TextMatrix(.Rows - 1, 4) = CStr(iSeq + 1)
                .TextMatrix(.Rows - 1, 5) = Trim(rs!kCustom)
                .TextMatrix(.Rows - 1, 6) = Trim(rs!Article)
                .TextMatrix(.Rows - 1, 7) = Trim(rs!Color)
                .TextMatrix(.Rows - 1, 8) = MakeOrderID(rs!OrderID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 9) = Format(rs!CardID, "00-00-0000")
                .TextMatrix(.Rows - 1, 10) = rs!SplitID
                .TextMatrix(.Rows - 1, 11) = rs!WaitProc
                .TextMatrix(.Rows - 1, 12) = Format(rs!Roll, "#,##0")
                .TextMatrix(.Rows - 1, 13) = Format(rs!Qty, "#,###,##0")
                .TextMatrix(.Rows - 1, 14) = rs!CustomID
                .TextMatrix(.Rows - 1, 15) = rs!ArticleID
                .TextMatrix(.Rows - 1, 16) = rs!ColorID
                .TextMatrix(.Rows - 1, 17) = rs!CardID
                .TextMatrix(.Rows - 1, 18) = rs!SplitID
                .TextMatrix(.Rows - 1, 19) = rs!Custom
                .TextMatrix(.Rows - 1, 20) = rs!OrderID
                .TextMatrix(.Rows - 1, 21) = rs!OrderSeq
                .TextMatrix(.Rows - 1, 22) = rs!AfterProc
                .TextMatrix(.Rows - 1, 23) = rs!SchID
                .TextMatrix(.Rows - 1, 24) = rs!DyeSeq
                
                
                If bToggle = True Then
                    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HE0E0E0
                Else
                    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = 0
                End If
               
                sWorkUnitID = rs!WorkUnitId
                
                iSeq = iSeq + 1
                rs.MoveNext
            Next iCount
            rs.Close
            Set rs = Nothing
    
            If .Rows > .FixedRows Then
                .HighLight = flexHighlightAlways
                .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
                .TopRow = .Row
                .Col = .FixedCols
                .ColSel = .Cols - 1
            End If
            
            If i = 0 Then
                For k = 0 To cboCustom.ListCount - 1
                    If cboCustom.List(k) = sCustom Then
                        cboCustom.Text = sCustom
                        Call cboCustom_Click
                        Exit For
                    End If
                Next k
                
                For k = 0 To cboArticle.ListCount - 1
                    If cboArticle.List(k) = sArticle Then
                        cboArticle.Text = sArticle
                        Exit For
                    End If
                Next k
            End If
            
            .Redraw = flexRDDirect
        End With
    Next i

    Call GridReSetting(cboCustom, cboArticle)
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.FillGridData", Err.Description)
End Sub

'''Private Sub FillSchData()
'''    Dim oRapid As PlusLib2.CRapid
'''    Dim rs As Recordset
'''    Dim iCount%, i%, iSeq%
'''    Dim sWorkUnitID$        ' 스케쥴번호(9자리) + 차수(2자리)
'''    Dim sDyeSchIDSeq$
'''    Dim iCntRec%
'''    Dim iBaseRow%
'''    Dim iCurRow%
'''    Dim sMachID$
'''
'''    Screen.MousePointer = vbHourglass
'''
'''    On Error GoTo ErrHandler
'''
'''    Set oRapid = New PlusLib2.CRapid
'''    oRapid.Connection = g_adoCon
'''    oRapid.UserName = g_sUserName
'''
'''    Call InitGrdTab
'''
'''    Set rs = oRapid.GetRapidScheduledData()
'''    If rs.RecordCount > 0 Then
'''        With grdTab(0)
'''            For iCount = 1 To rs.RecordCount
'''                If rs!wimachid <> sMachID Then
'''                    iBaseRow = 2
'''                End If
'''                If iBaseRow < 9 Then
'''                    If sDyeSchIDSeq <> Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00") Then
'''                        iBaseRow = iBaseRow + 1
'''
'''                        iCurRow = ((CInt(rs!wimachid) - 1) * 10) + iBaseRow
'''                        If rs!CardSeq > 0 Then
'''                            .TextMatrix(iCurRow, 3) = Trim(rs!kCustom)
'''                            .TextMatrix(iCurRow, 4) = MakeOrderID(rs!OrderID, OM_COMPACT)
'''                            .TextMatrix(iCurRow, 5) = Trim(rs!Article)
'''                            .TextMatrix(iCurRow, 6) = Trim(rs!Color)
'''                            .TextMatrix(iCurRow, 7) = IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
'''                            .TextMatrix(iCurRow, 8) = Format(rs!Roll, "##,##0")
'''                            .TextMatrix(iCurRow, 9) = Format(rs!Qty, "##,##0")
'''                            .TextMatrix(iCurRow, 10) = rs!BatJaNO
'''                            .TextMatrix(iCurRow, 11) = IIf(rs!rseq > 0, "○", "Ｘ")
'''                            .TextMatrix(iCurRow, 12) = IIf(Trim(rs!instclss) = "", "Ｘ", "○")
''''                            .TextMatrix(iCurRow, 13) = Trim(Rs!WorkClss)
'''                            .TextMatrix(iCurRow, 14) = Trim(rs!RapidClss)
'''                            .TextMatrix(iCurRow, 15) = rs!Process       ' 대기공정
''''                            .TextMatrix(iCurRow, 16) = Trim(Rs!Remark)
'''                            .TextMatrix(iCurRow, 17) = ""       ' 완료공정
'''1
''''                            .TextMatrix(iCurRow, 20) = Format(Rs!SchID, "000000000") & Format(Rs!DyeSeq, "00")
'''                            .TextMatrix(iCurRow, 21) = ""
'''                            .TextMatrix(iCurRow, 22) = rs!CustomID
'''                            .TextMatrix(iCurRow, 23) = rs!ArticleID
'''                            .TextMatrix(iCurRow, 24) = rs!OrderID
'''                            .TextMatrix(iCurRow, 25) = rs!OrderSeq
'''                            .TextMatrix(iCurRow, 26) = rs!CardID
'''                            .TextMatrix(iCurRow, 27) = rs!SplitID
'''                            .TextMatrix(iCurRow, 28) = rs!waitprocid       ' 대기공정코드
'''                            .TextMatrix(iCurRow, 29) = ""       ' 완료공정코드
'''                        End If
'''                        .TextMatrix(iCurRow, 13) = Trim(rs!workclss)
'''                        .TextMatrix(iCurRow, 16) = Trim(rs!Remark)
'''                        .TextMatrix(iCurRow, 18) = Format(rs!wiroll, "##,##0")
'''                        .TextMatrix(iCurRow, 19) = Format(rs!wiqty, "##,##0")
'''                        .TextMatrix(iCurRow, 20) = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
'''
'''
'''                        If (Trim(rs!UseClss) = "작업" Or Len(Trim(rs!UseClss)) = 8) And Left(rs!waitprocid, 2) = "43" Then
'''                            .Cell(flexcpForeColor, iCurRow, 1, iCurRow, .Cols - 1) = vbBlue
'''                            .Cell(flexcpFontBold, iCurRow, 1, iCurRow, .Cols - 1) = True
'''                        End If
'''                        sMachID = rs!wimachid
'''                        sDyeSchIDSeq = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
'''                    Else
'''                        iCurRow = ((CInt(rs!wimachid) - 1) * 10) + iBaseRow
'''                        .TextMatrix(iCurRow, 3) = .TextMatrix(iCurRow, 3) & vbCrLf & Trim(rs!kCustom)
'''                        .TextMatrix(iCurRow, 4) = .TextMatrix(iCurRow, 4) & vbCrLf & MakeOrderID(rs!OrderID, OM_COMPACT)
'''                        .TextMatrix(iCurRow, 5) = .TextMatrix(iCurRow, 5) & vbCrLf & Trim(rs!Article)
'''                        .TextMatrix(iCurRow, 6) = .TextMatrix(iCurRow, 6) & vbCrLf & Trim(rs!Color)
'''                        .TextMatrix(iCurRow, 7) = .TextMatrix(iCurRow, 7) & vbCrLf & IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
'''                        .TextMatrix(iCurRow, 8) = .TextMatrix(iCurRow, 8) & vbCrLf & Format(rs!Roll, "##,##0")
'''                        .TextMatrix(iCurRow, 9) = .TextMatrix(iCurRow, 9) & vbCrLf & Format(rs!Qty, "##,##0")
'''                        .TextMatrix(iCurRow, 10) = .TextMatrix(iCurRow, 10) & vbCrLf & rs!BatJaNO
'''                        .TextMatrix(iCurRow, 15) = .TextMatrix(iCurRow, 14) & vbCrLf & rs!Process       ' 대기공정
'''                    End If
'''                End If
'''                rs.MoveNext
'''            Next iCount
'''        End With
'''    End If
'''    rs.Close
'''    Set rs = Nothing
'''
'''
'''    sDyeSchIDSeq = ""
'''    sMachID = ""
'''    iBaseRow = 0
'''
'''    Set rs = oRapid.GetRapidWorkedEachData()
'''    If rs.RecordCount > 0 Then
'''        With grdTab(0)
'''            For iCount = 1 To rs.RecordCount
'''                If rs!wkmachid <> sMachID Then
'''                    iBaseRow = 0
'''                End If
'''                If sDyeSchIDSeq <> Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00") Then
'''                    iBaseRow = iBaseRow + 1
'''                    iCurRow = ((CInt(rs!wkmachid) - 1) * 10) + iBaseRow
'''
'''                    .TextMatrix(iCurRow, 3) = Trim(rs!kCustom)
'''                    .TextMatrix(iCurRow, 4) = MakeOrderID(rs!OrderID, OM_COMPACT)
'''                    .TextMatrix(iCurRow, 5) = Trim(rs!Article)
'''                    .TextMatrix(iCurRow, 6) = Trim(rs!Color)
'''                    .TextMatrix(iCurRow, 7) = IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
'''                    .TextMatrix(iCurRow, 8) = Format(rs!Roll, "##,##0")
'''                    .TextMatrix(iCurRow, 9) = Format(rs!Qty, "##,##0")
'''                    .TextMatrix(iCurRow, 10) = rs!BatJaNO
'''                    .TextMatrix(iCurRow, 11) = "○"
'''                    .TextMatrix(iCurRow, 12) = "○"
'''                    .TextMatrix(iCurRow, 13) = rs!workclss
'''                    .TextMatrix(iCurRow, 14) = rs!RapidClss
'''                    .TextMatrix(iCurRow, 15) = rs!Process       ' 대기공정
'''                    .TextMatrix(iCurRow, 16) = rs!Remark
'''                    .TextMatrix(iCurRow, 17) = ""       ' 완료공정
'''                    .TextMatrix(iCurRow, 18) = Format(rs!wkRoll, "##,##0")
'''                    .TextMatrix(iCurRow, 19) = Format(rs!wkqty, "##,##0")
'''                    .TextMatrix(iCurRow, 20) = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
'''                    .TextMatrix(iCurRow, 21) = ""
'''                    .TextMatrix(iCurRow, 22) = rs!CustomID
'''                    .TextMatrix(iCurRow, 23) = rs!ArticleID
'''                    .TextMatrix(iCurRow, 24) = rs!OrderID
'''                    .TextMatrix(iCurRow, 25) = rs!OrderSeq
'''                    .TextMatrix(iCurRow, 26) = rs!CardID
'''                    .TextMatrix(iCurRow, 27) = rs!SplitID
'''                    .TextMatrix(iCurRow, 28) = rs!waitprocid       ' 대기공정코드
'''                    .TextMatrix(iCurRow, 29) = ""       ' 완료공정코드
'''                    sMachID = rs!wkmachid
'''                    sDyeSchIDSeq = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
'''                Else
'''                    iCurRow = ((CInt(rs!wkmachid) - 1) * 10) + iBaseRow
'''                    .TextMatrix(iCurRow, 3) = .TextMatrix(iCurRow, 3) & vbCrLf & Trim(rs!kCustom)
'''                    .TextMatrix(iCurRow, 4) = .TextMatrix(iCurRow, 4) & vbCrLf & MakeOrderID(rs!OrderID, OM_COMPACT)
'''                    .TextMatrix(iCurRow, 5) = .TextMatrix(iCurRow, 5) & vbCrLf & Trim(rs!Article)
'''                    .TextMatrix(iCurRow, 6) = .TextMatrix(iCurRow, 6) & vbCrLf & Trim(rs!Color)
'''                    .TextMatrix(iCurRow, 7) = .TextMatrix(iCurRow, 7) & vbCrLf & IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
'''                    .TextMatrix(iCurRow, 8) = .TextMatrix(iCurRow, 8) & vbCrLf & Format(rs!Roll, "##,##0")
'''                    .TextMatrix(iCurRow, 9) = .TextMatrix(iCurRow, 9) & vbCrLf & Format(rs!Qty, "##,##0")
'''                    .TextMatrix(iCurRow, 10) = .TextMatrix(iCurRow, 10) & vbCrLf & rs!BatJaNO
'''                    .TextMatrix(iCurRow, 15) = .TextMatrix(iCurRow, 14) & vbCrLf & rs!Process       ' 대기공정
'''                End If
'''                rs.MoveNext
'''            Next iCount
'''        End With
'''    End If
'''
'''    rs.Close
'''    Set rs = Nothing
'''
'''    Set oRapid = Nothing
'''
'''    Call optMachine_Click(nMachineID)
'''
'''    Screen.MousePointer = vbDefault
'''
'''    Exit Sub
'''
'''ErrHandler:
'''    Screen.MousePointer = vbDefault
'''
'''    Set rs = Nothing
'''    Set oRapid = Nothing
'''
'''    Call ErrorBox(Err.Number, "frmInstRapid.FillSchData", Err.Description)
'''End Sub
Private Sub FillSchData()
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
    Dim iCount%, i%, iSeq%
    Dim sWorkUnitID$        ' 스케쥴번호(9자리) + 차수(2자리)
    Dim sDyeSchIDSeq$
    Dim iCntRec%
    Dim iBaseRow%
    Dim iCurRow%
    Dim sMachID$
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
    
    Call InitGrdTab

    Set rs = oRapid.GetRapidScheduledData()
    If rs.RecordCount > 0 Then
        With grdTab(0)
            For iCount = 1 To rs.RecordCount
                If rs!wimachid <> sMachID Then
                    iBaseRow = 2
                End If
                If iBaseRow < 9 Then
                    If sDyeSchIDSeq <> Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00") Then
                        iBaseRow = iBaseRow + 1
                        
                        iCurRow = ((CInt(rs!wimachid) - 1) * 10) + iBaseRow
                        If rs!CardSeq > 0 Then
                            .TextMatrix(iCurRow, 3) = Trim(rs!kCustom)
                            .TextMatrix(iCurRow, 4) = MakeOrderID(rs!OrderID, OM_COMPACT)
                            .TextMatrix(iCurRow, 5) = Trim(rs!Article)
                            .TextMatrix(iCurRow, 6) = Trim(rs!Color)
                            .TextMatrix(iCurRow, 7) = IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
                            .TextMatrix(iCurRow, 8) = Format(rs!Roll, "##,##0")
                            .TextMatrix(iCurRow, 9) = Format(rs!Qty, "##,##0")
                            .TextMatrix(iCurRow, 10) = rs!BatJaNO
                            .TextMatrix(iCurRow, 11) = IIf(rs!rseq > 0, "○", "Ｘ")
                            .TextMatrix(iCurRow, 12) = IIf(Trim(rs!instclss) = "", "Ｘ", "○")
'                            .TextMatrix(iCurRow, 13) = Trim(Rs!WorkClss)
                            .TextMatrix(iCurRow, 14) = Trim(rs!RapidClss)
                            .TextMatrix(iCurRow, 15) = rs!Process       ' 대기공정
'                            .TextMatrix(iCurRow, 16) = Trim(Rs!Remark)
                            .TextMatrix(iCurRow, 17) = ""       ' 완료공정
1
'                            .TextMatrix(iCurRow, 20) = Format(Rs!SchID, "000000000") & Format(Rs!DyeSeq, "00")
                            .TextMatrix(iCurRow, 21) = ""
                            .TextMatrix(iCurRow, 22) = rs!CustomID
                            .TextMatrix(iCurRow, 23) = rs!ArticleID
                            .TextMatrix(iCurRow, 24) = rs!OrderID
                            .TextMatrix(iCurRow, 25) = rs!OrderSeq
                            .TextMatrix(iCurRow, 26) = rs!CardID
                            .TextMatrix(iCurRow, 27) = rs!SplitID
                            .TextMatrix(iCurRow, 28) = rs!waitprocid       ' 대기공정코드
                            .TextMatrix(iCurRow, 29) = ""       ' 완료공정코드
                        End If
                        .TextMatrix(iCurRow, 13) = Trim(rs!workclss)
                        .TextMatrix(iCurRow, 16) = Trim(rs!Remark)
                        .TextMatrix(iCurRow, 18) = Format(rs!wiroll, "##,##0")
                        .TextMatrix(iCurRow, 19) = Format(rs!wiqty, "##,##0")
                        .TextMatrix(iCurRow, 20) = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
                        
                        
                        If (Trim(rs!UseClss) = "작업" Or Len(Trim(rs!UseClss)) = 8) And Left(rs!waitprocid, 2) = "43" Then
                            .Cell(flexcpForeColor, iCurRow, 1, iCurRow, .Cols - 1) = vbBlue
                            .Cell(flexcpFontBold, iCurRow, 1, iCurRow, .Cols - 1) = True
                        End If
                        sMachID = rs!wimachid
                        sDyeSchIDSeq = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
                    Else
                        iCurRow = ((CInt(rs!wimachid) - 1) * 10) + iBaseRow
                        .TextMatrix(iCurRow, 3) = .TextMatrix(iCurRow, 3) & vbCrLf & Trim(rs!kCustom)
                        .TextMatrix(iCurRow, 4) = .TextMatrix(iCurRow, 4) & vbCrLf & MakeOrderID(rs!OrderID, OM_COMPACT)
                        .TextMatrix(iCurRow, 5) = .TextMatrix(iCurRow, 5) & vbCrLf & Trim(rs!Article)
                        .TextMatrix(iCurRow, 6) = .TextMatrix(iCurRow, 6) & vbCrLf & Trim(rs!Color)
                        .TextMatrix(iCurRow, 7) = .TextMatrix(iCurRow, 7) & vbCrLf & IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
                        .TextMatrix(iCurRow, 8) = .TextMatrix(iCurRow, 8) & vbCrLf & Format(rs!Roll, "##,##0")
                        .TextMatrix(iCurRow, 9) = .TextMatrix(iCurRow, 9) & vbCrLf & Format(rs!Qty, "##,##0")
                        .TextMatrix(iCurRow, 10) = .TextMatrix(iCurRow, 10) & vbCrLf & rs!BatJaNO
                        .TextMatrix(iCurRow, 15) = .TextMatrix(iCurRow, 14) & vbCrLf & rs!Process       ' 대기공정
                    End If
                End If
                rs.MoveNext
            Next iCount
        End With
    End If
    rs.Close
    Set rs = Nothing


    sDyeSchIDSeq = ""
    sMachID = ""
    iBaseRow = 0

    Set rs = oRapid.GetRapidWorkedEachData()
    If rs.RecordCount > 0 Then
        With grdTab(0)
            For iCount = 1 To rs.RecordCount
                If rs!wkmachid <> sMachID Then
                    iBaseRow = 0
                End If
                If sDyeSchIDSeq <> Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00") Then
                    iBaseRow = iBaseRow + 1
                    iCurRow = ((CInt(rs!wkmachid) - 1) * 10) + iBaseRow
                    
                    .TextMatrix(iCurRow, 3) = Trim(rs!kCustom)
                    If Trim(rs!OrderID) = "" Then
                        .TextMatrix(iCurRow, 4) = ""
                    Else
                        .TextMatrix(iCurRow, 4) = MakeOrderID(rs!OrderID, OM_COMPACT)
                    End If
                    .TextMatrix(iCurRow, 5) = Trim(rs!Article)
                    .TextMatrix(iCurRow, 6) = Trim(rs!Color)
                    If Trim(rs!CardID) = "" Then
                        .TextMatrix(iCurRow, 7) = ""
                    Else
                        .TextMatrix(iCurRow, 7) = IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
                    End If
                    .TextMatrix(iCurRow, 8) = Format(rs!Roll, "##,##0")
                    .TextMatrix(iCurRow, 9) = Format(rs!Qty, "##,##0")
                    .TextMatrix(iCurRow, 10) = rs!BatJaNO
                    .TextMatrix(iCurRow, 11) = "○"
                    .TextMatrix(iCurRow, 12) = "○"
                    .TextMatrix(iCurRow, 13) = rs!workclss
                    .TextMatrix(iCurRow, 14) = rs!RapidClss
                    .TextMatrix(iCurRow, 15) = rs!Process       ' 대기공정
                    .TextMatrix(iCurRow, 16) = rs!Remark
                    .TextMatrix(iCurRow, 17) = ""       ' 완료공정
                    .TextMatrix(iCurRow, 18) = Format(rs!wkRoll, "##,##0")
                    .TextMatrix(iCurRow, 19) = Format(rs!wkqty, "##,##0")
                    .TextMatrix(iCurRow, 20) = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
                    .TextMatrix(iCurRow, 21) = ""
                    .TextMatrix(iCurRow, 22) = rs!CustomID
                    .TextMatrix(iCurRow, 23) = rs!ArticleID
                    .TextMatrix(iCurRow, 24) = rs!OrderID
                    .TextMatrix(iCurRow, 25) = rs!OrderSeq
                    .TextMatrix(iCurRow, 26) = rs!CardID
                    .TextMatrix(iCurRow, 27) = rs!SplitID
                    .TextMatrix(iCurRow, 28) = rs!waitprocid       ' 대기공정코드
                    .TextMatrix(iCurRow, 29) = ""       ' 완료공정코드
                    sMachID = rs!wkmachid
                    sDyeSchIDSeq = Format(rs!SchID, "000000000") & Format(rs!DyeSeq, "00")
                Else
                    iCurRow = ((CInt(rs!wkmachid) - 1) * 10) + iBaseRow
                    .TextMatrix(iCurRow, 3) = .TextMatrix(iCurRow, 3) & vbCrLf & Trim(rs!kCustom)
                    .TextMatrix(iCurRow, 4) = .TextMatrix(iCurRow, 4) & vbCrLf & MakeOrderID(rs!OrderID, OM_COMPACT)
                    .TextMatrix(iCurRow, 5) = .TextMatrix(iCurRow, 5) & vbCrLf & Trim(rs!Article)
                    .TextMatrix(iCurRow, 6) = .TextMatrix(iCurRow, 6) & vbCrLf & Trim(rs!Color)
                    .TextMatrix(iCurRow, 7) = .TextMatrix(iCurRow, 7) & vbCrLf & IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_COMPACT), MakeCardID(rs!CardID, OM_COMPACT) & "(" & rs!SplitID & ")")
                    .TextMatrix(iCurRow, 8) = .TextMatrix(iCurRow, 8) & vbCrLf & Format(rs!Roll, "##,##0")
                    .TextMatrix(iCurRow, 9) = .TextMatrix(iCurRow, 9) & vbCrLf & Format(rs!Qty, "##,##0")
                    .TextMatrix(iCurRow, 10) = .TextMatrix(iCurRow, 10) & vbCrLf & rs!BatJaNO
                    .TextMatrix(iCurRow, 15) = .TextMatrix(iCurRow, 14) & vbCrLf & rs!Process       ' 대기공정
                End If
                rs.MoveNext
            Next iCount
        End With
    End If

    rs.Close
    Set rs = Nothing

    Set oRapid = Nothing

    Call optMachine_Click(nMachineID)
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.FillSchData", Err.Description)
End Sub

Private Sub grdList_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Index = 4 Then
        Cancel = True
    Else
        If Col = 0 Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub grdList_Click(Index As Integer)
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
    Dim iCount%

    cboColor.Clear
    
    If Index = 4 Then
        If grdList(4).Row >= grdList(4).FixedRows Then
            pnlCardID.Caption = Trim(grdList(4).TextMatrix(grdList(4).Row, 17))
            pnlSplitID.Caption = Trim(grdList(4).TextMatrix(grdList(4).Row, 18))
            lblOrderID = Trim(grdList(4).TextMatrix(grdList(4).Row, 20))
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
            Set rs = oRapid.GetOrderColorList(pnlCardID, pnlSplitID)
            If rs.RecordCount > 0 Then
                For iCount = 1 To rs.RecordCount
                    cboColor.AddItem rs!Color
                    cboColor.ItemData(cboColor.NewIndex) = CLng(rs!OrderSeq)
                    rs.MoveNext
                Next iCount
                cboColor.ListIndex = FindComboBox(cboColor, CLng(grdList(4).TextMatrix(grdList(4).Row, 21)))
            End If
            Set rs = Nothing
            Set oRapid = Nothing
        Else
            pnlCardID.Caption = "카드번호"
            pnlSplitID.Caption = "분할"
            lblOrderID = ""
        End If
    Else
        If grdList(SSTab1.Tab).Row >= grdList(SSTab1.Tab).FixedRows Then
            pnlCardID.Caption = Trim(grdList(SSTab1.Tab).TextMatrix(grdList(SSTab1.Tab).Row, 17))
            pnlSplitID.Caption = Trim(grdList(SSTab1.Tab).TextMatrix(grdList(SSTab1.Tab).Row, 18))
            lblOrderID = Trim(grdList(SSTab1.Tab).TextMatrix(grdList(SSTab1.Tab).Row, 20))
            
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
            Set rs = oRapid.GetOrderColorList(pnlCardID, pnlSplitID)
            If rs.RecordCount > 0 Then
                For iCount = 1 To rs.RecordCount
                    cboColor.AddItem rs!Color
                    cboColor.ItemData(cboColor.NewIndex) = CLng(rs!OrderSeq)
                    rs.MoveNext
                Next iCount
                cboColor.ListIndex = FindComboBox(cboColor, CLng(grdList(SSTab1.Tab).TextMatrix(grdList(SSTab1.Tab).Row, 21)))
            End If
            Set rs = Nothing
            Set oRapid = Nothing
        Else
            pnlCardID.Caption = "카드번호"
            pnlSplitID.Caption = "분할"
            lblOrderID = ""
        End If
    End If
End Sub

Private Sub grdList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Index = 4 Then
        With grdList(Index)
            If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
                .ToolTipText = .TextMatrix(.MouseRow, .Cols - 3)
            End If
        End With
'    End If
End Sub

Private Sub grdList_RowColChange(Index As Integer)
'    Call grdList_Click(Index)
End Sub

Private Sub grdTab_Click(Index As Integer)
Dim iCol%, iCurrRow%

    fraWorkEnd.Visible = False
    With grdTab(Index)
        If .Row >= .FixedRows And .TextMatrix(.Row, 1) <> "실적" And .TextMatrix(.Row, 20) <> "" Then
            pnlCardID.Caption = Trim(.TextMatrix(.Row, 26))
            pnlSplitID.Caption = Trim(.TextMatrix(.Row, 27))
            .Col = 1
            shpBox.Left = .CellLeft
            shpBox.Top = .CellTop
            .Col = 12
            shpBox.Width = 14100
            shpBox.Height = .RowHeight(.Row)
                
            If Trim(.TextMatrix(.Row, 20)) <> "" Then
                If (.Row Mod 10) > 3 Then
                    Call VisibleUpDownFrame(True)
                    Call VisibleWorkFrame(False)
                    shpBox.Visible = True
                    
                    If .Cell(flexcpForeColor, .Row, 3) = vbBlue Then
                        Call ToggleShapeBox(True, True)
                        Call VisibleUpDownFrame(False)
                    End If
                    
                    Exit Sub
                End If
                If .Cell(flexcpForeColor, .Row, 3) = vbBlue Then
                    Call ToggleShapeBox(True, True)
                    Call VisibleUpDownFrame(False)
                Else
                    Call ToggleShapeBox(True, False)
                End If
                Call VisibleWorkFrame(bEnableWork)
                
            Else
                Call ToggleShapeBox(False, False)
            End If
        Else
            If .TextMatrix(.Row, 1) = "실적" Then
                pnlCardID.Caption = Trim(.TextMatrix(.Row, 26))
                pnlSplitID.Caption = Trim(.TextMatrix(.Row, 27))
                lblOrderID = .TextMatrix(.Row, 24)
            Else
                pnlCardID.Caption = ""
                pnlSplitID.Caption = ""
                lblOrderID = ""
            End If
            Call ToggleShapeBox(False, False)
        End If
    End With
End Sub

Private Sub ToggleShapeBox(bFlag As Boolean, bWorking As Boolean)
    shpBox.Visible = bFlag
    shpButton.Visible = bFlag
    cmdWorkStart.Visible = bFlag
    cmdWorkEnd.Visible = bFlag
    cmdCancelStart.Visible = bFlag
    cmdAddWork.Visible = bFlag
    
    cmdWorkStart.Enabled = Not (bWorking)
    cmdWorkEnd.Enabled = bWorking
    cmdCancelStart.Enabled = bWorking
    cmdAddWork.Enabled = bWorking
    
    shpBox.Visible = bFlag
    fraUpDown.Visible = bFlag
    
End Sub

Private Sub VisibleWorkButton(bFlag As Boolean)
    shpButton.Visible = bFlag
    cmdWorkStart.Visible = bFlag
    cmdWorkEnd.Visible = bFlag
    cmdCancelStart.Visible = bFlag
    cmdAddWork.Visible = bFlag
End Sub


Private Sub VisibleWorkFrame(bFlag As Boolean)
    shpButton.Visible = bFlag
    cmdWorkStart.Visible = bFlag
    cmdWorkEnd.Visible = bFlag
    cmdCancelStart.Visible = bFlag
    cmdAddWork.Visible = bFlag
End Sub

Private Sub VisibleUpDownFrame(bFlag As Boolean)
    fraUpDown.Visible = bFlag
End Sub

Private Sub grdTab_DblClick(Index As Integer)
    Dim iCol%, i%
    Dim oRapid As PlusLib2.CRapid
    Dim rs As Recordset
    Dim sRs As Recordset
    Dim iCount%, iSeq%
    Dim lTotRoll As Long, lTotQty As Long
    Dim iRapidSeq As Integer

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    fraWorkEnd.Visible = False
    With grdTab(Index)
        If .Row >= .FixedRows And .TextMatrix(.Row, 1) <> "실적" And .TextMatrix(.Row, 20) <> "" Then
            lblWork = ""
            If Trim(.TextMatrix(.Row, 20)) <> "" Then
                Call ToggleShapeBox(False, False)
                lblMachID = Left(.TextMatrix(.Row, 0), 2)
            
                Set oRapid = New PlusLib2.CRapid
                oRapid.Connection = g_adoCon
                oRapid.UserName = g_sUserName
                
                Set rs = oRapid.GetCheckDyeWorking(CLng(Left(.TextMatrix(.Row, 20), 9)), CInt(Right(.TextMatrix(.Row, 20), 2)))
                
                If rs.RecordCount > 0 Then
                    If (Trim(rs!UseClss) = "작업" Or Len(Trim(rs!UseClss)) = 8) And Left(rs!procid, 2) = "43" Then
                        MsgBox "선택되어진 건은 현재 작업중입니다" & vbCrLf & "편집이 불가능합니다", vbExclamation, "편집 불가"
                        If Not (.Cell(flexcpForeColor, .Row, 2) = vbBlue) Then
                            Set rs = Nothing
                            Set oRapid = Nothing
                            Call FillSchData
                            Exit Sub
                        End If
                        cmdDelete.Enabled = False
                        cmdConfirm.Enabled = False
                    Else
                        cmdDelete.Enabled = True
                        cmdConfirm.Enabled = True
                        cmdDelete.Visible = True
                        cmdConfirm.Caption = "저장"
                    End If
                    cmdToggle.Caption = "염색 스케쥴 조회"
                    pnlMsg.Caption = "편집 중 입니다...."
                    MoveScreen (True)
                    cmdScreen.Caption = "편집취소"
                    pnlView.Visible = False
                    pnlEdit.Visible = True
                    lblSchIDSeq = .TextMatrix(.Row, 20)
                    
                    .TopRow = .Row
                    .Cell(flexcpFontBold, .Row, 1, .Row, .Cols - 1) = True
                    grdList(4).Rows = grdList(4).FixedRows
                    If Len(Trim(rs!UseClss)) = 8 Or Len(Trim(rs!UseClss)) = 0 Then
                        Set sRs = oRapid.GetRapidSchedulingBox(CLng(Left(.TextMatrix(.Row, 20), 9)), CInt(Right(.TextMatrix(.Row, 20), 2)))
                    Else
                        Set sRs = oRapid.GetRapidScheduling(0, CLng(Left(.TextMatrix(.Row, 20), 9)))
                    End If
                    If sRs.RecordCount > 0 Then
                        For iCount = 1 To sRs.RecordCount
                            If iCount = 1 Then
                                lstArray(0).ListIndex = -1
                                For i = 0 To lstArray(0).ListCount - 1
                                    If Left(lstArray(0).List(i), 2) = Format(sRs!wimachid, "00") Then
                                        lstArray(0).Selected(i) = True
                                        Exit For
                                    End If
                                    
                                Next i
                                For i = 0 To lstArray(1).ListCount - 1
                                    If Left(lstArray(1).List(i), 2) = Format(sRs!PatternID, "00") Then
                                        lstArray(1).Selected(i) = True
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To lstArray(2).ListCount - 1
                                    If lstArray(2).List(i) = sRs!RapidClss Then
                                        lstArray(2).Selected(i) = True
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To lstArray(3).ListCount - 1
                                    If Right(lstArray(3).List(i), 8) = Format(sRs!PersonID, "00000000") Then
                                        lstArray(3).Selected(i) = True
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To lstArray(4).ListCount - 1
                                    If Left(lstArray(4).List(i), 2) = Format(sRs!wirapidseq, "00") Then
                                        lstArray(4).Selected(i) = True
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To lstArray(5).ListCount - 1
                                    If lstArray(5).List(i) = sRs!workclss Then
                                        lstArray(5).Selected(i) = True
                                        Exit For
                                    End If
                                Next i
                                
                                txtRemark = sRs!Remark
                            End If
                            If Len(Trim(rs!UseClss)) = 8 Or Len(Trim(rs!UseClss)) = 0 Then
                                Exit For
                            End If
                            
                            iSeq = iSeq + 1
                            lTotRoll = lTotRoll + CLng(sRs!Roll)
                            lTotQty = lTotQty + CLng(sRs!Qty)
                        
                            grdList(4).Rows = grdList(4).Rows + 1
                            grdList(4).RowHeight(grdList(4).Rows - 1) = 300
'                                grdList(4).TextMatrix(grdList(4).Rows - 1, 0) = ""
                            grdList(4).Cell(flexcpChecked, grdList(4).Rows - 1, 0, grdList(4).Rows - 1, 0) = flexChecked
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 1) = sRs!WorkUnitId
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 2) = sRs!WorkUnitSeq
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 3) = "" & sRs!BatJaNO
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 4) = CStr(iSeq)
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 5) = Trim(sRs!kCustom)
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 6) = Trim(sRs!Article)
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 7) = Trim(sRs!Color)
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 8) = MakeOrderID(sRs!OrderID, OM_EXPAND)
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 9) = Format(sRs!CardID, "00-00-0000")
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 10) = sRs!SplitID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 11) = sRs!WaitProc
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 12) = Format(sRs!Roll, "#,##0")
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 13) = Format(sRs!Qty, "#,###,##0")
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 14) = sRs!CustomID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 15) = sRs!ArticleID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 16) = sRs!ColorID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 17) = sRs!CardID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 18) = sRs!SplitID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 19) = sRs!waitprocid
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 20) = sRs!OrderID
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 21) = sRs!OrderSeq
                            grdList(4).TextMatrix(grdList(4).Rows - 1, 22) = sRs!AfterProc
                            sRs.MoveNext
                        Next iCount
                        grdList(4).Rows = grdList(4).Rows + 1
                        grdList(4).RowHeight(grdList(4).Rows - 1) = 300
                        grdList(4).Cell(flexcpText, grdList(4).Rows - 1, 0, grdList(4).Rows - 1, 11) = "선택되어진 카드 총 합계"
                        grdList(4).Cell(flexcpFontBold, grdList(4).Rows - 1, 0, grdList(4).Rows - 1, grdList(4).Cols - 1) = True
                        grdList(4).TextMatrix(grdList(4).Rows - 1, 12) = Format(lTotRoll, "#,##0")
                        grdList(4).TextMatrix(grdList(4).Rows - 1, 13) = Format(lTotQty, "#,###,##0")
                        grdList(4).MergeCells = flexMergeRestrictRows
                        grdList(4).MergeRow(grdList(4).Rows - 1) = True
                    
                    End If
                    sRs.Close
                    Set sRs = Nothing
                Else
                    MsgBox "선택된 건은 현재 작업이 완료되었습니다", vbOKOnly, "작업완료 건"
                End If
                rs.Close
                Set rs = Nothing
                Set oRapid = Nothing
            Else
                Call ToggleShapeBox(False, False)
            End If
            cmdExpandWait.Visible = False
            Call FoldOrUnFold(False)
        End If
    End With
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.grdTab_DblClick", Err.Description)
End Sub

Private Sub grdTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdTab(Index)
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            If Trim(.TextMatrix(.MouseRow, 20)) <> "" Then
                .ToolTipText = "절수: " & .TextMatrix(.MouseRow, 18) & ",  수량: " & .TextMatrix(.MouseRow, 19) & ",  " & .TextMatrix(.MouseRow, 16)
            Else
                .ToolTipText = ""
            End If
        End If
    End With
End Sub

Private Sub lstArray_Click(Index As Integer)
    Dim oRapid As PlusLib2.CRapid
    Dim oPerson  As PlusLib2.CPerson
    Dim rs As Recordset
    Dim iCount%

    If Index = 0 Then
        If Trim(lstArray(0).Text) <> "" Then
            Set oRapid = New PlusLib2.CRapid
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
            
            Set rs = oRapid.GetDyePatternList(1, CInt(Left(lstArray(0).Text, 2)), 0)
            
            Set oRapid = Nothing
            
            lstArray(1).Clear
            For iCount = 1 To rs.RecordCount
                lstArray(1).AddItem Format(rs!PtNo, "00") & ". " & rs!PtName
    '            lstArray(6).AddItem Format(rs!PtNo, "000") & " " & rs!PtName
                rs.MoveNext
            Next iCount
            rs.Close
            Set rs = Nothing
        End If
    End If
End Sub

Private Sub lstArray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 2, 8:
            lstArray(5).Selected(0) = True
            lstArray(7).Selected(0) = True
        Case 5, 7:
            lstArray(2).ListIndex = -1
            lstArray(8).ListIndex = -1
    End Select

End Sub


Private Sub optMachine_Click(Index As Integer)
    Call ToggleShapeBox(False, False)
    
    nMachineID = Index
    With grdTab(0)
        .TopRow = Index * 10 + 1
    End With
End Sub

Private Sub cboCustom_Click()
    Dim k, j As Integer
    Dim bCheck As Boolean
    
    cboArticle.Clear
    cboArticle.AddItem "전 체"
    
    With grdList(0)
        For k = .FixedRows To .Rows - 1
            If .RowHeight(k) > 0 Then
                bCheck = False
                If .TextMatrix(k, 5) = cboCustom.Text Then
                    For j = 0 To cboArticle.ListCount
                        If cboArticle.List(j) = .TextMatrix(k, 6) Then
                            bCheck = True
                            Exit For
                        End If
                    Next j
                    If Not bCheck Then
                        cboArticle.AddItem .TextMatrix(k, 6)
                    End If
                End If
            End If
        Next k
    End With
    
    cboArticle.ListIndex = 0
End Sub

Private Sub DataToCombo(pRs As Recordset)
    Dim k, j, l As Integer
    Dim bCheck As Boolean
    
    cboArticle.Clear
    cboCustom.Clear
    
    cboArticle.AddItem "전 체"
    cboCustom.AddItem "전 체"
    
    For k = 1 To pRs.RecordCount
        bCheck = False
        If Not pRs!SchID > 0 Then
            For j = 0 To cboCustom.ListCount
                If cboCustom.List(j) = pRs!kCustom Then
                    bCheck = True
                    Exit For
                End If
            Next j
            If Not bCheck Then
                cboCustom.AddItem pRs!kCustom
            End If
        
'            bCheck = False
'            For j = 0 To cboArticle.ListCount
'                If cboArticle.List(j) = pRs!Article Then
'                    bCheck = True
'                    Exit For
'                End If
'            Next j
'            If Not bCheck Then
'                cboArticle.AddItem pRs!Article
'            End If
        End If
        
        pRs.MoveNext
    Next k
    pRs.MoveFirst
    
    cboArticle.ListIndex = 0
    cboCustom.ListIndex = 0
End Sub

Private Sub GridReSetting(ByVal pCustom As String, ByVal pArticle As String)
    Dim i, j As Integer
    
'    For j = grdList.LBound To grdList.UBound
'        With grdList(j)
        With grdList(0)
            .Redraw = flexRDNone
        
            For i = .FixedRows To .Rows - 1
                If pCustom = "전 체" And pArticle = "전 체" Then
                    .RowHidden(i) = False
                ElseIf pCustom = "전 체" And pArticle <> "전 체" Then
                    If .TextMatrix(i, 6) = pArticle Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                ElseIf pCustom <> "전 체" And pArticle = "전 체" Then
                    If .TextMatrix(i, 5) = pCustom Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                Else
                    If .TextMatrix(i, 5) = pCustom And .TextMatrix(i, 6) = pArticle Then
                        .RowHidden(i) = False
                    Else
                        .RowHidden(i) = True
                    End If
                End If
            
            Next i
        
            .Redraw = flexRDDirect
        End With
'    Next j
End Sub

