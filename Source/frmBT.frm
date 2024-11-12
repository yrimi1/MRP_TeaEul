VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBT 
   ClientHeight    =   9255
   ClientLeft      =   2145
   ClientTop       =   1830
   ClientWidth     =   11865
   Icon            =   "frmBT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   9195
      TabIndex        =   81
      Top             =   135
      Width           =   1575
   End
   Begin Threed.SSPanel pnlSend 
      Height          =   2265
      Left            =   2295
      TabIndex        =   68
      Top             =   6270
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   3995
      _Version        =   196609
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel5 
         Height          =   1725
         Left            =   75
         TabIndex        =   75
         Top             =   420
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3043
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtSendPerson 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   90
            TabIndex        =   77
            Top             =   510
            Width           =   2295
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   330
            Left            =   90
            TabIndex        =   76
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            _Version        =   196609
            Caption         =   "담 당 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   7
            Left            =   2415
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   510
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   330
            Left            =   90
            TabIndex        =   79
            Top             =   915
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            _Version        =   196609
            Caption         =   "발송일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpSend 
            Height          =   330
            Left            =   75
            TabIndex        =   80
            Top             =   1305
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   582
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
            Format          =   54657024
            CurrentDate     =   37112
         End
      End
      Begin Threed.SSCommand cmdNO 
         Height          =   825
         Index           =   0
         Left            =   4515
         TabIndex        =   73
         Top             =   1305
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1455
         _Version        =   196609
         Caption         =   "닫기"
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   1725
         Left            =   3015
         TabIndex        =   72
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   3043
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "등록"
         AutoSize        =   1
         PictureAlignment=   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   345
         Left            =   15
         TabIndex        =   69
         Top             =   15
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   609
         _Version        =   196609
         ForeColor       =   -2147483634
         BackColor       =   8388608
         Caption         =   "  담당자 및 날짜를 입력하십시오."
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdNO 
            Height          =   300
            Index           =   1
            Left            =   5160
            TabIndex        =   70
            Top             =   30
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   529
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
            Caption         =   "X"
         End
         Begin Threed.SSCommand cmdCancel 
            Height          =   240
            Index           =   2
            Left            =   7050
            TabIndex        =   71
            Top             =   45
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
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
            Caption         =   "X"
         End
      End
      Begin Threed.SSCommand cmdUnSend 
         Height          =   825
         Left            =   4515
         TabIndex        =   74
         Top             =   420
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1455
         _Version        =   196609
         Enabled         =   0   'False
         Caption         =   "등록취소"
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   45
      TabIndex        =   65
      Top             =   8610
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   " 색상 세부내역 "
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   66
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "보임"
         Value           =   -1
      End
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   1
         Left            =   1035
         TabIndex        =   67
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "숨김"
      End
   End
   Begin Threed.SSCommand cmdRework 
      Height          =   690
      Left            =   3225
      TabIndex        =   43
      Top             =   8535
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      재등록"
      PictureAlignment=   1
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   2
      Left            =   5445
      TabIndex        =   10
      Top             =   840
      Width           =   1635
   End
   Begin Threed.SSPanel pnlModify 
      Height          =   6135
      Left            =   2130
      TabIndex        =   21
      Top             =   1260
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10821
      _Version        =   196609
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlText 
         Height          =   345
         Left            =   75
         TabIndex        =   33
         Top             =   510
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
         _Version        =   196609
         BackColor       =   16777215
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   630
         Index           =   0
         Left            =   6015
         TabIndex        =   64
         Top             =   5385
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1111
         _Version        =   196609
         Caption         =   "    닫기(&X)"
         PictureAlignment=   1
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Left            =   45
         TabIndex        =   31
         Top             =   45
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   582
         _Version        =   196609
         BackColor       =   8388608
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdCancel 
            Height          =   285
            Index           =   1
            Left            =   7020
            TabIndex        =   32
            Top             =   30
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
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
            Caption         =   "X"
         End
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   630
         Left            =   4575
         TabIndex        =   63
         Top             =   5400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1111
         _Version        =   196609
         Caption         =   "      저장(&S)"
         PictureAlignment=   1
      End
      Begin Threed.SSPanel pnlInfo 
         Height          =   4365
         Left            =   75
         TabIndex        =   22
         Top             =   930
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   7699
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtSendPer 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1350
            TabIndex        =   58
            Top             =   3540
            Width           =   2025
         End
         Begin VB.TextBox txtBTID 
            Height          =   300
            Left            =   1365
            TabIndex        =   50
            Top             =   120
            Width           =   2010
         End
         Begin VB.TextBox txtCustom 
            Height          =   300
            Left            =   1365
            TabIndex        =   51
            Top             =   480
            Width           =   2010
         End
         Begin VB.TextBox txtBTNO 
            Height          =   300
            Left            =   1350
            TabIndex        =   52
            Top             =   840
            Width           =   2010
         End
         Begin VB.TextBox txtArticle 
            Height          =   300
            Left            =   1350
            TabIndex        =   53
            Top             =   1185
            Width           =   2010
         End
         Begin VB.TextBox txtRecpPerson 
            Height          =   300
            Left            =   1350
            TabIndex        =   55
            Top             =   1905
            Width           =   2010
         End
         Begin VB.TextBox txtPerson 
            Height          =   300
            Left            =   1350
            TabIndex        =   54
            Top             =   1545
            Width           =   2010
         End
         Begin VB.TextBox txtRemark 
            Height          =   855
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   57
            Top             =   2640
            Width           =   2370
         End
         Begin MSComCtl2.DTPicker dtpBt 
            Height          =   300
            Index           =   0
            Left            =   1365
            TabIndex        =   56
            Top             =   2265
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   529
            _Version        =   393216
            Format          =   54657024
            CurrentDate     =   37112
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "B/T관리번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "B/T의뢰번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "거  래  처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "품       명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   3
            Left            =   3405
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   480
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   11
            Left            =   120
            TabIndex        =   29
            Top             =   2265
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "접수일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   120
            TabIndex        =   34
            Top             =   1545
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "실 험 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   5
            Left            =   3375
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1530
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   13
            Left            =   120
            TabIndex        =   45
            Top             =   1905
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "접수 작성자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   6
            Left            =   3375
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1905
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   18
            Left            =   120
            TabIndex        =   47
            Top             =   2655
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "비고 사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   4
            Left            =   3390
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1185
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   330
            Left            =   90
            TabIndex        =   84
            Top             =   3540
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   196609
            Caption         =   "담 당 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   8
            Left            =   3405
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   3540
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   330
            Left            =   75
            TabIndex        =   86
            Top             =   3930
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   196609
            Caption         =   "발송일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpSendDate 
            Height          =   330
            Left            =   1350
            TabIndex        =   59
            Top             =   3930
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   582
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
            Format          =   54657024
            CurrentDate     =   37112
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
         Height          =   4395
         Left            =   3960
         TabIndex        =   62
         Top             =   915
         Width           =   3375
         _cx             =   5953
         _cy             =   7752
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
      Begin Threed.SSCommand cmdAddNew 
         Height          =   450
         Left            =   5550
         TabIndex        =   60
         Top             =   420
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   794
         _Version        =   196609
         Caption         =   "색상추가"
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   450
         Left            =   6465
         TabIndex        =   61
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   794
         _Version        =   196609
         Caption         =   "색상삭제"
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   5445
      TabIndex        =   9
      Top             =   495
      Width           =   1635
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   0
      Left            =   5445
      TabIndex        =   8
      Top             =   135
      Width           =   1635
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "삭제(&D)"
      Height          =   675
      Index           =   2
      Left            =   9990
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   7
      ToolTipText     =   "자료 삭제"
      Top             =   510
      Width           =   795
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "수정(&U)"
      Height          =   675
      Index           =   1
      Left            =   9195
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   6
      ToolTipText     =   "자료 수정"
      Top             =   510
      Width           =   795
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "추가(&A)"
      Height          =   675
      Index           =   0
      Left            =   8400
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   5
      ToolTipText     =   "자료 추가"
      Top             =   510
      Width           =   795
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   1050
      Left            =   10935
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   2
      ToolTipText     =   "자료 저장"
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   465
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8580
      TabIndex        =   3
      Top             =   8535
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   10230
      TabIndex        =   4
      Top             =   8520
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   8130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   4
      Left            =   3915
      TabIndex        =   11
      Top             =   495
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품      명"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   1410
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   0
      Left            =   7125
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   135
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   9
      Left            =   3915
      TabIndex        =   14
      Top             =   135
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   60
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   2490
      TabIndex        =   16
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Format          =   54657025
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   2490
      TabIndex        =   17
      Top             =   480
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Format          =   54657025
      CurrentDate     =   36871
   End
   Begin VSFlex7LCtl.VSFlexGrid grdBt 
      Height          =   7245
      Left            =   30
      TabIndex        =   18
      Top             =   1230
      Width           =   11805
      _cx             =   20823
      _cy             =   12779
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
      Begin VSFlex7LCtl.VSFlexGrid grdBtShow 
         Height          =   2355
         Left            =   7275
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   3840
         _cx             =   6773
         _cy             =   4154
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
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   30
      Top             =   0
      Width           =   0
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   810
      TabIndex        =   35
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T 접수일자"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   36
         Top             =   60
         Value           =   1  '확인
         Width           =   1425
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   8
      Left            =   810
      TabIndex        =   37
      Top             =   825
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T 발송일자"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   38
         Top             =   60
         Width           =   1410
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   2
      Left            =   2490
      TabIndex        =   39
      Top             =   840
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   54657025
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   10
      Left            =   3915
      TabIndex        =   40
      Top             =   840
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "실 험 자"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   41
         Top             =   60
         Width           =   1410
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   7125
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   495
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   6930
      TabIndex        =   44
      Top             =   8535
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdSend 
      Height          =   690
      Left            =   5040
      TabIndex        =   49
      Top             =   8535
      Visible         =   0   'False
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "발송일등록"
      PictureAlignment=   9
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   2
      Left            =   7125
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   7680
      TabIndex        =   82
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T NO"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   83
         Top             =   60
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmBT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\BtList.rpt"

Private Const LIMIT_ROW1 = 25
Private Const LIMIT_ROW2 = 25
Private Const LIMIT_ROW3 = 5
Private Const LIMIT_ROW4 = 11
Private Const LIMIT_ROW5 = 10
Private Const LIMIT_WIDTH1 = 1380
Private Const LIMIT_WIDTH2 = 1635
Private Const LIMIT_WIDTH3 = 1965
Private Const LIMIT_WIDTH4 = 2085
Private Const LIMIT_WIDTH5 = 1890
Private Const LIMIT_WIDTH6 = 1000

Private m_sFlag         As String
Private m_nSelected     As Integer
Private m_bLoading      As Boolean
Private m_bSortForward  As Boolean
Private m_bSaved        As Boolean
Private m_sBtID         As String
Private m_nBtSeq        As Integer





Private Sub cmdCancel_Click(Index As Integer)
    pnlModify.Visible = False
    cmdoperate(0).Enabled = True
    cmdoperate(1).Enabled = True
    cmdoperate(2).Enabled = True
    
    txtBTID.Locked = False
    txtBTNO.Locked = False
    txtCustom.Locked = False
    txtArticle.Locked = False
    
    cmdFind(3).Enabled = True
    cmdFind(4).Enabled = True
End Sub

Private Sub cmdExcel_Click()
    If grdBt.Rows = 1 Then
        MsgBox LoadResString(203), vbInformation
        Exit Sub
    End If
    Call MakeExcelGrid(grdBt)
End Sub


Private Sub cmdRework_Click()
    
    If grdBt.Rows = grdBt.FixedRows Then Exit Sub
    
    If grdBt.IsSubtotal(grdBt.Row) = True Then Exit Sub
    
    If Len(grdBt.TextMatrix(grdBt.Row, 9)) = 0 Then
        MsgBox "발송 처리를 하지 않은 건입니다", vbInformation, "재 등록"
        Exit Sub
    End If

    pnlModify.Visible = True
    pnlText.Caption = "  새로운 B/T내역을 입력하십시오"
    
    Call ClearData
    m_sFlag = ID_ADDNEW

    cmdoperate(0).Enabled = False
    cmdoperate(1).Enabled = False
    cmdoperate(2).Enabled = False
                
    txtBTID.Locked = True
    txtBTNO.Locked = True
    txtCustom.Locked = True
    txtArticle.Locked = True
    
    cmdFind(3).Enabled = False
    cmdFind(4).Enabled = False
    
    Call ShowBTDetail
    
End Sub

Private Sub cmdSave_Click()
    If SaveData() Then
        pnlModify.Visible = False
        cmdoperate(0).Enabled = True
        cmdoperate(1).Enabled = True
        cmdoperate(2).Enabled = True
        
        txtBTID.Locked = False
        txtBTNO.Locked = False
        txtCustom.Locked = False
        txtArticle.Locked = False
        
        cmdFind(3).Enabled = True
        cmdFind(4).Enabled = True
        
        Call FillGridBt
        
    End If
End Sub


Private Sub cmdSend_Click()

    If grdBt.Rows = grdBt.FixedRows Then Exit Sub
    
    If grdBt.IsSubtotal(grdBt.Row) = True Then Exit Sub
        
    cmdUnSend.Enabled = False
    pnlSend.Visible = True
    pnlSend.Move 2835, 3900
    
    If Len(grdBt.TextMatrix(grdBt.Row, 9)) > 0 Then
        cmdUnSend.Enabled = True
        txtSendPerson = grdBt.TextMatrix(grdBt.Row, 10)
        txtSendPerson.Tag = grdBt.TextMatrix(grdBt.Row, 17)
        dtpSend = MakeDate(DF_FULL, grdBt.TextMatrix(grdBt.Row, 4))
    Else
        txtSendPerson = ""
        txtSendPerson.Tag = ""
        dtpSend = Now
    End If
End Sub


Private Sub cmdOK_Click()
    Dim oBt As PlusLib2.CBt
    Dim i%, sBTID$, nReworkSeq%, nBTSeq%
    Dim sSendDate$, sSendPerson$, sRecpDate$
    Dim bResult As Boolean
    Dim sMessage$

    On Error GoTo ErrHandler

    With grdBt
        sBTID = .TextMatrix(.Row, 13)
        nBTSeq = .TextMatrix(.Row, 2)
        sRecpDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 3))
    End With
    
    sSendDate = MakeDate(DF_SHORT, dtpSend)
    sSendPerson = txtSendPerson.Tag
    
    If sSendDate < sRecpDate Then
        MessageBox "발송일자는 접수일보다 빠를 수 없습니다."
        Exit Sub
    End If
    
    If Len(txtSendPerson.Tag) = 0 Then
        MessageBox "발송자를 입력하십시오"
        Exit Sub
    End If
    

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    
    bResult = oBt.UpdateBtSend(sBTID, nBTSeq, sSendDate, sSendPerson)
  
    If bResult = True Then
    
        sMessage = "B/T ID : " & MakeBTID(sBTID, OM_EXPAND) & "   차수 : " & nBTSeq & vbCrLf & "발송 등록이 정상적으로 이루어졌습니다"
        MessageBox sMessage
        pnlSend.Visible = False
        
        Set oBt = Nothing
        
        m_sBtID = sBTID
        m_nBtSeq = nBTSeq
        m_bSaved = True
        
        Call FillGridBt
    Else
        MessageBox "발송 등록이 실패하였습니다"
    End If
    
    Set oBt = Nothing
    
Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault

    Set oBt = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub



Private Sub cmdNO_Click(Index As Integer)
    pnlSend.Visible = False

End Sub


Private Sub cmdUnSend_Click()
    Dim oBt As PlusLib2.CBt
    Dim i%, sBTID$, nReworkSeq%, nBTSeq%
    Dim nMaxSeq%
    Dim sMessage$
    Dim bResult As Boolean

    On Error GoTo ErrHandler
    
    With grdBt
        sBTID = .TextMatrix(.Row, 13)
        nBTSeq = .TextMatrix(.Row, 2)
    End With
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    
    nMaxSeq = oBt.GetLastSeq(sBTID)
    
    If nMaxSeq <> nBTSeq Then
        MessageBox "마지막 차수의 B/T 내역만 발송 해제작업이 가능합니다"
        Set oBt = Nothing
        Exit Sub
    End If
        
    bResult = oBt.UpdateBtUnSend(sBTID, nBTSeq)
  
    If bResult = True Then
    
        sMessage = "B/T ID : " & MakeBTID(sBTID, OM_EXPAND) & "   차수 : " & nBTSeq & vbCrLf & "발송 등록이 해제되었습니다"
        MessageBox sMessage
        pnlSend.Visible = False
        
        Set oBt = Nothing
        
        m_sBtID = sBTID
        m_nBtSeq = nBTSeq
        m_bSaved = True
        
        Call FillGridBt
    Else
        MessageBox "발송 해제에 실패하였습니다"
    End If
    
    Set oBt = Nothing
    
Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault

    Set oBt = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub dtpBt_Change(Index As Integer)
    If Index = 0 Then
        dtpSendDate.Value = dtpBt(Index)
    End If
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11985, 9660
    Dim i%
    
    Call SetOperate(Me)
    ReDim m_nDeleteSeq(5)
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    dtpDate(2) = Now
    dtpSendDate = Now

    dtpBt(0) = Now
    
    cmdSave.MousePointer = ssCustom
    cmdSave.Picture = LoadResPicture("SAVE", vbResIcon)
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
    cmdCancel(0).Picture = LoadResPicture("EXIT", vbResIcon)
    cmdRework.Picture = LoadResPicture("ORDER", vbResIcon)
    cmdSave.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdSend.Picture = LoadResPicture("CLOSE", vbResIcon)
    
    For i = 0 To 8
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i

    Call InitGrid
    Call ClearData

    Show

    txtSearch(0).Enabled = False
    txtSearch(1).Enabled = False
    cmdFind(0).Enabled = False
    cmdFind(1).Enabled = False
    cmdFind(2).Enabled = False
    
 '   Call FillGridBt
End Sub

Private Sub chkSearch_Click(Index As Integer)

    If Index > 2 Then
        If Index = 3 Then
            If chkSearch(3).Value = vbChecked Then
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            Else
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            End If
        ElseIf Index = 4 Then
            If chkSearch(4).Value = vbChecked Then
                dtpDate(2).Enabled = True
            Else
                dtpDate(2).Enabled = False
            End If
        
        ElseIf Index = 5 Then
            If chkSearch(5).Value = vbChecked Then
                txtSearch(3).Enabled = True
            Else
                txtSearch(3).Enabled = False
            End If
        End If
    Else
        
        If chkSearch(Index) Then
            If Index = 0 Then
                cmdFind(0).Enabled = True
                txtSearch(0).Enabled = True
                txtSearch(0).SetFocus
            ElseIf Index = 1 Then
                cmdFind(1).Enabled = True
                txtSearch(1).Enabled = True
                txtSearch(1).SetFocus
            ElseIf Index = 2 Then
                cmdFind(2).Enabled = True
                txtSearch(2).Enabled = True
                txtSearch(2).SetFocus
            ElseIf Index = 5 Then
                txtSearch(3).Enabled = True
                txtSearch(3).SetFocus
            End If
        Else
            If Index = 0 Then
                cmdFind(0).Enabled = False
                txtSearch(0).Enabled = False
            ElseIf Index = 1 Then
                cmdFind(1).Enabled = False
                txtSearch(1).Enabled = False
            ElseIf Index = 2 Then
                cmdFind(2).Enabled = False
                txtSearch(2).Enabled = False
            ElseIf Index = 5 Then
                txtSearch(3).Enabled = False
            End If
        End If
    End If
End Sub



Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then       ' 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then   ' 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If

    cmdSearch.SetFocus
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub


Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
        Case 1
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
        Case 2
            Call ReturnCode(LG_PERSON, , False, txtSearch(2))
            
        Case 3
            Call ReturnCode(LG_CUSTOM, , False, txtCustom)
        Case 4
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Case 5
            Call ReturnCode(LG_PERSON, , False, txtPerson)
            txtSendPer = txtPerson
            txtSendPer.Tag = txtPerson.Tag
        Case 6
            Call ReturnCode(LG_PERSON, , False, txtRecpPerson)
        Case 7
            Call ReturnCode(LG_PERSON, , False, txtSendPerson)
        Case 8
            Call ReturnCode(LG_PERSON, , False, txtSendPer)
    End Select
End Sub

Private Sub cmdSearch_Click()
    Call FillGridBt
End Sub


Private Sub grdBt_DblClick()
    With grdBt
        If .MouseRow < .FixedRows Then Exit Sub
        
        .Row = .MouseRow
        If .IsSubtotal(.Row) Then Exit Sub
        
        Call cmdOperate_Click(ID_UPDATE)
    End With
End Sub


Private Sub grdBt_RowColChange()
    If m_bLoading Then Exit Sub

    If optView(1).Value = True Then
        grdBtShow.Visible = False
        Exit Sub
    Else
        grdBtShow.Visible = True
    
    End If
    
    With grdBt
        If .IsSubtotal(.Row) = True Then
            grdBtShow.Visible = False
            Exit Sub
        Else
            grdBtShow.Visible = True
        End If
    
        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Sub

        Call ShowBTData

        .SetFocus
    End With
End Sub




Private Sub grdDyeAux_DblClick()
    With grdDyeAux
        .EditCell
    End With
End Sub



Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyReturn Then Exit Sub
    
    If Index = 3 Then
        Call cmdFind_Click(1)
    ElseIf Index = 4 Then
        Call cmdFind_Click(2)
    
    End If
    
End Sub

Private Sub optView_Click(Index As Integer, Value As Integer)
    
    With grdBt
    
        If .Rows = .FixedRows Then Exit Sub
        
        If .IsSubtotal(.Row) = True Then Exit Sub
    
    End With
    
    If Index = 1 Then
        grdBtShow.Visible = False
    Else
        grdBtShow.Visible = True
        Call ShowBTData
    End If
End Sub



Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnRef(LG_ARTICLE, , False, txtArticle)
        'Call ReturnCode(LG_ARTICLE, , , txtArticle)
    End If
End Sub


Private Sub txtCustom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        Call ReturnRef(LG_CUSTOM, , False, txtCustom)
        
    End If
End Sub


Private Sub txtPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_PERSON, , , txtPerson)
        txtSendPer = txtPerson
        txtSendPer.Tag = txtPerson.Tag
    End If
End Sub


Private Sub txtRecpPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_PERSON, , , txtRecpPerson)
    End If
End Sub



Private Sub txtSendPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_PERSON, , , txtSendPerson)
    End If

End Sub


Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub


Private Sub chkSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
        ElseIf Index = 1 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_PERSON, , False, txtSearch(2))
        End If
        
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim i%

    Select Case Index
        Case ID_ADDNEW
            pnlModify.Visible = True
            pnlText.Caption = "  새로운 B/T내역을 입력하십시오"
            
            Call ClearData
            m_sFlag = ID_ADDNEW

            cmdoperate(0).Enabled = False
            cmdoperate(1).Enabled = False
            cmdoperate(2).Enabled = False
            
        Case ID_UPDATE
            If grdBt.IsSubtotal(grdBt.Row) Then
                MsgBox ("세부 데이터를 선택 하십시오")
                Exit Sub
            End If

            If grdBt.Rows = grdBt.FixedRows Then Exit Sub
            pnlText.Caption = "  수정할 B/T내역을 입력하십시오"
            
            pnlModify.Visible = True
            
            m_sFlag = ID_UPDATE
            cmdoperate(0).Enabled = False
            cmdoperate(1).Enabled = False
            cmdoperate(2).Enabled = False
            
            txtBTID.Locked = True
'            txtBTNO.Locked = True
            txtCustom.Locked = False
                        
            cmdFind(3).Enabled = True
                        
            Call ShowBTDetail
            
        Case ID_DELETE
            If grdBt.IsSubtotal(grdBt.Row) Then
                MsgBox ("세부 데이터를 선택 하십시오")
                Exit Sub
            End If
            
            If grdBt.Rows = grdBt.FixedRows Then Exit Sub
            

            If Not QuestionBox(LoadResString(201)) Then Exit Sub

            If DeleteData() Then Call FillGridBt
            
    End Select
End Sub

Private Sub ShowBTDetail()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, sBTID$, nBTSeq%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    With grdBt
        sBTID = .TextMatrix(.Row, 13)
        nBTSeq = .TextMatrix(.Row, 2)
    End With
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    Set rs = oBt.GetBtOne(sBTID, nBTSeq)
    
'    txtCustom.Locked = True
    
    txtBTID = MakeBTID(rs!BTID, OM_EXPAND)
    txtBTID.Tag = rs!BTID
    txtCustom = rs!kCustom
    txtCustom.Tag = rs!CustomID
    txtBTNO = rs!BTNO
    txtArticle = CheckNull(rs!Article)
    txtArticle.Tag = CheckNull(rs!ArticleID)
    
    txtRecpPerson = CheckNull(rs!RecpName)
    txtRecpPerson.Tag = CheckNull(rs!RecpPerID)
    txtPerson = CheckNull(rs!Name)
    txtPerson.Tag = CheckNull(rs!PersonID)
    txtRemark = CheckNull(rs!Remark)
    txtSendPer = CheckNull(rs!SendName)
    txtSendPer.Tag = CheckNull(rs!SendPerID)
    
    dtpBt(0) = MakeDate(DF_LONG, rs!Recpdate)
    dtpSendDate = MakeDate(DF_LONG, rs!SendDate)
        
    Set rs = Nothing
    
    Set rs = oBt.GetBtSub(sBTID, nBTSeq)
    With grdDyeAux
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!Color & vbTab & rs!ColorSeq

            rs.MoveNext
        Next i
        .Redraw = flexRDDirect
    End With
    
    Set rs = Nothing
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set oBt = Nothing
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub

Private Sub cmdAddNew_Click()
    Dim i%

    With grdDyeAux
        .Rows = .Rows + 1
    
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = CStr(i)
        Next i
        
        .SetFocus
        .Select .Rows - 1, 1
        
        If .Row = .FixedRows Then
            .TextMatrix(.Rows - 1, 2) = 1
        Else
            i = .TextMatrix(.Rows - 2, 2) + 1
            .TextMatrix(.Rows - 1, 2) = i
        End If
        
        .EditCell
        
    End With
End Sub

Private Sub cmdDelete_Click()
    With grdDyeAux
        If .Rows = 1 Or .Row < 1 Then
            MsgBox LoadResString(204), vbInformation
        Else
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                
                .RemoveItem .Row
            End If

        End If
    End With

End Sub


Private Sub grdDyeAux_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdDyeAux
        If Col = 1 Then
            If .TextMatrix(.Row, 1) = "" Then
                If QuestionBox("색상명이 입력되지 않았습니다" & vbCrLf & vbCrLf & "입력을 계속하시겠습니까?") Then
                    .EditCell
                Else
                    .RemoveItem .Row
                End If
            Else
                If Row = .Rows - 1 Then
                    If QuestionBox("색상" & "을 계속 추가하시겠습니까 ?") Then
                        Call cmdAddNew_Click
                    Else
                        cmdSave.SetFocus
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub cmdPrint_Click()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim nChkDate%, sDate$, eDate$
    Dim nChkSendDate%, SendDate$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim nChkPerson%, sPersonID$
    Dim nChkBTNO%, sBTNO$
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    If grdBt.Rows = grdBt.FixedRows Then Exit Sub
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    nChkDate = IIf(chkSearch(3), 1, 0)          ' 접수일
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkSendDate = IIf(chkSearch(4), 1, 0)      ' 발송일
    SendDate = MakeDate(DF_SHORT, dtpDate(2))
    nChkCustom = IIf(chkSearch(0), 1, 0)        ' 거래처
    sCustom = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(1), 1, 0)       ' 품명
    sArticle = txtSearch(1).Tag
    nChkPerson = IIf(chkSearch(2), 1, 0)        ' 작성자
    sPersonID = txtSearch(2).Tag
    nChkBTNO = IIf(chkSearch(5), 1, 0)
    sBTNO = txtSearch(3)
    
    Set rs = oBt.GetBtList(nChkDate, sDate, eDate, nChkSendDate, SendDate, nChkCustom, sCustom, _
            nChkArticle, sArticle, nChkPerson, sPersonID, nChkBTNO, sBTNO)
    
    Set oBt = Nothing
    
    
    ReDim sParam(1)
    sParam(0) = "B/T 접수대장"
    sParam(1) = CompanyName
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
    
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
    
    With grdBt
        .Cols = 18
        .Rows = 1
        
        .Redraw = flexRDNone
        
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarVertical
        .ExplorerBar = flexExSortShow
        .ScrollTrack = True
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 550
        .ColWidth(0) = 360
        .RowHeightMin = 450

        .TextArray(1) = "거래처":                       .ColWidth(1) = 2500:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "접수번호" & vbCrLf & "차수":   .ColWidth(2) = 1600:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "B/T NO" & vbCrLf & "접수일자": .ColWidth(3) = 2500:    .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "품명" & vbCrLf & "발송일자":   .ColWidth(4) = 3000:    .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "색상수":                       .ColWidth(5) = 1000:     .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "실험자":                       .ColWidth(6) = 900:     .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "접수" & vbCrLf & "등록일":     .ColWidth(7) = 1100:    .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(8) = "접수" & vbCrLf & "작성자":     .ColWidth(8) = 1100:    .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(9) = "발송" & vbCrLf & "등록일":     .ColWidth(9) = 1100:    .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(10) = "발송" & vbCrLf & "작성자":    .ColWidth(10) = 900:    .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(11) = "거래처":                      .ColWidth(11) = 0
        .TextArray(12) = "거래처ID":                    .ColWidth(12) = 0
        .TextArray(13) = "BTID":                        .ColWidth(13) = 0
        .TextArray(14) = "BTNO":                        .ColWidth(14) = 0
        .TextArray(15) = "품명":                        .ColWidth(15) = 0
        .TextArray(16) = "품명ID":                      .ColWidth(16) = 0
        .TextArray(17) = "발송작성자ID":                .ColWidth(17) = 0

        For i = 6 To 10
            .ColHidden(i) = True
            
        Next i
        
        .Redraw = flexRDDirect
    End With

    With grdBtShow
        .Cols = 2
        Call SetVSFlexGrid(grdBtShow)

        .Redraw = False

        .TextArray(1) = "색상명":     .ColWidth(1) = 900:             .ColAlignment(1) = flexAlignLeftCenter

        .Redraw = True
    End With
    
    With grdDyeAux
        .Cols = 3
        Call SetVSFlexGrid(grdDyeAux)
        .Redraw = flexRDNone
        
        .TextArray(1) = "색상명":       .ColWidth(1) = 1800:         .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "일련번호":     .ColWidth(2) = 0  '        .ColHidden(2) = True
        
        .ExtendLastCol = True
        
        .FocusRect = flexFocusSolid
        .FloodColor = RGB(255, 0, 0)
        .Redraw = flexRDDirect
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub FillGridBt()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim nChkDate%, sDate$, eDate$
    Dim nChkSendDate%, SendDate$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim nChkPerson%, sPersonID$
    Dim nChkBTNO%, sBTNO$
    Dim sPreBTID$, nCnt%, nBeforeTop%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bLoading = True

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    nChkDate = IIf(chkSearch(3), 1, 0)          ' 접수일
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkSendDate = IIf(chkSearch(4), 1, 0)      ' 발송일
    SendDate = MakeDate(DF_SHORT, dtpDate(2))
    nChkCustom = IIf(chkSearch(0), 1, 0)        ' 거래처
    sCustom = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(1), 1, 0)       ' 품명
    sArticle = txtSearch(1).Tag
    nChkPerson = IIf(chkSearch(2), 1, 0)        ' 작성자
    sPersonID = txtSearch(2).Tag
    nChkBTNO = IIf(chkSearch(5), 1, 0)
    sBTNO = txtSearch(3)
    
    Set rs = oBt.GetBtList(nChkDate, sDate, eDate, nChkSendDate, SendDate, nChkCustom, sCustom, _
            nChkArticle, sArticle, nChkPerson, sPersonID, nChkBTNO, sBTNO)
    
    Set oBt = Nothing

    nCnt = 1
    
    With grdBt
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            
            If sPreBTID <> rs!BTID Then
                sPreBTID = rs!BTID
                
                .AddItem CStr(nCnt) & vbTab & rs!kCustom & vbTab & MakeBTID(rs!BTID, OM_EXPAND) & vbTab & rs!BTNO & vbTab & rs!Article
            
                Call DoFlexGridGroup(.Rows - 1, 1)  ' 그리드 서브토탈
'                Call GridCollapse(nBeforeTop)       ' 서브토탈 row를 접힌 상태로 출력
'
                nBeforeTop = .Rows - 1
                nCnt = nCnt + 1
            End If
            
            .AddItem "" & vbTab & " " & vbTab & rs!BTIDSeq & vbTab & MakeDate(DF_LONG, CheckNull(rs!Recpdate)) & vbTab & MakeDate(DF_LONG, CheckNull(rs!SendDate)) & vbTab & _
                        rs!ColorCnt & vbTab & CheckNull(rs!Name) & vbTab & MakeDate(DF_LONG, CheckNull(rs!RecpDTime)) & vbTab & CheckNull(rs!RecpName) & vbTab & _
                        MakeDate(DF_LONG, CheckNull(rs!SendDTime)) & vbTab & CheckNull(rs!SendName) & vbTab & CheckNull(rs!kCustom) & vbTab & CheckNull(rs!CustomID) & vbTab & _
                        rs!BTID & vbTab & rs!BTNO & vbTab & CheckNull(rs!Article) & vbTab & CheckNull(rs!ArticleID) & vbTab & CheckNull(rs!SendPerID)
                        
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If

            rs.MoveNext
        Next i

        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
        
        If .Rows > .FixedRows Then
            cmdPrint.Enabled = True
            .HighLight = flexHighlightAlways
            cmdoperate(ID_UPDATE).Enabled = True
            cmdoperate(ID_DELETE).Enabled = True

            If m_bSaved = True Then
                Call FindNewRow(m_sBtID, m_nBtSeq)
            End If
        Else
            cmdPrint.Enabled = False
            .HighLight = flexHighlightNever
            cmdoperate(ID_UPDATE).Enabled = False
            cmdoperate(ID_DELETE).Enabled = False
            grdBtShow.Visible = False
            
            MsgBox LoadResString(203), vbInformation
        End If
        
        .SetFocus
    End With

    Screen.MousePointer = vbDefault

    m_bSaved = False
    m_bLoading = False

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oBt = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub DoFlexGridGroup(Row As Integer, Level As Integer)
    With grdBt
        ' Set the row as a group
        .IsSubtotal(Row) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(Row) = 1

        Select Case Level
            Case 0
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = &HE0E0E0
                .Cell(flexcpFontBold, Row, 0, Row, .Cols - 1) = True
            Case 1, 2
                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = &HE0E0E0
        End Select
        
        
    End With
End Sub

Private Sub GridCollapse(Row As Integer)
    
    With grdBt
    
        If Row >= .FixedRows Then
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub


Private Function MakeBTID(sBTID As String, nType As EORDERMAKE) As String
     If nType = OM_EXPAND Then
        MakeBTID = Left(sBTID, 2) & "-" & Mid(sBTID, 3, 2) & "-" & Mid(sBTID, 5, 4)
    Else
        MakeBTID = Replace(sBTID, "-", "")
    End If
    

End Function

Private Sub FindNewRow(sBTID As String, nSeq As Integer)
    Dim i%
    
    With grdBt
        For i = .FixedRows To .Rows - 1
            If .IsSubtotal(i) = False Then
                If (.TextMatrix(i, 13) = sBTID) And (.TextMatrix(i, 2) = nSeq) Then
                    .Row = i
                    .TopRow = i
                    Exit Sub
                End If
            End If
        Next i
    
    End With

End Sub

'
'Private Function IsGetOrder() As Boolean
'    IsGetOrder = False
'
'    With grdOrder
'        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Function
'    End With
'    With grdColor
'        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Function
'    End With
'
'    IsGetOrder = True
'End Function

Private Sub ClearData()

    txtBTID = ""
    txtCustom = ""
    txtCustom.Tag = ""
    txtBTNO = ""
    txtArticle = ""
    txtArticle.Tag = ""
    txtPerson = ""
    txtPerson.Tag = ""
    txtRecpPerson = ""
    txtRecpPerson.Tag = ""
    dtpBt(0) = Now
    txtRemark = ""
    
    grdDyeAux.Rows = grdBtShow.FixedRows

    m_sFlag = ID_ADDNEW
'
End Sub



Private Sub ShowBTData()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, sBTID$, nReworkSeq%, nBTSeq%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    With grdBt
        sBTID = .TextMatrix(.Row, 13)
        nBTSeq = .TextMatrix(.Row, 2)
    End With

    Set rs = oBt.GetBtSub(sBTID, nBTSeq)
    With grdBtShow
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!Color

            rs.MoveNext
        Next i

        
        If .Rows > .FixedRows Then
            
            If .Rows < LIMIT_ROW5 Then
                .Height = (.RowHeight(.FixedRows) + 40) * .Rows + 350
                .ScrollBars = flexScrollBarNone
            Else
                .Height = 2700
                .ScrollBars = flexScrollBarVertical
            End If
        Else
            .Height = .RowHeight(0) + 110
        End If
        
        .Redraw = True
        .SetFocus
    End With
    
    With grdBt
        If .Rows = .FixedRows Then Exit Sub

        If .Row < (.TopRow + 7) Then
            grdBtShow.Top = 4400
        Else
            grdBtShow.Top = 900
        End If
    End With

    rs.Close

    Set rs = Nothing
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oBt = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Function CheckData() As Boolean
    Dim oBt As PlusLib2.CBt
    
    
    CheckData = False

    
    If Len(txtBTNO) <= 0 Then
        Call MessageBox("의뢰번호를 입력하십시오.")
        cmdSearch.SetFocus
        Exit Function
    End If
    
    If Len(txtArticle.Tag) = 0 Then
        Call MessageBox("품명을 입력하십시오.")
        txtArticle.SetFocus
        Exit Function
    End If
    
    If Len(txtCustom.Tag) <> 4 Then
        Call MessageBox("거래처를 입력하십시오.")
        txtCustom.SetFocus
        Exit Function
    End If
    
    

    Dim i%

    With grdDyeAux
        If .Rows = .FixedRows Then
            Call MessageBox("색상을 입력하십시오.")
            cmdAddNew.SetFocus
            Exit Function
        End If
    End With

    CheckData = True
End Function



Private Function SaveData() As Boolean
    Dim tBtList   As PlusLib2.TBt
    Dim tBtListSub() As PlusLib2.TBtSub
    Dim oBt   As PlusLib2.CBt
    Dim i%, nColorCnt%, nBTSub%

    SaveData = False
    If m_sFlag = ID_ADDNEW Then
        If Not CheckData Then Exit Function
    End If
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    With tBtList
        .sBTID = Replace(txtBTID, "-", "")
        If grdBt.Rows = grdBt.FixedRows Then
            .nBTSeq = 0
        Else
            .nBTSeq = grdBt.TextMatrix(grdBt.Row, 2)
        End If
        .sCustom = txtCustom
        .sCustomID = Format(txtCustom.Tag, "0000")
        .sBTNO = txtBTNO
        .sArticle = txtArticle
        .sArticleID = Format(txtArticle.Tag, "0000")
        .nColorCnt = 0
        .sRecpDate = MakeDate(DF_SHORT, dtpBt(0))
        .sPersonID = Format(txtPerson.Tag, "00000000")
        .RecpPersonID = Format(txtRecpPerson.Tag, "00000000")
        .RecpDTime = Now
        .Remark = txtRemark
        .sSendDate = MakeDate(DF_SHORT, dtpSendDate)
        .sSendPersonID = Format(txtSendPer.Tag, "00000000")
    End With

    nBTSub = (grdDyeAux.Rows - 2)
    With grdDyeAux
        For i = 0 To nBTSub
            If Not .TextMatrix(i + 1, 1) = "" Then
                ReDim Preserve tBtListSub(nColorCnt)
                tBtListSub(i).sBTID = Replace(txtBTID, "-", "")
                tBtListSub(i).nBTSeq = 0
                tBtListSub(i).nColorSeq = .TextMatrix(i + 1, 2)
                tBtListSub(i).sColor = .TextMatrix(i + 1, 1)
                
                nColorCnt = nColorCnt + 1
            End If
        Next i
    End With
    tBtList.nColorCnt = nColorCnt
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    oBt.UserName = g_sUserName

    If m_sFlag = ID_ADDNEW Then
        SaveData = oBt.AddNewBt(tBtList, tBtListSub)
    Else
        tBtList.sBTID = tBtList.sBTID
        SaveData = oBt.UpdateBt(tBtList, tBtListSub)
    End If

    Set oBt = Nothing
    Screen.MousePointer = vbDefault
    
    m_bSaved = True
    m_sBtID = tBtList.sBTID
    m_nBtSeq = tBtList.nBTSeq
    
    txtBTID.Locked = False
    txtBTNO.Locked = False

    Exit Function

ErrHandler:
    SaveData = False
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

Private Function DeleteData() As Boolean
    Dim oBt As PlusLib2.CBt
    Dim nMaxSeq%, sBTID$, nBTSeq%
    
    On Error GoTo ErrHandler

    DeleteData = False

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    oBt.UserName = g_sUserName
    
    With grdBt
        sBTID = .TextMatrix(.Row, 13)
        nBTSeq = .TextMatrix(.Row, 2)
    End With

    nMaxSeq = oBt.GetLastSeq(sBTID)
    
    If nMaxSeq <> nBTSeq Then
        MessageBox "마지막 차수의 B/T 내역만 삭제가능합니다"
        Set oBt = Nothing
        Exit Function
    End If

    DeleteData = oBt.DeleteBt(sBTID, nBTSeq)

    Set oBt = Nothing

    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Set oBt = Nothing
End Function

Private Sub ChangeScrollBT()
    With grdBt
        .ColWidth(3) = IIf(.Rows > LIMIT_ROW1 + .FixedRows, LIMIT_WIDTH1 - 240, LIMIT_WIDTH1)
    End With
End Sub


Private Sub ChangeScrollBtShow()
    With grdBtShow
   '     .ColWidth(1) = IIf(.Rows > LIMIT_ROW5 + .FixedRows, LIMIT_WIDTH5 - 240, LIMIT_WIDTH5)
    End With
End Sub


Private Sub ChangeScrollDyeAux()
    With grdDyeAux
        .ColWidth(1) = IIf(.Rows > LIMIT_ROW4 + .FixedRows, LIMIT_WIDTH4 - 240, LIMIT_WIDTH4)
    End With
End Sub


