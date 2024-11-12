VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrderHistory 
   Caption         =   "수주별 진행"
   ClientHeight    =   9255
   ClientLeft      =   1200
   ClientTop       =   2205
   ClientWidth     =   15180
   Icon            =   "frmOrderHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15180
   Begin VB.TextBox txtWorkWidth 
      Height          =   330
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   0
      Width           =   1245
   End
   Begin VB.TextBox txtWorkName 
      Height          =   300
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   1305
   End
   Begin VB.TextBox txtPattern 
      Height          =   270
      Left            =   1005
      TabIndex        =   9
      Top             =   2625
      Width           =   14265
   End
   Begin VB.Frame fraSearch 
      Height          =   645
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   4965
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No"
         Height          =   375
         Index           =   1
         Left            =   1395
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   210
         Width           =   1290
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   375
         Index           =   0
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.TextBox txtSearch 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         MaxLength       =   20
         TabIndex        =   0
         Top             =   210
         Width           =   2190
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   660
      Left            =   13515
      TabIndex        =   1
      Top             =   0
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1164
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   -15
      TabIndex        =   3
      Top             =   2925
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   2
      TabHeight       =   600
      TabMaxWidth     =   4410
      TabCaption(0)   =   "입  고"
      TabPicture(0)   =   "frmOrderHistory.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdStuffIN"
      Tab(0).Control(1)=   "grdTotal"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "공정 대기"
      TabPicture(1)   =   "frmOrderHistory.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel1"
      Tab(1).Control(1)=   "grdProcWait"
      Tab(1).Control(2)=   "grdProcWaitDetail"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "실 적"
      TabPicture(2)   =   "frmOrderHistory.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "grdWorkResult"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin Threed.SSPanel SSPanel1 
         Height          =   765
         Left            =   -67320
         TabIndex        =   24
         Top             =   420
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   1349
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtThreadName 
            Height          =   300
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   390
            Width           =   1545
         End
         Begin VB.TextBox txtCustom 
            Height          =   300
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   60
            Width           =   3165
         End
         Begin VB.TextBox txtUseClss 
            Height          =   300
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   390
            Width           =   1545
         End
         Begin VB.TextBox txtExpectDate 
            Height          =   300
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   60
            Width           =   1545
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   25
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
            Caption         =   "완료예정일"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   1
            Left            =   60
            TabIndex        =   27
            Top             =   390
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
            Caption         =   "카드상태"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   2
            Left            =   3000
            TabIndex        =   29
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
            Caption         =   "제직처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   3
            Left            =   3000
            TabIndex        =   31
            Top             =   390
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
            Caption         =   "사종"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdWorkResult 
         Height          =   5925
         Left            =   60
         TabIndex        =   4
         Top             =   390
         Width           =   15030
         _cx             =   26511
         _cy             =   10451
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
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
         FixedRows       =   2
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
      Begin VSFlex7LCtl.VSFlexGrid grdProcWait 
         Height          =   5865
         Left            =   -74940
         TabIndex        =   19
         Top             =   420
         Width           =   7605
         _cx             =   13414
         _cy             =   10345
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
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
      Begin VSFlex7LCtl.VSFlexGrid grdProcWaitDetail 
         Height          =   5115
         Left            =   -67335
         TabIndex        =   20
         Top             =   1170
         Width           =   7425
         _cx             =   13097
         _cy             =   9022
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
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
      Begin VSFlex7LCtl.VSFlexGrid grdStuffIN 
         Height          =   5475
         Left            =   -74940
         TabIndex        =   21
         Top             =   390
         Width           =   15060
         _cx             =   26564
         _cy             =   9657
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
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
      Begin VSFlex7LCtl.VSFlexGrid grdTotal 
         Height          =   390
         Left            =   -63960
         TabIndex        =   33
         Top             =   5880
         Width           =   4050
         _cx             =   7144
         _cy             =   688
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
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   1950
      Left            =   0
      TabIndex        =   5
      Top             =   660
      Width           =   15150
      _cx             =   26723
      _cy             =   3440
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
      Begin VB.CheckBox chkExpand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "공정확장"
         Height          =   510
         Left            =   0
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   270
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   2625
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   476
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
      Caption         =   "공정 패턴"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlName 
      Height          =   300
      Index           =   9
      Left            =   5040
      TabIndex        =   11
      Top             =   15
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
      Caption         =   "접수 일자"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlName 
      Height          =   300
      Index           =   10
      Left            =   5040
      TabIndex        =   12
      Top             =   345
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
      Caption         =   "납기 일자"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlName 
      Height          =   300
      Index           =   12
      Left            =   7575
      TabIndex        =   13
      Top             =   15
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
      Caption         =   "가공 구분"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlName 
      Height          =   300
      Index           =   17
      Left            =   10080
      TabIndex        =   14
      Top             =   15
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
      Caption         =   "가공 폭"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlName 
      Height          =   300
      Index           =   33
      Left            =   7575
      TabIndex        =   15
      Top             =   330
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
      Caption         =   "가공 밀도"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin MRPPlus2.WizText txtWorkDensity 
      Height          =   300
      Left            =   8760
      TabIndex        =   16
      Top             =   315
      Width           =   1320
      _ExtentX        =   2328
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
      Locked          =   -1  'True
   End
   Begin MRPPlus2.WizText txtAcptDate 
      Height          =   300
      Left            =   6240
      TabIndex        =   17
      Top             =   30
      Width           =   1320
      _ExtentX        =   2328
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
      Locked          =   -1  'True
      Text            =   "2003-11-13"
   End
   Begin MRPPlus2.WizText txtDvlyDate 
      Height          =   300
      Left            =   6240
      TabIndex        =   18
      Top             =   330
      Width           =   1320
      _ExtentX        =   2328
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
      Locked          =   -1  'True
      Text            =   "2003-11-13"
   End
End
Attribute VB_Name = "frmOrderHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub chkExpand_Click()
    Dim i%
    With grdOrder
        For i = 13 To 21
            .ColHidden(i) = IIf(chkExpand.Value = vbChecked, False, True)
        Next i
        
        .ColHidden(15) = True
        .ColHidden(18) = True
        
        If chkExpand.Value = vbChecked Then
             .ScrollBars = flexScrollBarBoth
        Else
            .ScrollBars = flexScrollBarVertical
        End If
    End With
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub
Private Sub FillgrdWorkResult()
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    Dim II As Long, GoodQty_int As Double, NgQty_int As Double, OutQty_int As Double, GrpRow As Long

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OrdHistWorkResult"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, Trim(txtSearch.Text))

        Set dRS = .Execute
    End With
    
    Set adoCmd = Nothing
    With grdWorkResult
        .Rows = .FixedRows
        .ExplorerBar = flexExNone
        Do Until dRS.EOF
       '     If .Rows <> .FixedRows Then
                If (.TextMatrix(.Rows - 1, 2) <> Trim(dRS!ColorName)) Then
                    .AddItem "" & vbTab & "" & vbTab & Trim(dRS!ColorName)
                    Call DoFlexGridGroup(grdWorkResult, .Rows - 1, 1)
                End If
                .AddItem "" & vbTab & IIf(dRS!ReWorkClss = "*", "■", "") & vbTab & Trim(dRS!ColorName) & vbTab & _
                            MakeCardID(dRS!CardID, OM_EXPAND, dRS!SplitID) & vbTab & _
                            SetCurrency(dRS!Qty) & vbTab & Trim(dRS!UseClss) & vbTab & _
                            IIf(Trim(dRS!ResultProc) = "", Trim(dRS!AfterProc), Trim(dRS!ResultProc) & "->" & Trim(dRS!AfterProc)) & vbTab & _
                            SetCurrency(dRS!GOOD, 0) & vbTab & SetCurrency(dRS!NG, 0) & vbTab & SetCurrency(dRS!OutWare, 0)
                .RowHeight(.Rows - 1) = 550
                
                Select Case Trim(dRS!UseClss)
                    Case "보류"
                        .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, 3) = vbRed
                        .Cell(flexcpForeColor, .Rows - 1, 3, .Rows - 1, 3) = vbWhite
                    Case "작업"
                        .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, 3) = vbBlue
                        .Cell(flexcpForeColor, .Rows - 1, 3, .Rows - 1, 3) = vbWhite
                End Select
                
                dRS.MoveNext
  '          End If
        Loop
        dRS.Close
        Set dRS = Nothing
    End With
    
    GoodQty_int = 0:  NgQty_int = 0: OutQty_int = 0: GrpRow = 2
    With grdWorkResult
        For II = .FixedRows To .Rows - 1
            ' GoodQty_int As Integer, NgQty_int As Integer, OutQty_int As Integer, GrpRow As Long
            If .IsSubtotal(II) = True Then
                If II <> 0 Then
                    .TextMatrix(GrpRow, 7) = GoodQty_int
                    .TextMatrix(GrpRow, 8) = NgQty_int
                    .TextMatrix(GrpRow, 9) = OutQty_int
                End If
                GoodQty_int = 0:  NgQty_int = 0: OutQty_int = 0: GrpRow = II
            End If
            
            If .IsSubtotal(II) = False Then
                .TextMatrix(II, 2) = ""
                GoodQty_int = GoodQty_int + .TextMatrix(II, 7)
                NgQty_int = NgQty_int + .TextMatrix(II, 8)
                OutQty_int = OutQty_int + .TextMatrix(II, 9)
            End If
        Next II
        If II > 2 Then
            .TextMatrix(GrpRow, 7) = GoodQty_int
            .TextMatrix(GrpRow, 8) = NgQty_int
            .TextMatrix(GrpRow, 9) = OutQty_int
        End If
        If .Rows > .FixedRows Then
            .Cell(flexcpFontSize, .FixedRows, 6, .Rows - 1, 6) = 8
            .Row = .FixedRows
        End If
    End With
    

    
''    With grdOrderWait
''        For II = .FixedRows To .Rows - 1
''            If Trim(.TextMatrix(II, 3)) <> "" Then
''            ElseIf Trim(.TextMatrix(II, 4)) <> "" Then
''                Call DoFlexGridGroup(grdOrderWait, II, 2)
''            End If
''
''        Next II
''    End With

End Sub
Private Sub FillgrdProcWait()
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OrdHistProcWait"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, Trim(txtSearch.Text))

        Set dRS = .Execute
    End With
    
    Set adoCmd = Nothing
    With grdProcWait
        .Rows = .FixedRows
        Do Until dRS.EOF
            '-- 1단계
            '-- 2단계
            '-- 3단계
 '           WaitProcID ProcName             OrderSeq    ColorName                      CardID   SplitID Card_EA     Roll        Qty

            .AddItem ""
            .RowHeight(.Rows - 1) = 350
            
            If Len(Trim(dRS!ColorName)) = 0 Then   '1단계
                .TextMatrix(.Rows - 1, 3) = dRS!ProcName
                .TextMatrix(.Rows - 1, 5) = dRS!Card_EA
                .TextMatrix(.Rows - 1, 6) = dRS!Roll
                .TextMatrix(.Rows - 1, 7) = dRS!Qty
                Call DoFlexGridGroup(grdProcWait, .Rows - 1, 1)
                
            ElseIf Len(Trim(dRS!CardID)) = 0 Then    '2단계
                .TextMatrix(.Rows - 1, 4) = dRS!ColorName
                .TextMatrix(.Rows - 1, 5) = dRS!Card_EA
                .TextMatrix(.Rows - 1, 6) = dRS!Roll
                .TextMatrix(.Rows - 1, 7) = dRS!Qty
                Call DoFlexGridGroup(grdProcWait, .Rows - 1, 2)
                
            Else
                .TextMatrix(.Rows - 1, 5) = IIf(Trim(dRS!SplitID) <> "", Format(Trim(dRS!CardID), "0#-##-####") & "-" & dRS!SplitID, Format(Trim(dRS!CardID), "0#-##-####"))
                .TextMatrix(.Rows - 1, 6) = dRS!Roll
                .TextMatrix(.Rows - 1, 7) = dRS!Qty
                
                Select Case Trim(dRS!UseClss)
                    Case "보류"
                        .Cell(flexcpBackColor, .Rows - 1, 5, .Rows - 1, 5) = vbRed
                        .Cell(flexcpForeColor, .Rows - 1, 5, .Rows - 1, 5) = vbWhite
                    Case "작업"
                        .Cell(flexcpBackColor, .Rows - 1, 5, .Rows - 1, 5) = vbBlue
                        .Cell(flexcpForeColor, .Rows - 1, 5, .Rows - 1, 5) = vbWhite
                End Select
                
            End If
            dRS.MoveNext
        Loop
        dRS.Close
        Set dRS = Nothing
    End With
    
''    With grdOrderWait
''        For II = .FixedRows To .Rows - 1
''            If Trim(.TextMatrix(II, 3)) <> "" Then
''            ElseIf Trim(.TextMatrix(II, 4)) <> "" Then
''                Call DoFlexGridGroup(grdOrderWait, II, 2)
''            End If
''
''        Next II
''    End With

End Sub
Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15300, 9660
    
    SSTab1.Tab = 2
    Call InitGrid
    Call SetOperate(Me)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdOrder
        .Cols = 28

        .Redraw = flexRDNone

        .Rows = 2
        .FixedRows = 2
        .FixedCols = 0
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        .TextArray(0) = " ":            .ColWidth(0) = 500
        .TextArray(1) = "거래처":       .ColWidth(1) = 1750:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "품명":         .ColWidth(2) = 1550:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호" & vbCrLf & "색  상  명":             .ColWidth(3) = 1350:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "Order No." & vbCrLf & "색  상  명":            .ColWidth(4) = 0:               .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "축율":         .ColWidth(5) = 800:             .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "수주량":       .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "입고량":       .ColWidth(7) = 900:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "미계획량":     .ColWidth(8) = 900:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "계획량":       .ColWidth(9) = 900:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "배색":        .ColWidth(10) = 900:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "배색":        .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "공정량":      .ColWidth(12) = 900:            .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "정련":        .ColWidth(13) = 900:            .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "SETT":        .ColWidth(14) = 900:            .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "PEACH":       .ColWidth(15) = 900:            .ColAlignment(15) = flexAlignRightCenter
        .TextArray(16) = "CPB":         .ColWidth(16) = 900:            .ColAlignment(16) = flexAlignRightCenter
        .TextArray(17) = "염색":        .ColWidth(17) = 900:            .ColAlignment(17) = flexAlignRightCenter
        .TextArray(18) = "DRY":         .ColWidth(18) = 900:            .ColAlignment(18) = flexAlignRightCenter
        .TextArray(19) = "가공":        .ColWidth(19) = 900:            .ColAlignment(19) = flexAlignRightCenter
        .TextArray(20) = "검사":        .ColWidth(20) = 900:            .ColAlignment(20) = flexAlignRightCenter
        .TextArray(21) = "보류":        .ColWidth(21) = 900:            .ColAlignment(21) = flexAlignRightCenter
        .TextArray(22) = "검사":        .ColWidth(22) = 900:            .ColAlignment(22) = flexAlignRightCenter
        .TextArray(23) = "검사":        .ColWidth(23) = 900:            .ColAlignment(23) = flexAlignRightCenter
        .TextArray(24) = "출고량":      .ColWidth(24) = 1000:           .ColAlignment(24) = flexAlignRightCenter
        .TextArray(25) = "공정패턴코드": .ColWidth(25) = 0
        .TextArray(26) = "가공폭":       .ColWidth(26) = 0
        .TextArray(27) = "반입량":      .ColWidth(27) = 1000
        
        .TextArray(.Cols + 0) = " "
        .TextArray(.Cols + 1) = "거래처"
        .TextArray(.Cols + 2) = "품명"
        .TextArray(.Cols + 3) = "관리번호" & vbCrLf & "색  상  명"
        .TextArray(.Cols + 4) = "Order No." & vbCrLf & "색  상  명"
        .TextArray(.Cols + 5) = "축율"
        .TextArray(.Cols + 6) = "수주량"
        .TextArray(.Cols + 7) = "입고량"
        .TextArray(.Cols + 8) = "미계획량"
        .TextArray(.Cols + 9) = "계획량"
        .TextArray(.Cols + 10) = "대기량"
        .TextArray(.Cols + 11) = "배색량"
        .TextArray(.Cols + 12) = "공정량"
        .TextArray(.Cols + 13) = "정련"
        .TextArray(.Cols + 14) = "SETT"
        .TextArray(.Cols + 15) = "PEACH"
        .TextArray(.Cols + 16) = "CPB"
        .TextArray(.Cols + 17) = "염색"
        .TextArray(.Cols + 18) = "DRY"
        .TextArray(.Cols + 19) = "가공"
        .TextArray(.Cols + 20) = "검사"
        .TextArray(.Cols + 21) = "보류"
        .TextArray(.Cols + 22) = "합격"
        .TextArray(.Cols + 23) = "불합격"
        .TextArray(.Cols + 24) = "출고량"
        .TextArray(.Cols + 25) = "공정패턴코드"
        .TextArray(.Cols + 26) = "가공폭"
        .TextArray(.Cols + 27) = "반입량"

        .ColFormat(6) = "#,##0"
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        .ColFormat(11) = "#,##0"
        .ColFormat(12) = "#,##0"
        .ColFormat(13) = "#,##0"
        .ColFormat(14) = "#,##0"
        .ColFormat(15) = "#,##0"
        .ColFormat(16) = "#,##0"
        .ColFormat(17) = "#,##0"
        .ColFormat(18) = "#,##0"
        .ColFormat(19) = "#,##0"
        .ColFormat(20) = "#,##0"
        .ColFormat(21) = "#,##0"
        .ColFormat(22) = "#,##0"
        .ColFormat(23) = "#,##0"
        .ColFormat(24) = "#,##0"
        .ColFormat(27) = "#,##0"
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
        For i = 0 To 9
            .MergeCol(i) = True
        Next i
        
        For i = 12 To 21
            .MergeCol(i) = True
        Next i
        .MergeCol(24) = True
        .MergeCol(25) = True
        .MergeCol(26) = True
        .MergeCol(27) = True
       
        For i = 1 To .Cols - 1
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
        Next i
        
        For i = 13 To 21
            .ColHidden(i) = True
        Next i
        
        .ColHidden(15) = True
        .ColHidden(18) = True
        
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusSolid
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 0
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
    
    With grdStuffIN
        .Cols = 9
        Call SetVSFlexGrid(grdStuffIN)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1
        .RowHeight(0) = 400

        .TextArray(0) = " "
        .TextArray(1) = "입고일자":     .ColWidth(1) = 1000:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "입고구분":     .ColWidth(2) = 2000:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "입고순번":     .ColWidth(3) = 1000:             .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "제직처":       .ColWidth(4) = 3500:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "사종":         .ColWidth(5) = 2500:             .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "입고단위":     .ColWidth(6) = 1500:             .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "입고절수":     .ColWidth(7) = 1300:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "입고수량":     .ColWidth(8) = 1500:             .ColAlignment(8) = flexAlignRightCenter
        
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        
        .ColHidden(3) = True
        .WordWrap = False
        
        .Redraw = flexRDDirect
    End With
    
    
    '--- 공정대기
    '--- 그룹설정
    Call SetGridGroup(grdProcWait)
    With grdProcWait
        .Redraw = flexRDNone
        
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 1
        .Cols = 8 ' 17
        .RowHeight(0) = 400
        
        .TextArray(0) = "":                               .ColWidth(0) = 250   '1단계
        .TextArray(1) = "":                               .ColWidth(1) = 250   '2단계
        .TextArray(2) = "":                               .ColWidth(2) = 250   '3단계
        
        .TextArray(3) = "공정명":                         .ColAlignment(3) = flexAlignCenterCenter:
        .TextArray(4) = "색상명":                         .ColAlignment(4) = flexAlignLeftCenter:
        .TextArray(5) = "카드수" & vbCrLf & "카드번호":   .ColAlignment(5) = flexAlignLeftCenter:
        .TextArray(6) = "절수":                           .ColAlignment(6) = flexAlignRightCenter:
        .TextArray(7) = "수량":                           .ColAlignment(7) = flexAlignRightCenter:
        

        .ColWidth(3) = 1560
        .ColWidth(4) = 1600
        .ColWidth(5) = 1800
        .ColWidth(6) = 800
        .ColWidth(7) = 1020
        For i = 1 To .Cols - 1
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
        Next i
        .Redraw = flexRDDirect
    End With
    
    '--- Card 세부내역
    With grdProcWaitDetail
        .Redraw = flexRDNone
        .Cols = 8
        
        Call SetVSFlexGrid(grdProcWaitDetail)
        .Rows = 1
        
        .TextArray(0) = "순서":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColHidden(1) = True:       .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "완료여부":     .ColHidden(3) = True:       .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "실적일":       .ColWidth(4) = "1200":      .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "요구폭":       .ColWidth(5) = "700":       .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "지시사항":     .ColWidth(6) = "3000":      .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "수정":         .ColWidth(7) = "400":       .ColAlignment(7) = flexAlignCenterCenter
        
        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
        
    '--- 실적
 '   Call SetGridGroup(grdWorkResult)
    With grdWorkResult
        .Cols = 10

        .Redraw = flexRDDirect

        .Rows = 2
        .FixedRows = 2
        .FixedCols = 1
    
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250
        
        .TextArray(0) = " ":                        .ColWidth(0) = 200:                      .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = " ":                        .ColWidth(1) = 200:                      .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":                   .ColWidth(2) = 1000:                     .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "카드번호":                 .ColWidth(3) = 1300:                     .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "현수량":                   .ColWidth(4) = 800:                      .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "카드" & vbCr & "상태":     .ColWidth(5) = 600:                      .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "공정":                     .ColWidth(6) = 8300:                     .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "검사":                     .ColWidth(7) = 800:                      .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "검사":                     .ColWidth(8) = 800:                      .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "출고":                     .ColWidth(9) = 980:                      .ColAlignment(9) = flexAlignCenterCenter
        
        .TextArray(.Cols + 0) = " "
        .TextArray(.Cols + 1) = " "
        .TextArray(.Cols + 2) = "색상명"
        .TextArray(.Cols + 3) = "카드번호"
        .TextArray(.Cols + 4) = "현수량"
        .TextArray(.Cols + 5) = "카드" & vbCr & "상태"
        .TextArray(.Cols + 6) = "공정"
        .TextArray(.Cols + 7) = "합격"
        .TextArray(.Cols + 8) = "불량"
        .TextArray(.Cols + 9) = "출고"
        
        
        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarVertical
        .MergeCells = flexMergeFixedOnly
        
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        For i = 0 To 9
            .MergeCol(i) = True
        Next
        
        For i = .FixedCols To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .WordWrap = True
        .Redraw = flexRDDirect
    End With
    
    Call SetVSFlexGrid(grdProcWaitDetail)
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .ScrollBars = flexScrollBarNone
        .FixedRows = 0
        .Rows = 1
        .Cols = 4
        .ExtendLastCol = True
        
        .RowHeight(0) = 300
        .TextArray(0) = "합계":         .ColWidth(0) = 1000:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "":             .ColWidth(1) = 700:    .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "":             .ColWidth(2) = 900:    .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "":             .ColWidth(3) = 1700:    .ColAlignment(3) = flexAlignRightCenter
        .Redraw = flexRDDirect
    End With
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

Private Sub FillGridOrder()
    Dim oPlanInput As Pluslib2.CPlanInput
    Dim rs As ADODB.Recordset
    Dim dRS As New ADODB.Recordset
    Dim i%, nTop%, dSql_str$
    Dim nNoPlanQty#, nProceTotalQty#
    
    On Error GoTo ErrHandler
    
    dSql_str = " SELECT  WorkName = ISNULL( ( SELECT WorkName" & vbCr & _
               "                                FROM [mt_work] BB " & vbCr & _
               "                               WHERE AA.WorkID = BB.WorkID ), '' ) " & vbCr & _
               "       , WorkWidth = ISNULL( ( SELECT StuffWidth " & vbCr & _
               "                                 FROM [mt_Stuffwidth] BB " & vbCr & _
               "                                WHERE AA.WorkWidth = BB.StuffWidthID ), '' ) " & vbCr & _
               "       , AcptDate, DvlyDate, WorkDensity, PatternID " & vbCr & _
               "    FROM [Order] AA " & vbCr & _
               "    WHERE OrderID = '" & Trim(txtSearch) & "' "
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount = 1 Then
        txtAcptDate = MakeDate(DF_LONG, Trim(dRS!AcptDate))
        txtDvlyDate = MakeDate(DF_LONG, Trim(dRS!DvlyDate))
        txtWorkName = Trim(dRS!WorkName)
        txtWorkWidth = Trim(dRS!WorkWidth)
        txtWorkDensity = Trim(dRS!WorkDensity)
        txtPattern = GetPatternProc(dRS!PatternID)
    Else
        txtUseClss = ""
        txtExpectDate = ""
        txtCustom = ""
        txtThreadName = ""
    End If
    dRS.Close
    Set dRS = Nothing
    
    
    Set oPlanInput = New Pluslib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    Set rs = oPlanInput.GetOrderHistory(, , , , , , , IIf(optOrder(0).Value = True, 1, 2), txtSearch, 1, 0)
    
    Set oPlanInput = Nothing
        
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            nNoPlanQty = IIf(rs!UnitClss = "0", rs!ColorQty, CLng(rs!ColorQty / 0.9144)) * (1 + rs!ChunkRate / 100) - rs!InstQty  '미계획량
            nProceTotalQty = rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty + rs!셋팅Qty + _
                            rs!PeachQty + rs!C염색Qty + rs!염색Qty + rs!P염색Qty + rs!R수세Qty + rs!건조Qty + rs!가공Qty + _
                            rs!검사Qty + rs!PauseQty
                            
            If rs!OrderID <> MakeOrderID(.TextMatrix(nTop, 3), OM_REDUCE) Then
                .AddItem "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    rs!ChunkRate & vbTab & rs!OrderQty & vbTab & rs!StuffInQty & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & CheckNull(rs!PatternID) & vbTab & rs!WorkWidth & vbTab & rs!반입Qty
                
                
                Call DoFlexGridGroup(grdOrder, .Rows - 1, 1)
                Call GridCollapse(grdOrder, nTop)
                
                nTop = .Rows - 1
            End If
        
            .AddItem "" & vbTab & "" & vbTab & "" & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & rs!Color & vbTab & rs!Color & vbTab & _
                "" & vbTab & IIf(rs!ReWorkClss = "*", "", rs!ColorQty) & vbTab & " " & vbTab & IIf(rs!ReWorkClss = "*", "", nNoPlanQty) & vbTab & _
                rs!InstQty & vbTab & rs!InstQty - rs!배색TQty & vbTab & rs!배색TQty & vbTab & nProceTotalQty & vbTab & _
                rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty & vbTab & _
                rs!셋팅Qty & vbTab & rs!PeachQty & vbTab & rs!C염색Qty & vbTab & _
                rs!염색Qty + rs!P염색Qty + rs!R수세Qty & vbTab & rs!건조Qty & vbTab & rs!가공Qty & vbTab & _
                rs!검사Qty & vbTab & rs!PauseQty & vbTab & rs!PassQty & vbTab & rs!DefectQty & vbTab & rs!OutQty & vbTab & CheckNull(rs!PatternID) & vbTab & rs!WorkWidth & vbTab & rs!반입Qty
        
            .TextMatrix(nTop, 8) = CLng(.TextMatrix(nTop, 8)) + nNoPlanQty
            .TextMatrix(nTop, 9) = CLng(.TextMatrix(nTop, 9)) + rs!InstQty
            .TextMatrix(nTop, 10) = CLng(.TextMatrix(nTop, 10)) + rs!InstQty - rs!배색TQty
            .TextMatrix(nTop, 11) = CLng(.TextMatrix(nTop, 11)) + rs!배색TQty
            .TextMatrix(nTop, 12) = CLng(.TextMatrix(nTop, 12)) + nProceTotalQty
            .TextMatrix(nTop, 13) = CLng(.TextMatrix(nTop, 13)) + rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty
            .TextMatrix(nTop, 14) = CLng(.TextMatrix(nTop, 14)) + rs!셋팅Qty
            .TextMatrix(nTop, 15) = CLng(.TextMatrix(nTop, 15)) + rs!PeachQty
            .TextMatrix(nTop, 16) = CLng(.TextMatrix(nTop, 16)) + rs!C염색Qty
            .TextMatrix(nTop, 17) = CLng(.TextMatrix(nTop, 17)) + rs!염색Qty + rs!P염색Qty + rs!R수세Qty
            .TextMatrix(nTop, 18) = CLng(.TextMatrix(nTop, 18)) + rs!건조Qty
            .TextMatrix(nTop, 19) = CLng(.TextMatrix(nTop, 19)) + rs!가공Qty
            .TextMatrix(nTop, 20) = CLng(.TextMatrix(nTop, 20)) + rs!검사Qty
            .TextMatrix(nTop, 21) = CLng(.TextMatrix(nTop, 21)) + rs!PauseQty
            .TextMatrix(nTop, 22) = CLng(.TextMatrix(nTop, 22)) + rs!PassQty
            .TextMatrix(nTop, 23) = CLng(.TextMatrix(nTop, 23)) + rs!DefectQty
            .TextMatrix(nTop, 24) = CLng(.TextMatrix(nTop, 24)) + rs!OutQty
            .TextMatrix(nTop, 27) = CLng(.TextMatrix(nTop, 27)) + rs!반입Qty
            
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
            MsgBox LoadResString(203), vbInformation
        End If
        
        
''        For i = 3 To 9
''            .MergeCol(i) = True
''        Next i
''        .MergeCells = flexMergeRestrictColumns
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    Exit Sub

ErrHandler:
    Set oPlanInput = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOrderTotal.FillGridOrder", Err.Description)
End Sub

Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(iRow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2, 3
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        End Select
    End With
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

Private Sub FillGridStuffIN()
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    Dim dRollQty_int As Long, dQty_int As Long, dRec_int As Integer
    
    dRollQty_int = 0: dQty_int = 0: dRec_int = 0

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "GetStuffINRec"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, Trim(txtSearch.Text))

        Set dRS = .Execute
    End With
    
    Set adoCmd = Nothing
    With grdStuffIN
        .Rows = .FixedRows
        Do Until dRS.EOF
            .AddItem "" & vbTab & MakeDate(DF_LONG, dRS!StuffDate) & vbTab & dRS!StuffClssName & vbTab & dRS!StuffSeq & vbTab & _
            dRS!Custom & vbTab & dRS!ThreadName & vbTab & dRS!UnitClss & vbTab & dRS!TotRoll & vbTab & dRS!TotQty
            dRollQty_int = dRollQty_int + dRS!TotRoll
            dQty_int = dQty_int + dRS!TotQty
            dRec_int = dRec_int + 1
            dRS.MoveNext
        Loop
        dRS.Close
        Set dRS = Nothing
    End With
    
    grdTotal.TextArray(1) = Format(dRec_int, "#,##0 건")
    grdTotal.TextArray(2) = Format(dRollQty_int, "##,##0 절")
    grdTotal.TextArray(3) = Format(dQty_int, "###,##0 YDS")

End Sub
Private Sub FillgrdProcWaitDetail(ByVal CardID As String, ByVal SplitID As String)
    Dim adoCmd As ADODB.Command
    Dim dRS As New ADODB.Recordset
    Dim dSql_str As String
    
    dSql_str = " SELECT UseClss, ExpectDate, Custom, ThreadName " & vbCr & _
               "   FROM [Card] " & vbCr & _
               "  WHERE CardID = '" & Trim(CardID) & "' and SplitID = '" & SplitID & "' "
    
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount = 1 Then
        txtUseClss = Trim(dRS!UseClss)
        txtExpectDate = MakeDate(DF_LONG, Trim(dRS!ExpectDate))
        txtCustom = Trim(dRS!Custom)
        txtThreadName = Trim(dRS!ThreadName)
    Else
        txtUseClss = ""
        txtExpectDate = ""
        txtCustom = ""
        txtThreadName = ""
    End If
    dRS.Close
    Set dRS = Nothing
    
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OrdHistResult"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, CardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, SplitID)

        Set dRS = .Execute
    End With
    'PlanSeq ProcessID ProcessName    CompleteClss      ResultDate NeedWidth InstRemark
    Set adoCmd = Nothing
    With grdProcWaitDetail
        .Rows = .FixedRows
        Do Until dRS.EOF
            .AddItem dRS!PlanSeq & vbTab & dRS!ProcessID & vbTab & Trim(dRS!ProcessName) & vbTab & _
            Trim(dRS!CompleteClss) & vbTab & IIf(Trim(dRS!ResultDate) = "", "", MakeDate(DF_LONG, dRS!ResultDate)) & vbTab & _
            Trim(dRS!NeedWidth) & vbTab & Trim(dRS!InstRemark) & vbTab & IIf(dRS!ReWorkClss = "*", "■", "")
            dRS.MoveNext
        Loop
        dRS.Close
        Set dRS = Nothing
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub





Private Sub grdProcWait_Click()
    Dim dCardID As String
    Dim CardID As String, SplitID As String
    With grdProcWait
    End With
    
    With grdProcWait
        If .IsSubtotal(.Row) = True Then
            MsgBox "하위 내용을 선택하십시오", vbInformation
            Exit Sub
        End If
        dCardID = Replace(Trim(.TextMatrix(.Row, 5)), "-", "")
        
        Call FillgrdProcWaitDetail(Left(dCardID, 8), Mid(dCardID, 9))
    End With
End Sub

Private Sub FormClear()
    grdOrder.Rows = grdOrder.FixedRows
    grdStuffIN.Rows = grdStuffIN.FixedRows
    grdProcWait.Rows = grdProcWait.FixedRows
    grdProcWaitDetail.Rows = grdProcWaitDetail.FixedRows
    grdWorkResult.Rows = grdWorkResult.FixedRows
    
    txtUseClss = ""
    txtExpectDate = ""
    txtCustom = ""
    txtThreadName = ""
End Sub



Public Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtSearch)) > 0 And Len(Trim(txtSearch)) = 10 Then
            Call FormClear
            Call FillGridOrder
            Call FillGridStuffIN
            Call FillgrdProcWait
            Call FillgrdWorkResult
        Else
            MsgBox "해당 번호를 입력 하지 않았습니다", vbInformation, "Key 입력"
        End If
    End If
End Sub
