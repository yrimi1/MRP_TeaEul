VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcessWait 
   Caption         =   "공정대기현황"
   ClientHeight    =   9255
   ClientLeft      =   75
   ClientTop       =   750
   ClientWidth     =   15180
   Icon            =   "frmProcessWait.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15180
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   1635
      TabIndex        =   37
      Top             =   8625
      Width           =   1320
      _ExtentX        =   2328
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
      Begin VB.CheckBox chkCount 
         Caption         =   "인쇄 매수"
         Height          =   180
         Left            =   60
         TabIndex        =   38
         Top             =   60
         Width           =   1140
      End
   End
   Begin VB.TextBox txtCount 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   36
      Top             =   8610
      Width           =   945
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   30
      TabIndex        =   33
      Top             =   8640
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1058
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optPrint 
         Caption         =   "공정별 출력"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   345
         Width           =   1275
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "전체출력"
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   60
         Value           =   -1  'True
         Width           =   1320
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   4005
      TabIndex        =   32
      Top             =   8565
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab prTab 
      Height          =   8505
      Left            =   4020
      TabIndex        =   17
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   15002
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   600
      TabCaption(0)   =   "대기(Order별)"
      TabPicture(0)   =   "frmProcessWait.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdOrderWait"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "대기(작업단위별)"
      TabPicture(1)   =   "frmProcessWait.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdWorkWait"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "보류"
      TabPicture(2)   =   "frmProcessWait.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdHold"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "보류작성내역"
         Height          =   2025
         Left            =   -74895
         TabIndex        =   20
         Top             =   6450
         Width           =   10965
         Begin VB.TextBox txtOccuDate 
            Height          =   300
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   270
            Width           =   1935
         End
         Begin VB.TextBox txtOccuProc 
            Height          =   300
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtHoldReason 
            Height          =   975
            Left            =   1380
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   930
            Width           =   5085
         End
         Begin VB.TextBox txtHoldPersonID 
            Height          =   315
            Left            =   8070
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   210
            Width           =   1545
         End
         Begin VB.TextBox txtHoldSetDate 
            Height          =   315
            Left            =   8070
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   570
            Width           =   1545
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   51
            Left            =   180
            TabIndex        =   26
            Top             =   600
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
            Caption         =   "발생공정"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   55
            Left            =   180
            TabIndex        =   27
            Top             =   930
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
            Caption         =   "보류원인"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   3
            Left            =   6870
            TabIndex        =   28
            Top             =   210
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
            Index           =   4
            Left            =   180
            TabIndex        =   29
            Top             =   270
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
            Caption         =   "발생일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   0
            Left            =   6870
            TabIndex        =   30
            Top             =   570
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
            Caption         =   "작성일시"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOrderWait 
         Height          =   8040
         Left            =   -74940
         TabIndex        =   18
         Top             =   420
         Width           =   11010
         _cx             =   19420
         _cy             =   14182
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
         AllowUserResizing=   1
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
      Begin VSFlex7LCtl.VSFlexGrid grdWorkWait 
         Height          =   8040
         Left            =   60
         TabIndex        =   19
         Top             =   390
         Width           =   11010
         _cx             =   19420
         _cy             =   14182
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
         AllowUserResizing=   3
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
      Begin VSFlex7LCtl.VSFlexGrid grdHold 
         Height          =   6000
         Left            =   -74940
         TabIndex        =   31
         Top             =   390
         Width           =   11010
         _cx             =   19420
         _cy             =   10583
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
         AllowUserResizing=   1
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
   Begin VSFlex7LCtl.VSFlexGrid grdProcess 
      Height          =   5400
      Left            =   15
      TabIndex        =   3
      Top             =   1590
      Width           =   3960
      _cx             =   6985
      _cy             =   9525
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
   Begin VB.Frame fraSearch 
      Height          =   1695
      Left            =   30
      TabIndex        =   2
      Top             =   -90
      Width           =   3945
      Begin Threed.SSPanel SSPanel2 
         Height          =   375
         Left            =   75
         TabIndex        =   41
         Top             =   150
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   120
            Width           =   1125
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   1590
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   975
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Top             =   1320
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   735
         Left            =   3030
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         ToolTipText     =   "자료 저장"
         Top             =   150
         Width           =   840
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   7125
         TabIndex        =   5
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   615
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   45
            Width           =   1065
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   975
         Width           =   1230
         _ExtentX        =   2170
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
         Index           =   0
         Left            =   2850
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   975
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
         Left            =   75
         TabIndex        =   13
         Top             =   1320
         Width           =   1230
         _ExtentX        =   2170
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
            Width           =   1050
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1320
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
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13500
      TabIndex        =   1
      Top             =   8550
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
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   375
      Left            =   15
      TabIndex        =   16
      Top             =   6990
      Width           =   3960
      _cx             =   6985
      _cy             =   661
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
      ScrollBars      =   0
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
   Begin Threed.SSCommand cmdOrderDetail 
      Height          =   690
      Left            =   7710
      TabIndex        =   39
      Top             =   8550
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "수주상세"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdCardDetail 
      Height          =   690
      Left            =   9570
      TabIndex        =   40
      Top             =   8550
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "카드상세"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTube 
      Height          =   1110
      Left            =   15
      TabIndex        =   44
      Top             =   7380
      Width           =   3960
      _cx             =   6985
      _cy             =   1958
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
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
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
      ScrollBars      =   0
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
Attribute VB_Name = "frmProcessWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const REPORTFILE = "\Report\WaitCardList.rpt"

Private Type TParaType
    nCheckOrderID   As Integer
    sOrderID        As String
    nCheckOrderNo   As Integer
    sOrderNO        As String
    nCheckCutom     As Integer
    sCustomID       As String
    nCheckArticle   As Integer
    sArticleID      As String
End Type
Dim m_bloading As Boolean
Dim TParaType As TParaType

Private Sub chkCount_Click()
    If chkCount.Value = vbChecked Then
        txtCount.Enabled = True
    Else
        txtCount.Enabled = False
    End If
End Sub

Private Sub chkSearch_Click(Index As Integer)
    
    Select Case Index
        Case 0:
            If chkSearch(0).Value = vbChecked Then
                txtSearch(0).Enabled = True
            Else
                txtSearch(0).Enabled = False
                txtSearch(0).Text = ""
            End If
        Case 1:
            If chkSearch(1).Value = vbChecked Then
                cmdFind(0).Enabled = True
                txtSearch(1).Enabled = True
            Else
                cmdFind(0).Enabled = False
                txtSearch(1).Enabled = False
                txtSearch(1).Text = ""
            End If
        Case 2
            If chkSearch(2).Value = vbChecked Then
                cmdFind(1).Enabled = True
                txtSearch(2).Enabled = True
            Else
                cmdFind(1).Enabled = False
                txtSearch(2).Enabled = False
                txtSearch(2).Text = ""
            End If
    End Select
End Sub

Private Sub cmdCardDetail_Click()
    Dim sCardID As String
    If prTab.Tab = 1 Then
        With grdWorkWait
            If .Rows > .FixedRows Then
                sCardID = MakeCardID(.TextMatrix(.Row, 8), OM_REDUCE, "-")
                frmCardHistory.txtCard.Text = sCardID
                frmCardHistory.txtCard_KeyPress (vbKeyReturn)
            End If
        End With
    Else
        With grdHold
            If .Rows > .FixedRows Then
                sCardID = MakeCardID(.TextMatrix(.Row, 2), OM_REDUCE, "-")
                frmCardHistory.txtCard.Text = sCardID
                frmCardHistory.txtCard_KeyPress (vbKeyReturn)
            End If
        End With
    
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0             '[3] 거래처 코드
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
        Case 1             '[4] 품명
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End Select

End Sub


Private Sub cmdOrderDetail_Click()
    Dim sOrderID As String
    If prTab.Tab = 1 Then
        With grdWorkWait
            If .Rows > .FixedRows Then
                sOrderID = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
                frmOrderHistory.optOrder(0).Value = True
                frmOrderHistory.txtSearch.Text = sOrderID
                frmOrderHistory.txtSearch_KeyPress (vbKeyReturn)
            End If
        End With
        
        
    Else
        With grdHold
            If .Rows > .FixedRows Then
                sOrderID = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
                frmOrderHistory.optOrder(0).Value = True
                frmOrderHistory.txtSearch.Text = sOrderID
                frmOrderHistory.txtSearch_KeyPress (vbKeyReturn)
            End If
        End With
        
    End If
End Sub

Private Sub SetPrint()
    On Error GoTo ErrHandler
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim oProcess As PlusLib2.CProcess
    Dim nChkProcessID%, sProcessID$
    Dim i%, nCount%, bChkPrev As Boolean
    Dim nChkOrder%, sOrder$, nChkCustomID%, sCustomID$, nChkArticleID%, sArticleID$
    
    If grdProcess.Rows = grdProcess.FixedRows Then Exit Sub

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    nChkOrder = IIf(chkSearch(0).Value, 1, 0)
    sOrder = txtSearch(0)
    nChkCustomID = IIf(chkSearch(1).Value, 1, 0)
    sCustomID = txtSearch(1).Tag
    nChkArticleID = IIf(chkSearch(2).Value, 1, 0)
    sArticleID = txtSearch(2).Tag
    nChkProcessID = IIf(optPrint(0).Value = True, 0, 1)
    sProcessID = grdProcess.TextMatrix(grdProcess.Row, 4)
    
    If chkCount.Value = vbChecked Then
        If IsNumeric(txtCount) Then
            nCount = txtCount
        Else
            nCount = 1
        End If
    Else
        nCount = 1
    End If
    
    For i = 1 To nCount
        Set oProcess = New PlusLib2.CProcess
        oProcess.Connection = g_adoCon
    
        Set rs = oProcess.GetWaitCardList(nChkOrder, sOrder, nChkCustomID, sCustomID, nChkArticleID, sArticleID, nChkProcessID, sProcessID)
        Set oProcess = Nothing
      
        ReDim Preserve sParam(0)
        
        sParam(0) = "진 호 염 직 (주)"
        
        Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
    Next i
    
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oProcess = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub


Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String
    
    Call ColResize(grdHold, ES_REDUCE, 20)
    
    With grdHold
        .Redraw = flexRDBuffered

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        .ExtendLastCol = False

        .FontSize = 7
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "보류현황"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 2, 7, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .SheetBorder = vbBlack
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .PrintGrid "태을염직", True, 1, 100, 500

        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True

        .FontSize = 9
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    
    Call ColResize(grdHold, ES_EXPAND, 20)
    
End Sub

Private Sub cmdPrint_Click()
    If prTab.Tab = 2 Then
        Call FillGrdPrint
        
    Else
        Call SetPrint
    End If
End Sub

Private Sub cmdSearch_Click()
    grdOrderWait.Rows = grdOrderWait.FixedRows
    grdWorkWait.Rows = grdWorkWait.FixedRows
    Call ClearData
    Call FillGridOrder
    Call FillGridTube(True)
    Call FillGrdWaitHold
    Call grdProcess_Click
End Sub
Private Sub FillGridTube(Optional ByVal bTotal As Boolean, Optional ByVal sProcID As String)
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetCardTube(IIf(bTotal, 0, 1), _
                IIf(bTotal, "", sProcID), _
                IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
                IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
                IIf(chkSearch(0).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0), txtSearch(0))
    Set oCard = Nothing
    
    With grdTube
        .Redraw = flexRDNone
        If rs.RecordCount > 0 Then
            If bTotal Then
                .TextMatrix(2, 1) = Format(rs!Tube1, "#,###")
                .TextMatrix(2, 2) = Format(rs!Tube2, "#,###")
                .TextMatrix(2, 3) = Format(rs!Tube3, "#,###")
            Else
                .TextMatrix(1, 0) = rs!Process
                .TextMatrix(1, 1) = Format(rs!Tube1, "#,###")
                .TextMatrix(1, 2) = Format(rs!Tube2, "#,###")
                .TextMatrix(1, 3) = Format(rs!Tube3, "#,###")
            End If
        Else
            If bTotal Then
                .TextMatrix(2, 1) = ""
                .TextMatrix(2, 2) = ""
                .TextMatrix(2, 3) = ""
            Else
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
    
    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    
    grdTube.Redraw = flexRDDirect
    
    Call ErrorBox(Err.Number, "Card.FillGridTube", Err.Description)
End Sub
Private Sub FillGridOrder()
    Dim rs As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Dim lNowRow&, lNowSum&, i%
    Dim sOrderID As String
    'Dim TParaType As TParaType
    Dim nTotRoll As Long, nTotCard As Integer, nTotQty As Long
    
    On Error GoTo ErrHandler

    '------ Parameter 넘겨줄 값 Move

    With TParaType
        If chkSearch(0).Value = vbChecked Then
            If optOrder(0).Value = True Then  'Order NO
                .nCheckOrderID = 0
                .sOrderID = ""
                
                .nCheckOrderNo = 1
                .sOrderNO = txtSearch(0).Text
            Else
                .nCheckOrderID = 1
                .sOrderID = txtSearch(0).Text
                
                .nCheckOrderNo = 0
                .sOrderNO = ""
            End If
        Else
            .nCheckOrderID = 0
            .sOrderID = ""
            .nCheckOrderNo = 0
            .sOrderNO = ""
        End If
        .nCheckCutom = IIf(chkSearch(1) = vbChecked, 1, 0)
        .sCustomID = Trim(txtSearch(1).Tag)
        .nCheckArticle = IIf(chkSearch(2) = vbChecked, 1, 0)
        .sArticleID = Trim(txtSearch(2).Tag)
    End With
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sDraftOrder"
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaType.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TParaType.sOrderNO)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckCutom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sArticleID)
        Set rs = .Execute
    End With
    Set adoCmd = Nothing
    
    
    nTotRoll = 0
    nTotCard = 0
    nTotQty = 0
            
    '---- Recordset의 데이터를 Grid에 나타낸다.
    With grdProcess
        .Redraw = flexRDNone
        .Rows = .FixedRows
            
        Do Until rs.EOF
            nTotRoll = nTotRoll + rs!Roll_EA
            nTotCard = nTotCard + rs!Card_EA
            nTotQty = nTotQty + rs!Qty_EA
            
            .AddItem rs!Process & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & rs!Card_EA & vbTab & rs!Roll_EA & vbTab & rs!Qty_EA & vbTab & rs!waitprocid
            i = i + 1
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If
            .RowHeight(.Rows - 1) = 350

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
        Else
            .HighLight = flexHighlightNever
                    
        End If
        .Redraw = flexRDDirect
    End With
    
    grdTotal.TextArray(1) = Format(nTotCard, "#,##0")
    grdTotal.TextArray(2) = Format(nTotRoll, "#,##0")
    grdTotal.TextArray(3) = Format(nTotQty, "#,##0")
    
    If grdProcess.Rows > grdProcess.FixedRows Then
        grdProcess.Row = grdProcess.FixedRows
    Else
        MsgBox LoadResString(203), vbInformation
    End If
    
    m_bloading = False
    
    Exit Sub
ErrHandler:
    Set rs = Nothing
    
    grdProcess.Redraw = flexRDDirect
    
    Call ErrorBox(Err.Number, "frmProcessWaitTEMP.FillGridOrder", Err.Description)
End Sub

Private Sub Form_Load()
    Dim i%
    
    PlusMDI.pnlMenu.Visible = False
    Call SetOperate(Me)
    
    Me.Move 0, 0, 15300, 9660
    
    prTab.Tab = 1 '대기(작업단위별)을 맨 위로
    
    Call InitGrid
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)    '---거래처
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)    '---품명
    
    cmdFind(0).Enabled = False
    cmdFind(1).Enabled = False
    
    If prTab.Tab = 0 Then
        cmdOrderDetail.Visible = False
        cmdCardDetail.Visible = False
    Else
        cmdOrderDetail.Visible = True
        cmdCardDetail.Visible = True
        
    End If
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
Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, irow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(irow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(irow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, irow, 0, irow, .Cols - 1) = &HE0E0E0
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2, 3
            .Cell(flexcpBackColor, irow, 0, irow, .Cols - 1) = &HE0E0E0
        End Select
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Sub FillGrdWaitWork(ByVal dProcessID As String, ByVal ReWorkClss As String)
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim iTop(2) As Integer
    Dim II%, JJ%
    Dim dColor As String

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sWorkProc"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, dProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaType.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TParaType.sOrderNO)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckCutom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ReWorkClss)
        
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing
    
  '  Call SetVSFlexGrid(grdWorkWait)
    
    With grdWorkWait
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            .Rows = 1
        End If
        
        II = 1
        dColor = "1"
        Do Until rs.EOF
            
            If Trim(.TextMatrix(.Rows - 1, 1)) <> Trim(rs!WorkUnitId) And (.Rows <> .FixedRows) Then
                .AddItem " "
                .RowHidden(.Rows - 1) = True
                II = 1
                dColor = dColor & ", " & CStr(.Rows)
            End If
            
            .AddItem "" & vbTab & Trim(rs!WorkUnitId) & vbTab & II & vbTab & MakeOrderID(Trim(rs!OrderID), OM_EXPAND) & vbTab & _
                    Trim(rs!OrderNo) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & _
                    Trim(rs!Color) & vbTab & MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID) & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!BatJaNO & vbTab & Trim(rs!Procss) & vbTab & Trim(rs!UseClss)
            .RowHeight(.Rows - 1) = 400
            
            II = II + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Cell(flexcpFontSize, .FixedRows, 11, .Rows - 1, 11) = 8
            .Row = .FixedRows
        End If
    End With
    
    Dim dWorkUnitID As Variant
    
    dWorkUnitID = Split(dColor, ",")
    
    If UBound(dWorkUnitID) > 1 Then
        With grdWorkWait
            For II = 0 To UBound(dWorkUnitID) Step 2
                If II = UBound(dWorkUnitID) Then
                     JJ = .Rows - 1
                Else
                     JJ = dWorkUnitID(II + 1) - 1
                End If
                .Cell(flexcpBackColor, dWorkUnitID(II), 0, JJ, .Cols - 1) = &HE0E0E0
            Next II
            
        End With
    End If
    
    Dim iCount As Integer
    With grdWorkWait
        For iCount = .FixedRows To .Rows - 1
            If .TextMatrix(iCount, 13) = "보류" Then
                .Cell(flexcpBackColor, iCount, 8, iCount, 8) = vbRed
                .Cell(flexcpForeColor, iCount, 8, iCount, 8) = vbWhite
            ElseIf .TextMatrix(iCount, 13) = "작업" Then
                .Cell(flexcpBackColor, iCount, 8, iCount, 8) = vbBlue
                .Cell(flexcpForeColor, iCount, 8, iCount, 8) = vbWhite
            End If
        Next iCount
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
    
    
 '.Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
    
'    Dim dWorkUnitClss As String
'    Dim dColor As Long
'
'    dColor = &HE0E0E0
'
'    With grdWorkWait
'        For II = .FixedRows To .Rows - 1
'            If II = .FixedRows Then
'                dWorkUnitClss = Trim(.TextMatrix(II, 1))
'            End If
'
'            If .TextMatrix(II, 1) = dWorkUnitClss Then
'                .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
'            End If
'
'        Next II
'    End With

End Sub
' ---  대기(Order별)
Sub FillGrdWaitOrder(ByVal dProcessID As String, ByVal ReWorkClss As String)
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim iTop(2) As Integer

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sOrder"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, dProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaType.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TParaType.sOrderNO)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckCutom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ReWorkClss)
        
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing
    
    Call SetVSFlexGrid(grdOrderWait)
    With grdOrderWait
        .Redraw = flexRDNone
        .ExplorerBar = flexExNone
        
        If .Rows > .FixedRows Then
            .Rows = 1
        End If
            
        Do Until rs.EOF
            
            .AddItem " " & vbTab & " " & vbTab & " " & vbTab & _
                    Trim(rs!kCustom) & vbTab & MakeOrderID(Trim(rs!OrderID), OM_EXPAND) & vbTab & Trim(rs!OrderNo) & vbTab & Trim(rs!Article) & vbTab & _
                    Trim(rs!Color) & vbTab & MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID) & vbTab & rs!Roll & vbTab & rs!Qty
            .RowHeight(.Rows - 1) = 350
            
            If rs!UseClss = "보류" Then
                .Cell(flexcpBackColor, .Rows - 1, 8, .Rows - 1, 8) = vbRed
                .Cell(flexcpForeColor, .Rows - 1, 8, .Rows - 1, 8) = vbWhite
            ElseIf rs!UseClss = "작업" Then
                .Cell(flexcpBackColor, .Rows - 1, 8, .Rows - 1, 8) = vbBlue
                .Cell(flexcpForeColor, .Rows - 1, 8, .Rows - 1, 8) = vbWhite
            End If
            
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
    End With
    Dim II%, JJ%
    
    With grdOrderWait
        For II = .Rows - 1 To .FixedRows Step -1
            For JJ = 7 To 3 Step -1
                If Trim(.TextMatrix(II, JJ)) = Trim(.TextMatrix(II - 1, JJ)) Then
                    .TextMatrix(II, JJ) = ""
                End If
                
            Next JJ

        Next II
    End With
    grdOrderWait.Redraw = flexRDDirect
    
    '-- Group의 단계설정
    With grdOrderWait
        For II = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(II, 3)) <> "" Then
                Call DoFlexGridGroup(grdOrderWait, II, 1)
            ElseIf Trim(.TextMatrix(II, 4)) <> "" Then
                Call DoFlexGridGroup(grdOrderWait, II, 2)
            End If
            
        Next II
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
Private Sub InitGrid()
    '--- main Grid
    Dim II%
    
    Call SetVSFlexGrid(grdProcess)
    With grdProcess
        .Redraw = False
        .Cols = 6
            
        .TextArray(0) = "공정명":         .ColWidth(0) = 1300:    .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "수정":           .ColWidth(1) = 300:     .ColAlignment(1) = flexAlignCenterCenter
        
        .TextArray(2) = "카드수":         .ColWidth(2) = 600:     .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "절수":           .ColWidth(3) = 700:     .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "수량":           .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "ProcID":         .ColWidth(5) = 0
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .Redraw = True
    End With
    
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 4
        .ExtendLastCol = True
        
        .RowHeight(0) = 300
        .TextArray(0) = "합계":         .ColWidth(0) = 1000:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "":             .ColWidth(1) = 700:    .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "":             .ColWidth(2) = 900:    .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "":             .ColWidth(3) = 1700:    .ColAlignment(3) = flexAlignRightCenter
        .RowHeight(0) = 450
        .Redraw = flexRDDirect
    End With
    
    With grdTube
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 3
        .Cols = 4
        .ExtendLastCol = True
        
        .TextMatrix(0, 0) = "Tube":     .ColWidth(0) = 1000:    .ColAlignment(0) = flexAlignCenterCenter:   .FixedAlignment(0) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "1":        .ColWidth(1) = 950:    .ColAlignment(1) = flexAlignRightCenter:   .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "2":        .ColWidth(2) = 950:    .ColAlignment(2) = flexAlignRightCenter:   .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(0, 3) = "3":        .ColWidth(3) = 950:    .ColAlignment(3) = flexAlignRightCenter:   .FixedAlignment(3) = flexAlignCenterCenter
        
        .TextMatrix(2, 0) = "전체"
        
        .RowHeight(0) = 350: .RowHeight(1) = 350: .RowHeight(2) = 350:
        
        .Cell(flexcpFontBold, 1, 1, 1, .Cols - 1) = True
        .Cell(flexcpFontBold, 2, 1, 2, .Cols - 1) = True
        .Cell(flexcpFontSize, 1, 1, 1, .Cols - 1) = 14
        .Cell(flexcpFontSize, 2, 1, 2, .Cols - 1) = 14
        .Redraw = flexRDDirect
    End With
    
'    '--- 대기(Order별)
'    Call SetVSFlexGrid(grdOrderWait)
    Call SetGridGroup(grdOrderWait)
    With grdOrderWait
        .Redraw = False
        .Cols = 14
        .FixedRows = 1
        .RowHeight(0) = 450
        
        .TextArray(1) = " ":                .ColWidth(1) = 200:      .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = " ":                .ColWidth(2) = 200:      .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "거래처명":         .ColWidth(3) = 1900:     .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "관리번호":         .ColWidth(4) = 1300:     .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNO":          .ColWidth(5) = 1300:     .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "품명":             .ColWidth(6) = 1300:     .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "색상명":           .ColWidth(7) = 1200:     .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "카드번호":           .ColWidth(8) = 1300:     .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "절수":             .ColWidth(9) = 900:      .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "수량":            .ColWidth(10) = 900:     .ColAlignment(10) = flexAlignCenterCenter
        
        
        .TextArray(11) = "거래처명"
        .TextArray(12) = "Orderid"
        .TextArray(13) = "color"
        
        .ColHidden(0) = True
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True
        
        
        .Redraw = flexRDDirect
    End With


'
'    '--- 대기(작업단위별)
    Call SetVSFlexGrid(grdWorkWait)
    With grdWorkWait
      '  .Redraw = False
        .Cols = 14
        .Rows = 1
        .FixedRows = 1
        
        .RowHeight(0) = 450
        
        .TextArray(1) = "작업단위ID":                     .ColWidth(1) = 0:       .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "밧자" & vbCrLf & "순위":         .ColWidth(2) = 400:     .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "관리번호":                       .ColWidth(3) = 1200:    .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "OrderNO":                        .ColWidth(4) = 1300:    .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "거래처명":                       .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "품명":                           .ColWidth(6) = 2400:    .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "색상명":                         .ColWidth(7) = 1300:    .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "카드번호":                       .ColWidth(8) = 1400:    .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "절수":                           .ColWidth(9) = 500:     .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "수량":                          .ColWidth(10) = 500:    .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(11) = "실물" & vbCrLf & "밧자기":      .ColWidth(11) = 800:    .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "공정진행":                      .ColWidth(12) = 1600:   .ColAlignment(12) = flexAlignLeftCenter
        .TextArray(13) = "카드" & vbCrLf & "상태":        .ColWidth(13) = 600:    .ColAlignment(13) = flexAlignCenterCenter
        
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(4) = True
        
        For II = 0 To .Cols - 1
            .FixedAlignment(II) = flexAlignCenterCenter
        Next II
        
        .ExtendLastCol = True
        .FrozenCols = 4
        
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
        
    End With
'
''    '--- 보류
    Call SetVSFlexGrid(grdHold)
    With grdHold
        .Rows = 4
        .FixedRows = 4
        .Cols = 10
        .RowHeight(0) = 450
        
        .TextMatrix(3, 0) = ""
        .TextMatrix(3, 1) = "공정명":         .ColWidth(1) = 1000:    .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "카드번호":       .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(3, 3) = "관리번호":      .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "OrderNO":       .ColWidth(4) = 1300:    .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "업체명":        .ColWidth(5) = 1800:    .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "품명":          .ColWidth(6) = 1600:    .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(3, 7) = "색상명":        .ColWidth(7) = 1200:    .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(3, 8) = "절수":          .ColWidth(8) = 500:     .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(3, 9) = "수량":          .ColWidth(9) = 800:     .ColAlignment(9) = flexAlignRightCenter

        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .RowHeight(3) = 400
        .Redraw = flexRDDirect
    End With
End Sub

'--- 보류
Sub FillGrdWaitHold()
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sHold"
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaType.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TParaType.sOrderNO)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckCutom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaType.sArticleID)
        
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing
    

    With grdHold
        .Redraw = flexRDNone
        .ExplorerBar = flexExSort
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows - .FixedRows + 1) & vbTab & rs!Process & vbTab & Trim(rs!CardID) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    Trim(rs!OrderNo) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & Trim(rs!ColorName) & vbTab & _
                            rs!Roll & vbTab & rs!Qty
            .RowHeight(.Rows - 1) = 350
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub


Sub SetHoldDetail(ByVal CardID As String, ByVal SplitID As String)

    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_GetHoldNonProc"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, CardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, SplitID)
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing

    If rs.RecordCount > 0 Then
        txtOccuDate.Text = Format$(rs!OccuDate, "####-##-##")
        txtOccuProc.Text = Trim(rs!OccuProc)
        txtHoldReason.Text = Trim(rs!HoldReason)
        txtHoldPersonID.Text = Trim(rs!PersonName)
        txtHoldSetDate.Text = Format$(rs!SetDate, "yyyy-mm-dd HH:MM")
    End If
    rs.Close
    Set rs = Nothing
    
End Sub
Sub ClearData()
    txtOccuDate.Text = ""
    txtOccuProc.Text = ""
    txtHoldReason.Text = ""
    txtHoldPersonID.Text = ""
    txtHoldSetDate.Text = ""

End Sub


Private Sub grdHold_Click()
    Dim CardID As String
    With grdHold
        If .Rows > .FixedRows Then
            CardID = Replace(grdHold.TextMatrix(grdHold.Row, 2), "-", "")
           Call SetHoldDetail(Left(CardID, 8), Mid(CardID, 9))
        End If
    End With
End Sub

Private Sub grdHold_RowColChange()
    Dim CardID As Variant
    With grdHold
        If .Rows > .FixedRows Then
            CardID = Split(grdHold.TextMatrix(grdHold.Row, 2), "-")
            Call SetHoldDetail(CardID(0), CardID(1))
        End If
    End With

End Sub

Private Sub grdOrderWait_DblClick()
    With grdOrderWait
        If .Row < .FixedRows Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With

End Sub

Private Sub grdProcess_Click()
    Dim dProcessID As String, ReWorkClss As String
    
    With grdProcess
        dProcessID = GetProcessID(Trim(.TextMatrix(.Row, 0)))
        ReWorkClss = IIf(Trim(.TextMatrix(.Row, 1)) = "", "", "*")
    End With
    
    Call ShowData(dProcessID, ReWorkClss)
    Call FillGridTube(False, dProcessID)
End Sub

'Private Sub grdProcess_RowColChange()
'    Dim dProcessID As String
'
'    dProcessID = GetProcessID(Trim(grdProcess.TextMatrix(grdProcess.Row, 0)))
'    Call ShowData(dProcessID)
'End Sub

Sub ShowData(ByVal dProcessID As String, ByVal ReWorkClss As String)
    Call ClearData
    Call FillGrdWaitOrder(dProcessID, ReWorkClss)
    Call FillGrdWaitWork(dProcessID, ReWorkClss)
End Sub

Private Sub optOrder_Click(Index As Integer)
    chkSearch(0).Caption = optOrder(Index).Caption
End Sub



Private Sub prTab_Click(PreviousTab As Integer)
    If prTab.Tab = 0 Then
        cmdOrderDetail.Visible = False
        cmdCardDetail.Visible = False
    Else
        cmdOrderDetail.Visible = True
        cmdCardDetail.Visible = True
        
    End If
End Sub


Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
                Call MoveFocus(KeyAscii)
            End If
        Case 2
            If KeyAscii = vbKeyReturn Then
                Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
                Call MoveFocus(KeyAscii)
            End If
    End Select
End Sub

'Private Sub txtSearch_LostFocus(Index As Integer)
'    Select Case Index
'        Case 1
'            If Len(txtSearch(Index)) > 0 Then
'                Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
'                Call NextFocus
'            End If
'        Case 2
'            If Len(txtSearch(Index)) Then
'                Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
'                Call NextFocus
'            End If
'    End Select
'
'End Sub
