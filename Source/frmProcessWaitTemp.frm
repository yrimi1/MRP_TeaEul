VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcessWaitTemp 
   Caption         =   "공정대기현황"
   ClientHeight    =   9315
   ClientLeft      =   2295
   ClientTop       =   2115
   ClientWidth     =   15270
   Icon            =   "frmProcessWaitTemp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15270
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3960
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   8700
      Width           =   1605
   End
   Begin VSFlex7LCtl.VSFlexGrid grdProcess 
      Height          =   6165
      Left            =   15
      TabIndex        =   3
      Top             =   1710
      Width           =   3960
      _cx             =   6985
      _cy             =   10874
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
   Begin VB.Frame fraSearch 
      Height          =   1695
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   3945
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   13
         Top             =   615
         Width           =   1545
      End
      Begin VB.Frame fraOrder 
         Height          =   810
         Left            =   90
         TabIndex        =   9
         Top             =   120
         Width           =   1305
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   525
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   1290
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   750
         Left            =   3090
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         ToolTipText     =   "자료 저장"
         Top             =   180
         Width           =   780
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   7125
         TabIndex        =   12
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
         Left            =   1440
         TabIndex        =   14
         Top             =   210
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   45
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   17
         Top             =   960
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   18
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   3030
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   1305
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
         Begin VB.CheckBox chkSearch 
            Caption         =   "품     명"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1140
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3030
         TabIndex        =   22
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
      Left            =   13590
      TabIndex        =   1
      Top             =   8580
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8505
      Left            =   4020
      TabIndex        =   4
      Top             =   120
      Width           =   32520
      _ExtentX        =   57362
      _ExtentY        =   15002
      _Version        =   393216
      TabsPerRow      =   8
      TabHeight       =   600
      TabMaxWidth     =   4410
      TabCaption(0)   =   "  대기(Order별)  "
      TabPicture(0)   =   "frmProcessWaitTemp.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdOrderWait"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "  대기(작업단위별)  "
      TabPicture(1)   =   "frmProcessWaitTemp.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdWorkWait"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "보  류  "
      TabPicture(2)   =   "frmProcessWaitTemp.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdHold"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "보류작성내역"
         Height          =   2025
         Left            =   -74910
         TabIndex        =   24
         Top             =   6390
         Width           =   10965
         Begin VB.TextBox txtHoldSetDate 
            Height          =   315
            Left            =   8070
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   570
            Width           =   1545
         End
         Begin VB.TextBox txtHoldPersonID 
            Height          =   315
            Left            =   8070
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   210
            Width           =   1545
         End
         Begin VB.TextBox txtHoldReason 
            Height          =   975
            Left            =   1380
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   930
            Width           =   5085
         End
         Begin VB.TextBox txtOccuProc 
            Height          =   300
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtOccuDate 
            Height          =   300
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   270
            Width           =   1935
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   51
            Left            =   180
            TabIndex        =   29
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
            TabIndex        =   30
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
            TabIndex        =   31
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
            TabIndex        =   32
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
            TabIndex        =   34
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
      Begin VSFlex7LCtl.VSFlexGrid grdHold 
         Height          =   5910
         Left            =   -74925
         TabIndex        =   7
         Top             =   420
         Width           =   11130
         _cx             =   19632
         _cy             =   10425
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
      Begin VSFlex7LCtl.VSFlexGrid grdWorkWait 
         Height          =   8040
         Left            =   -74940
         TabIndex        =   6
         Top             =   420
         Width           =   11130
         _cx             =   19632
         _cy             =   14182
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
      Begin VSFlex7LCtl.VSFlexGrid grdOrderWait 
         Height          =   8040
         Left            =   90
         TabIndex        =   5
         Top             =   420
         Width           =   11130
         _cx             =   19632
         _cy             =   14182
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
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   705
      Left            =   15
      TabIndex        =   23
      Top             =   7905
      Width           =   3960
      _cx             =   6985
      _cy             =   1244
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
Attribute VB_Name = "frmProcessWaitTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TParaType
    nCheckOrderID   As Integer
    sOrderID        As String
    nCheckOrderNo   As Integer
    sOrderNo        As String
    nCheckCutom     As Integer
    sCustomID       As String
    nCheckArticle   As Integer
    sArticleID      As String
End Type
Dim m_bLoading As Boolean

Private Sub chkSearch_Click(Index As Integer)
    
    Select Case Index
        Case 0:
            If chkSearch(1).Value = vbChecked Then
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

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0             '[3] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(1))
        Case 1             '[4] 품명
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If

End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub
Private Sub FillGridOrder()
    Dim rs As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Dim lNowRow&, lNowSum&, i%
    Dim sOrderID As String
    Dim TParaType As TParaType
    Dim nTotRoll As Integer, nTotCard As Integer, nTotQty As Integer
    
'    On Error GoTo ErrHandler


    '------ Parameter 넘겨줄 값 Move

    With TParaType
        If chkSearch(0).Value = vbChecked Then
            If optOrder(0).Value = True Then  'Order NO
                .nCheckOrderID = 0
                .sOrderID = ""
                
                .nCheckOrderNo = 1
                .sOrderNo = txtSearch(0).Text
            Else
                .nCheckOrderID = 1
                .sOrderID = txtSearch(0).Text
                
                .nCheckOrderNo = 0
                .sOrderNo = ""
            End If
        Else
            .nCheckOrderID = 0
            .sOrderID = ""
            .nCheckOrderNo = 0
            .sOrderNo = ""
        End If
        .nCheckCutom = IIf(chkSearch(1) = vbChecked, 1, 0)
        .sCustomID = Trim(txtSearch(1).Text)
        .nCheckArticle = IIf(chkSearch(2) = vbChecked, 1, 0)
        .sArticleID = Trim(txtSearch(2).Text)
    End With
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sDraftOrder"
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaType.sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaType.nCheckOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TParaType.sOrderNo)
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
            
            .AddItem rs!Process & vbTab & rs!Card_EA & vbTab & rs!Roll_EA & vbTab & rs!Qty_EA
            i = i + 1
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If
            .RowHeight(.Rows - 1) = 300

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
    
    grdTotal.TextArray(1) = Format(nTotCard, "#,##0 장")
    grdTotal.TextArray(2) = Format(nTotRoll, "#,##0 절")
    grdTotal.TextArray(3) = Format(nTotQty, "#,##0 YDS")
    
    If grdProcess.Rows > grdProcess.FixedRows Then
        grdProcess.Row = grdProcess.FixedRows
    End If
    
    m_bLoading = False
    
    Exit Sub
ErrHandler:
    Set rs = Nothing
    
    grdProcess.Redraw = flexRDDirect
    
    Call ErrorBox(Err.Number, "frmProcessWaitTEMP.FillGridOrder", Err.Description)
End Sub

Private Sub Form_Load()
    Dim i%
    
    PlusMDI.pnlMenu.Visible = False
    
    Me.Move 0, 0, 15360, 9840
    
    Call InitGrid
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)    '---거래처
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)    '---품명
    
    cmdFind(0).Enabled = False
    cmdFind(1).Enabled = False
    
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
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        End Select
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Sub FillGrdWaitWork(ByVal dProcessID As String)
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
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing
    
    With grdWorkWait
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            .Rows = 1
        End If
        
''        .TextArray(1) = "작업단위ID":                    .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter
''        .TextArray(2) = "단위" & vbCrLf & "순위":        .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignCenterCenter
''        .TextArray(3) = "관리번호":       .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignCenterCenter
''        .TextArray(4) = "OrderNO":        .ColWidth(4) = 1300:    .ColAlignment(4) = flexAlignCenterCenter
''        .TextArray(5) = "거래처명":       .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignCenterCenter
''        .TextArray(6) = "품명":           .ColWidth(6) = 1500:    .ColAlignment(6) = flexAlignCenterCenter
''        .TextArray(7) = "색상명":         .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignCenterCenter
''        .TextArray(8) = "CardID":         .ColWidth(8) = 1000:    .ColAlignment(8) = flexAlignCenterCenter
''        .TextArray(9) = "절수":           .ColWidth(9) = 600:     .ColAlignment(9) = flexAlignCenterCenter
''        .TextArray(10) = "수량":          .ColWidth(10) = 600:    .ColAlignment(10) = flexAlignCenterCenter
''        .TextArray(11) = "공정진행":      .ColWidth(11) = 1000:   .ColAlignment(11) = flexAlignCenterCenter
        II = 1
        dColor = "1"
        Do Until rs.EOF
            
            If Trim(.TextMatrix(.Rows - 1, 1)) <> Trim(rs!WorkUnitID) And (.Rows <> .FixedRows) Then
                .AddItem " "
                .RowHidden(.Rows - 1) = True
                II = 1
                dColor = dColor & ", " & CStr(.Rows)
            End If
            
            .AddItem CStr(.Rows) & vbTab & Trim(rs!WorkUnitID) & vbTab & II & vbTab & MakeOrderID(Trim(rs!OrderID), OM_EXPAND) & vbTab & _
                    Trim(rs!OrderNo) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & _
                    Trim(rs!Color) & vbTab & Trim(rs!CardID) & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & Trim(rs!Procss)
            .RowHeight(.Rows - 1) = 500
            II = II + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
    End With
    
    Dim dWorkUnitID As Variant
    
    dWorkUnitID = Split(dColor, ",")
    
    If UBound(dWorkUnitID) > 0 Then
        With grdWorkWait
            For II = 0 To UBound(dWorkUnitID) Step 2
            
                .Cell(flexcpBackColor, dWorkUnitID(II), 0, dWorkUnitID(II + 1) - 1, .Cols - 1) = &HE0E0E0
            
            Next II
            
        End With
    End If
    
    
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
Sub FillGrdWaitOrder(ByVal dProcessID As String)
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim iTop(2) As Integer

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sOrder"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, dProcessID)

    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing
    
    With grdOrderWait
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            .Rows = 1
        End If
        
''        .TextArray(3) = "거래처명":         .ColWidth(3) = 1900:     .ColAlignment(3) = flexAlignCenterCenter
''        .TextArray(4) = "관리번호":         .ColWidth(4) = 1300:     .ColAlignment(4) = flexAlignCenterCenter
''        .TextArray(5) = "OrderNO":          .ColWidth(5) = 1300:     .ColAlignment(5) = flexAlignCenterCenter
''        .TextArray(6) = "품명":             .ColWidth(6) = 1300:     .ColAlignment(6) = flexAlignCenterCenter
''        .TextArray(7) = "색상명":           .ColWidth(7) = 1200:     .ColAlignment(7) = flexAlignCenterCenter
''        .TextArray(8) = "CardID":           .ColWidth(8) = 1200:     .ColAlignment(8) = flexAlignCenterCenter
''        .TextArray(9) = "절수":             .ColWidth(9) = 900:      .ColAlignment(9) = flexAlignCenterCenter
''        .TextArray(10) = "수량":            .ColWidth(10) = 900:     .ColAlignment(10) = flexAlignCenterCenter

''        .TextArray(11) = "거래처명"
''        .TextArray(12) = "Orderid"
''        .TextArray(13) = "color"
        
            
        Do Until rs.EOF
            
            .AddItem " " & vbTab & " " & vbTab & " " & vbTab & _
                    Trim(rs!kCustom) & vbTab & MakeOrderID(Trim(rs!OrderID), OM_EXPAND) & vbTab & Trim(rs!OrderNo) & vbTab & Trim(rs!Article) & vbTab & _
                    Trim(rs!Color) & vbTab & Trim(rs!CardID) & vbTab & rs!Roll & vbTab & rs!Qty
            
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
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
    
    With grdOrderWait
        For II = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(II, 3)) <> "" Then
                Call DoFlexGridGroup(grdOrderWait, II, 1)
                Call GridCollapse(grdOrderWait, II)
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
        .Cols = 4
            
        .TextArray(0) = "공정명":         .ColWidth(0) = 1500:    .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "카드수":         .ColWidth(1) = 700:     .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "절수":           .ColWidth(2) = 700:     .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "수량":           .ColWidth(3) = 1200:    .ColAlignment(3) = flexAlignRightCenter
        
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
        
        .Redraw = flexRDDirect
    End With
    
    
'    '--- 대기(Order별)
    Call SetVSFlexGrid(grdOrderWait)
    With grdOrderWait
        .Redraw = False
        .Cols = 14
        .FixedRows = 1
        .RowHeight(0) = 700

        .TextArray(1) = " ":                .ColWidth(1) = 200:      .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = " ":                .ColWidth(2) = 200:      .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "거래처명":         .ColWidth(3) = 1900:     .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "관리번호":         .ColWidth(4) = 1300:     .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNO":          .ColWidth(5) = 1300:     .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "품명":             .ColWidth(6) = 1300:     .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "색상명":           .ColWidth(7) = 1200:     .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "CardID":           .ColWidth(8) = 1200:     .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "절수":             .ColWidth(9) = 900:      .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "수량":            .ColWidth(10) = 900:     .ColAlignment(10) = flexAlignCenterCenter
        
        
        .TextArray(11) = "거래처명"
        .TextArray(12) = "Orderid"
        .TextArray(13) = "color"
        
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True
        
        
        .Redraw = flexRDDirect
    End With


'
'    '--- 대기(작업단위별)
    Call SetVSFlexGrid(grdWorkWait)
    With grdWorkWait
        .Redraw = False
        .Cols = 12
        .FixedRows = 1
         
        .TextArray(0) = ""
        .TextArray(1) = "작업단위ID":                    .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "단위" & vbCrLf & "순위":        .ColWidth(2) = 600:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "관리번호":       .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "OrderNO":        .ColWidth(4) = 1300:    .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "거래처명":       .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "품명":           .ColWidth(6) = 1300:    .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "색상명":         .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "CardID":         .ColWidth(8) = 1000:    .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "절수":           .ColWidth(9) = 500:     .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "수량":          .ColWidth(10) = 500:    .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(11) = "공정진행":      .ColWidth(11) = 1000:   .ColAlignment(11) = flexAlignLeftCenter
        
        .ColHidden(1) = True
        .ColHidden(4) = True
        .ColHidden(5) = True
        .Redraw = flexRDDirect
    End With
'
''    '--- 보류
    Call SetVSFlexGrid(grdHold)
    With grdHold
        .Redraw = False
        .Cols = 9
            
        .TextArray(0) = ""
        .TextArray(1) = "관리번호":       .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "OrderNO":        .ColWidth(2) = 1300:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "업체명":         .ColWidth(3) = 1900:    .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "품명":           .ColWidth(4) = 1400:    .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "카드번호":       .ColWidth(5) = 1200:    .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "색상명":         .ColWidth(6) = 1400:    .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "절수":           .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "수량":           .ColWidth(8) = 1000:    .ColAlignment(8) = flexAlignCenterCenter

        .Redraw = True
    End With
End Sub
'--- 보류
Sub FillGrdWaitHold(ByVal dProcessID As String)
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessWait_sHold"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, dProcessID)
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing

    With grdHold
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            .Rows = 1
        End If
            
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & Trim(rs!OrderNo) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & _
                            Trim(rs!CardID) & vbTab & Trim(rs!ColorName) & vbTab & _
                            rs!Roll & vbTab & rs!Qty
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
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
Private Sub grdHold_Click()
   ' grdHold.TextMatrix (grdHold.Row, 5)
   Dim CardID As Variant
   
   CardID = Split(grdHold.TextMatrix(grdHold.Row, 5), "-")
   Call SetHoldDetail(CardID(0), CardID(1))
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
    Dim dProcessID As String
    
    dProcessID = GetProcessID(Trim(grdProcess.TextMatrix(grdProcess.Row, 0)))
    Call ShowData(dProcessID)
End Sub

Private Sub grdProcess_RowColChange()
    Dim dProcessID As String
    
    dProcessID = GetProcessID(Trim(grdProcess.TextMatrix(grdProcess.Row, 0)))
'    Call ShowData(dProcessID)
End Sub

Sub ShowData(ByVal dProcessID As String)
    Call FillGrdWaitOrder(dProcessID)
    Call FillGrdWaitWork(dProcessID)
    Call FillGrdWaitHold(dProcessID)
End Sub

Private Sub optOrder_Click(Index As Integer)
    chkSearch(0).Caption = optOrder(Index).Caption
End Sub

