VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRecipe 
   ClientHeight    =   9255
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmRecipe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin Threed.SSPanel spDyeAux 
      Height          =   4185
      Left            =   3030
      TabIndex        =   91
      Top             =   1110
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   7382
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame faDyeAux 
         Caption         =   "염료선택"
         ForeColor       =   &H00000000&
         Height          =   3315
         Left            =   60
         TabIndex        =   95
         Top             =   60
         Width           =   4755
         Begin VSFlex7LCtl.VSFlexGrid grdDyeAuxFind 
            Height          =   2985
            Left            =   90
            TabIndex        =   96
            Top             =   240
            Width           =   4590
            _cx             =   8096
            _cy             =   5265
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
      Begin VB.TextBox txtDyeAuxClss 
         Height          =   285
         Left            =   90
         TabIndex        =   94
         Top             =   3510
         Visible         =   0   'False
         Width           =   915
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   555
         Left            =   3630
         TabIndex        =   93
         Top             =   3510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
         _Version        =   196609
         Caption         =   "닫기"
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   555
         Left            =   2310
         TabIndex        =   92
         Top             =   3510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
         _Version        =   196609
         Caption         =   "선택"
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   765
      Left            =   0
      TabIndex        =   58
      Top             =   8475
      Width           =   1470
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   195
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   495
         Width           =   1200
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   8385
      TabIndex        =   51
      Top             =   8535
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      저장(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8385
      TabIndex        =   37
      Top             =   8535
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Enabled         =   0   'False
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   10185
      TabIndex        =   38
      Top             =   8535
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab tabForm 
      Height          =   8460
      Left            =   0
      TabIndex        =   13
      Top             =   15
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   14923
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   2
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmRecipe.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdFind(1)"
      Tab(0).Control(1)=   "pnlCaption(14)"
      Tab(0).Control(2)=   "dtpDate(1)"
      Tab(0).Control(3)=   "dtpDate(0)"
      Tab(0).Control(4)=   "pnlCaption(9)"
      Tab(0).Control(5)=   "cmdFind(0)"
      Tab(0).Control(6)=   "pnlCaption(4)"
      Tab(0).Control(7)=   "pnlCaption(0)"
      Tab(0).Control(8)=   "grdRecipe"
      Tab(0).Control(9)=   "cmdOperate(2)"
      Tab(0).Control(10)=   "cmdOperate(1)"
      Tab(0).Control(11)=   "cmdOperate(0)"
      Tab(0).Control(12)=   "txtSearch(2)"
      Tab(0).Control(13)=   "txtSearch(0)"
      Tab(0).Control(14)=   "txtSearch(1)"
      Tab(0).Control(15)=   "SSFrame1"
      Tab(0).Control(16)=   "SSFrame2"
      Tab(0).Control(17)=   "cmdSearch(0)"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmRecipe.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "pnlCaption(15)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdFind(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dtpDateI(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dtpDateI(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "pnlCaption(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdFind(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "pnlCaption(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "pnlCaption(10)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "grdOrder"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "pnlEdit"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "pnlProgress"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtSearchI(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtSearchI(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtSearchI(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdSearch(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdCancel"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "grdColor"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   1950
         Left            =   3435
         TabIndex        =   85
         Top             =   825
         Width           =   8325
         _cx             =   14684
         _cy             =   3440
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
      Begin VB.CommandButton cmdCancel 
         Caption         =   "작성취소"
         Height          =   735
         Left            =   10770
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   84
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   1020
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   735
         Index           =   1
         Left            =   9630
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   83
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   1020
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Index           =   0
         Left            =   -66585
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   82
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   930
      End
      Begin VB.TextBox txtSearchI 
         Height          =   300
         Index           =   1
         Left            =   4410
         TabIndex        =   77
         Top             =   465
         Width           =   1800
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2325
         Left            =   -69930
         TabIndex        =   68
         Top             =   6060
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   4101
         _Version        =   196609
         Caption         =   "[처방내역]"
         Begin VSFlex7LCtl.VSFlexGrid grdShowDyeAux 
            Height          =   1995
            Index           =   0
            Left            =   45
            TabIndex        =   71
            Top             =   270
            Width           =   3300
            _cx             =   5821
            _cy             =   3519
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
         Begin VSFlex7LCtl.VSFlexGrid grdShowDyeAux 
            Height          =   1995
            Index           =   1
            Left            =   3375
            TabIndex        =   70
            Top             =   270
            Width           =   3300
            _cx             =   5821
            _cy             =   3519
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
      Begin Threed.SSFrame SSFrame1 
         Height          =   2325
         Left            =   -74925
         TabIndex        =   67
         Top             =   6060
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   4101
         _Version        =   196609
         Caption         =   " [변경내역] "
         Begin VSFlex7LCtl.VSFlexGrid grdHistory 
            Height          =   1995
            Left            =   60
            TabIndex        =   69
            Top             =   270
            Width           =   4815
            _cx             =   8493
            _cy             =   3519
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
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   -70605
         TabIndex        =   63
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox txtSearchI 
         Height          =   300
         Index           =   0
         Left            =   4410
         TabIndex        =   18
         Top             =   105
         Width           =   1800
      End
      Begin VB.TextBox txtSearchI 
         Height          =   300
         Index           =   2
         Left            =   6690
         TabIndex        =   21
         Top             =   450
         Width           =   1800
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   0
         Left            =   -70605
         TabIndex        =   6
         Top             =   150
         Width           =   1635
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   2
         Left            =   -68550
         TabIndex        =   9
         Top             =   495
         Width           =   1905
      End
      Begin Threed.SSPanel pnlProgress 
         Height          =   870
         Left            =   -8055
         TabIndex        =   39
         Top             =   6240
         Visible         =   0   'False
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1535
         _Version        =   196609
         Alignment       =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin MSComctlLib.ProgressBar proProgress 
            Height          =   390
            Left            =   90
            TabIndex        =   40
            Top             =   375
            Width           =   10890
            _ExtentX        =   19209
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblCount 
            AutoSize        =   -1  'True
            Caption         =   "180"
            Height          =   180
            Left            =   195
            TabIndex        =   41
            Top             =   120
            Width           =   270
         End
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   -65580
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   10
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   -64785
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   11
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin Threed.SSPanel pnlEdit 
         Height          =   5595
         Left            =   3420
         TabIndex        =   42
         Top             =   2805
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   9869
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
            Height          =   2805
            Index           =   0
            Left            =   60
            TabIndex        =   33
            Top             =   2715
            Width           =   4020
            _cx             =   7091
            _cy             =   4948
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
         Begin Threed.SSPanel pnlInfo 
            Height          =   2160
            Left            =   60
            TabIndex        =   43
            Top             =   75
            Width           =   8250
            _ExtentX        =   14552
            _ExtentY        =   3810
            _Version        =   196609
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txtColorQty 
               Height          =   300
               Left            =   3240
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   780
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txtQty 
               Height          =   300
               Left            =   3240
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   450
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txtBox 
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   6
               Left            =   5580
               TabIndex        =   87
               Top             =   1140
               Width           =   1800
            End
            Begin VB.TextBox txtBox 
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   5
               Left            =   1365
               TabIndex        =   75
               Top             =   1140
               Width           =   1800
            End
            Begin VB.TextBox txtRemark 
               Height          =   570
               Left            =   1350
               MultiLine       =   -1  'True
               TabIndex        =   73
               Top             =   1485
               Width           =   6345
            End
            Begin VB.TextBox txtBox 
               BackColor       =   &H00FFC0C0&
               Height          =   300
               Index           =   4
               Left            =   1365
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   465
               Width           =   1800
            End
            Begin VB.TextBox txtTemp 
               Height          =   300
               Left            =   3240
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txtBox 
               Height          =   300
               Index           =   3
               Left            =   5580
               TabIndex        =   29
               Top             =   465
               Width           =   1800
            End
            Begin MSComCtl2.DTPicker dtpRecipe 
               Height          =   300
               Left            =   5580
               TabIndex        =   28
               Top             =   135
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   529
               _Version        =   393216
               Format          =   113639425
               CurrentDate     =   37112
            End
            Begin VB.TextBox txtBox 
               BackColor       =   &H00FFFFC0&
               Height          =   300
               Index           =   2
               Left            =   5580
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   810
               Width           =   1800
            End
            Begin VB.TextBox txtBox 
               BackColor       =   &H00FFC0C0&
               Height          =   300
               Index           =   1
               Left            =   1365
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   795
               Width           =   1800
            End
            Begin VB.TextBox txtBox 
               BackColor       =   &H00FFC0C0&
               Height          =   300
               Index           =   0
               Left            =   1365
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   135
               Width           =   1800
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   135
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "관리번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   2
               Left            =   120
               TabIndex        =   45
               Top             =   795
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "색      상"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   5
               Left            =   4335
               TabIndex        =   46
               Top             =   810
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "처방전번호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   6
               Left            =   4335
               TabIndex        =   47
               Top             =   135
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "처방 일자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   7
               Left            =   4335
               TabIndex        =   48
               Top             =   465
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "처  방  자"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSCommand cmdFind 
               Height          =   300
               Index           =   4
               Left            =   7410
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   465
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   529
               _Version        =   196609
               ButtonStyle     =   3
               Outline         =   0   'False
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   8
               Left            =   7230
               TabIndex        =   50
               Top             =   1440
               Visible         =   0   'False
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   196609
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.CheckBox chkRework 
                  Caption         =   "재 처 방"
                  Height          =   180
                  Left            =   150
                  TabIndex        =   25
                  Top             =   60
                  Width           =   1515
               End
            End
            Begin Threed.SSCommand cmdFind 
               Height          =   300
               Index           =   5
               Left            =   7410
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   810
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
               TabIndex        =   61
               Top             =   465
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "품      명"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   16
               Left            =   120
               TabIndex        =   72
               Top             =   1485
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "특기사항"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin VB.TextBox txtModify 
               Height          =   300
               Left            =   6945
               TabIndex        =   74
               Top             =   1440
               Visible         =   0   'False
               Width           =   255
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   3
               Left            =   120
               TabIndex        =   76
               Top             =   1140
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "g / yd"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   17
               Left            =   4335
               TabIndex        =   88
               Top             =   1140
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "축      율"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   390
            Index           =   0
            Left            =   2925
            TabIndex        =   32
            Top             =   2295
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "염료삭제(&W)"
         End
         Begin Threed.SSCommand cmdAddNew 
            Height          =   390
            Index           =   0
            Left            =   1680
            TabIndex        =   31
            Top             =   2295
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "염료추가(&Q)"
         End
         Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
            Height          =   2805
            Index           =   1
            Left            =   4290
            TabIndex        =   36
            Top             =   2715
            Width           =   4020
            _cx             =   7091
            _cy             =   4948
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
         Begin Threed.SSCommand cmdDelete 
            Height          =   390
            Index           =   1
            Left            =   7155
            TabIndex        =   35
            Top             =   2295
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "조제삭제(&R)"
         End
         Begin Threed.SSCommand cmdAddNew 
            Height          =   390
            Index           =   1
            Left            =   5910
            TabIndex        =   34
            Top             =   2295
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "조제추가(&E)"
         End
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   -63990
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   12
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VSFlex7LCtl.VSFlexGrid grdRecipe 
         Height          =   5055
         Left            =   -74955
         TabIndex        =   0
         Top             =   900
         Width           =   11775
         _cx             =   20770
         _cy             =   8916
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
      Begin MSFlexGridLib.MSFlexGrid grdOrder 
         Height          =   7605
         Left            =   30
         TabIndex        =   22
         Top             =   810
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   13414
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   -74295
         TabIndex        =   52
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "처방일자"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   86
            Top             =   60
            Value           =   1  '확인
            Width           =   1065
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   10
         Left            =   705
         TabIndex        =   53
         Top             =   105
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearchI 
            Caption         =   "수주일자"
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   14
            Top             =   60
            Width           =   1020
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   -68550
         TabIndex        =   54
         Top             =   165
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "Order No"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1380
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   -68940
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   165
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
         Left            =   -71715
         TabIndex        =   55
         Top             =   135
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거래처"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   5
            Top             =   60
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   -73095
         TabIndex        =   3
         Top             =   135
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113639425
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   -73095
         TabIndex        =   4
         Top             =   495
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113639425
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   11
         Left            =   6705
         TabIndex        =   56
         Top             =   105
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearchI 
            Caption         =   "Order No"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1290
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   6225
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   105
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   12
         Left            =   3285
         TabIndex        =   57
         Top             =   105
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearchI 
            Caption         =   "거래처"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   17
            Top             =   60
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpDateI 
         Height          =   300
         Index           =   0
         Left            =   1935
         TabIndex        =   15
         Top             =   105
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113639425
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDateI 
         Height          =   300
         Index           =   1
         Left            =   1935
         TabIndex        =   16
         Top             =   465
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113639425
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   14
         Left            =   -71715
         TabIndex        =   64
         Top             =   480
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품 명"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   65
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   -68940
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   495
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   3
         Left            =   6225
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   465
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   15
         Left            =   3270
         TabIndex        =   79
         Top             =   465
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearchI 
            Caption         =   "품명"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   80
            Top             =   60
            Width           =   975
         End
      End
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   8835
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   3585
      TabIndex        =   81
      Top             =   8640
      Visible         =   0   'False
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   900
      _Version        =   196609
      BackColor       =   65535
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "frmRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\Recipe.rpt"
'Private Const REPORTFILE1 = "\Report\RecipeList.rpt"

Private Const LIMIT_ROW1 = 25
Private Const LIMIT_ROW2 = 25
Private Const LIMIT_ROW3 = 5
Private Const LIMIT_ROW4 = 11
Private Const LIMIT_ROW5 = 7
Private Const LIMIT_WIDTH1 = 1380
Private Const LIMIT_WIDTH2 = 1635
Private Const LIMIT_WIDTH3 = 1965
Private Const LIMIT_WIDTH4 = 2085
Private Const LIMIT_WIDTH5 = 1890

Private m_sFlag         As String
Private m_nSelected     As Integer
Private m_bloading      As Boolean
Private m_bSortForward  As Boolean
Private m_sOrderID      As String
Private m_sColorID      As String
Private m_nRecipeSeq    As Integer
Private m_nModifySeq    As Integer
Private m_bSaved        As Boolean

Private Type DyeRecord
    sDyeID     As String * 2
    sDyeSeq    As String * 2
    sDye       As String * 30
    sDyeRate   As String * 9
    sTankNo   As String * 2
End Type

Private Type AuxRecord
    sAuxID     As String * 2
    sAuxSeq    As String * 2
    sAux       As String * 30
    sAuxRate   As String * 9
    sTankNo    As String * 2
End Type




Private Sub cmdCancel_Click()
    txtSearchI(2) = ""
    chkSearchI(2).Value = vbUnchecked
    txtQty = ""
    tabForm.Tab = 0
    
End Sub


Private Sub Form_Activate()
    m_bloading = False
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11985, 9660
    
    If PlusMDI.pnlMenu.Visible = False Then
        PlusMDI.pnlMenu.Visible = True
    End If
    spDyeAux.Visible = False
    Call SetOperate(Me)

    pnlEdit.Enabled = True
    dtpDate(0) = Now
    dtpDate(1) = Now
    dtpDateI(0) = Now
    dtpDateI(1) = Now
    cmdSave.MousePointer = ssCustom
    cmdSave.Picture = LoadResPicture("SAVE", vbResIcon)
    cmdSearch(0).Picture = LoadResPicture("SEARCH", vbResIcon)
    cmdSearch(1).Picture = LoadResPicture("SEARCH", vbResIcon)
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(3).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(4).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(5).Picture = LoadResPicture("FIND", vbResIcon)
    cmdCancel.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdSave.MousePointer = ssCustom
    cmdSave.Picture = LoadResPicture("SAVE", vbResIcon)
    cmdSave.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    Call InitGrid
    Call ClearData
    
    txtQty = ""
    
    txtSearch(0).Enabled = False
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
    cmdFind(0).Enabled = False
    cmdFind(1).Enabled = False

    dtpDateI(0).Enabled = False
    dtpDateI(1).Enabled = False
    txtSearchI(0).Enabled = False
    txtSearchI(1).Enabled = False
    txtSearchI(2).Enabled = False
    cmdFind(2).Enabled = False
    cmdFind(3).Enabled = False
        
    m_bloading = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting(LoadResString(100), Me.Name, "Custom", IIf(chkSearch(0) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Order", IIf(chkSearch(1) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "DateI", IIf(chkSearchI(0) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "CustomI", IIf(chkSearchI(1) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "OrderI", IIf(chkSearchI(2) = vbChecked, "1", "0"))
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 3 Then
        If chkSearch(Index) Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else
        If chkSearch(Index) Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 0 Then
                cmdFind(0).Enabled = True
            ElseIf Index = 1 Then
                cmdFind(1).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 0 Then
                cmdFind(0).Enabled = False
            ElseIf Index = 1 Then
                cmdFind(1).Enabled = False
            End If
        End If
    End If
End Sub

Private Sub chkSearchI_Click(Index As Integer)
    If chkSearchI(Index) Then
        If Index = 3 Then
            dtpDateI(0).Enabled = True
            dtpDateI(1).Enabled = True
            dtpDateI(0).SetFocus
        Else
            txtSearchI(Index).Enabled = True
            txtSearchI(Index).SetFocus
            If Index = 0 Then
                cmdFind(2).Enabled = True
            ElseIf Index = 1 Then
                cmdFind(3).Enabled = True
            End If
        End If
    Else
        If Index = 3 Then
            dtpDateI(0).Enabled = False
            dtpDateI(1).Enabled = False
        Else
            txtSearchI(Index).Enabled = False
        
            cmdSearch(1).SetFocus
            If Index = 0 Then
                cmdFind(2).Enabled = False
            ElseIf Index = 1 Then
                cmdFind(3).Enabled = False
            End If
        End If
    End If
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then       ' 금일
        If tabForm.Tab = 0 Then
            dtpDate(0) = Date
            dtpDate(1) = Date
        Else
            dtpDateI(0) = Date
            dtpDateI(1) = Date
        End If
    ElseIf Index = 1 Then   ' 금월
        If tabForm.Tab = 0 Then
            dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
            dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
        Else
            dtpDateI(0) = DateSerial(Year(Date), Month(Date), 1)
            dtpDateI(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
        End If
    End If

'    cmdSearch.SetFocus
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub dtpDateI_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub cmdFind_Click(Index As Integer)
    ' 조회 - 거래처
    If Index = 0 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(0))
            
    ' 조회 - 품명
    ElseIf Index = 1 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
    
    ' 입력 -거래처
    ElseIf Index = 2 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearchI(0))
                    
    ' 입력 품명
    ElseIf Index = 3 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearchI(1))
        
    ' 입력 - 처방자
    ElseIf Index = 4 Then
        Call ReturnCode(LG_PERSON, , False, txtBox(3))
        
    ' 입력 - 처방전 번호
    ElseIf Index = 5 Then
        Dim sRecipeNO$

        sRecipeNO = InputBox("처방전 번호를 입력하십시오.")
        If Len(sRecipeNO) <= 0 Then Exit Sub

        Call GetRecipeOne(sRecipeNO)
    End If
End Sub

Public Sub cmdSearch_Click(Index As Integer)
    Select Case Index
        Case 0
            Call FillGridRecipe
        Case 1
            Call ClearData
            Call FillGridOrder
    End Select
End Sub



Private Sub grdDyeAuxFind_Click()
    Dim II%, nMax%
    nMax = 0
    With grdDyeAuxFind

        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
        '    .TextMatrix(.Row, 5) = .Row
            For II = .FixedRows To .Rows - 1
                If val(.ValueMatrix(II, 5)) <> 0 Then
                    If val(.ValueMatrix(II, 5)) > nMax Then
                        nMax = val(.ValueMatrix(II, 5))
                    End If
                End If
            Next II
            .TextMatrix(.Row, 5) = nMax + 1
        Else
            .TextMatrix(.Row, 5) = 0
            
        End If

    End With
End Sub

Private Sub grdHistory_RowColChange()
    Dim sOrderID$, nOrderSeq%
    Dim nRecipeSeq%, nModifySeq%
    
    If grdRecipe.Rows = grdRecipe.FixedRows Then Exit Sub
    
    If grdHistory.Rows = grdHistory.FixedRows Then Exit Sub
    
    With grdRecipe
        sOrderID = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
        nOrderSeq = .TextMatrix(.Row, 5)
    End With
    
    With grdHistory
        nRecipeSeq% = CInt(.TextMatrix(.Row, 2))
        nModifySeq = CInt(.TextMatrix(.Row, 3))
    End With

    Call ShowDyeAuxData(sOrderID, nOrderSeq, nRecipeSeq%, nModifySeq)

    
End Sub



Private Sub grdOrder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdOrder
        If .Rows = .FixedRows Or .MouseRow < 0 Or .MouseRow >= .FixedRows Then Exit Sub

        Call SortGrid(grdOrder, .MouseCol, m_bSortForward)
        m_bSortForward = Not m_bSortForward
    End With
End Sub

Private Sub grdRecipe_AfterSort(ByVal Col As Long, Order As Integer)
    Call grdRecipe_RowColChange
End Sub

Private Sub grdRecipe_DblClick()
    With grdRecipe
        If .MouseRow < .FixedRows Then Exit Sub

        txtQty = ""
        txtQty = .TextMatrix(.MouseRow, 19)
        .Row = .MouseRow
        
    End With
    
    
    Call cmdOperate_Click(ID_UPDATE)
End Sub

Private Sub grdRecipe_RowColChange()
    Dim sOrderID$, nOrderSeq%
    Dim nRecipeSeq%, nModifySeq%
        
    If m_bloading Then Exit Sub

    With grdRecipe
        If .Rows > .FixedRows Then
        
            sOrderID = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
            nOrderSeq = .TextMatrix(.Row, 5)
            nRecipeSeq = CInt(.TextMatrix(.Row, 9))
            nModifySeq = CInt(.TextMatrix(.Row, 10))
    
            Call ShowDyeAuxHistory(sOrderID, nOrderSeq, nRecipeSeq)
            
            Call ShowDyeAuxData(sOrderID, nOrderSeq, nRecipeSeq, nModifySeq)
            .SetFocus
        Else
            grdHistory.Rows = grdHistory.FixedRows
            grdShowDyeAux(0).Rows = grdShowDyeAux(0).FixedRows
            grdShowDyeAux(1).Rows = grdShowDyeAux(1).FixedRows
            
        End If
        
    End With
End Sub

Private Sub grdOrder_RowColChange()
    If m_bloading Then Exit Sub

    Call FillGridColor
End Sub

Private Sub grdColor_DblClick()
'    chkRework.SetFocus
End Sub

Private Sub grdColor_RowColChange()
    Call ShowSelOrder
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows
End Sub


Private Sub SSCommand1_Click()
    Dim II%
    With grdDyeAuxFind
        .Select 1, 5
        .Sort = flexSortGenericAscending
        For II = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, II, 1) = flexChecked Then
                grdDyeAux(txtDyeAuxClss).AddItem ""
                grdDyeAux(txtDyeAuxClss).TextMatrix(grdDyeAux(txtDyeAuxClss).Rows - 1, 1) = .TextMatrix(II, 2)      '염료(조제)명
                grdDyeAux(txtDyeAuxClss).TextMatrix(grdDyeAux(txtDyeAuxClss).Rows - 1, 5) = .TextMatrix(II, 2)      '염료(조제)명
                grdDyeAux(txtDyeAuxClss).TextMatrix(grdDyeAux(txtDyeAuxClss).Rows - 1, 4) = .TextMatrix(II, 3)      '염료(조제) 코드
''                grdDyeAux(txtDyeAuxClss).Cell(flexcpPicture, grdDyeAux(txtDyeAuxClss).Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
''                grdDyeAux(txtDyeAuxClss).Cell(flexcpPictureAlignment, grdDyeAux(txtDyeAuxClss).Rows - 1, 2) = flexPicAlignCenterCenter

            End If
        Next II
        .Redraw = flexRDDirect
    End With
    spDyeAux.Visible = False
End Sub

Private Sub SSCommand2_Click()
    spDyeAux.Visible = False
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
        End If
    End If
End Sub

Private Sub txtSearchI_GotFocus(Index As Integer)
    Call GotFocusText(txtSearchI(Index))
End Sub

Private Sub chkSearchI_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub txtSearchI_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearchI(0))
        ElseIf Index = 1 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearchI(1))
        
        End If
    End If
    
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim i%

    Select Case Index
        Case ID_ADDNEW
            tabForm.Tab = 1

            Call ClearData
            chkSearchI(2) = vbUnchecked
            m_sFlag = ID_ADDNEW

            cmdSearch(1).SetFocus
        Case ID_UPDATE
        
            With grdRecipe
               
                If .Rows = .FixedRows Then Exit Sub
            End With
            
            tabForm.Tab = 1

            chkSearchI(0) = vbUnchecked
            chkSearchI(1) = vbUnchecked
            chkSearchI(2) = vbChecked
            txtSearchI(2) = MakeOrderID(grdRecipe.TextMatrix(grdRecipe.Row, 3), OM_REDUCE)
            optOrder(1) = vbChecked
            Call cmdSearch_Click(1)

            With grdColor
                For i = .FixedRows To .Rows - 1
                    If grdRecipe.TextMatrix(grdRecipe.Row, 5) = .TextMatrix(i, 1) Then
                        .Row = i
                        .Col = .FixedCols
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next i
            End With

            Call ShowData
            grdColor.Enabled = False
            m_sFlag = ID_UPDATE
            cmdFind(2).Visible = False
            pnlMsg.Caption = LoadResString(122)

            txtBox(2).SetFocus
        Case ID_DELETE
            If grdRecipe.Rows = grdRecipe.FixedRows Then Exit Sub

            If Not QuestionBox(LoadResString(201)) Then Exit Sub

            If DeleteData() Then Call FillGridRecipe
    End Select
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
    Call GotFocusText(txtBox(Index))
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = vbKeyReturn Then Call ReturnCode(LG_PERSON, , False, txtBox(3))
    
    KeyAscii = KeyPress(txtBox(Index), KeyAscii)
End Sub

Private Sub chkRework_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub dtpRecipe_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub cmdAddNew_Click(Index As Integer)
'    Dim nType As Integer
    
    spDyeAux.Visible = True
'    If Index = 0 Then
'        nType = 1
'    Else
'        nType = 0
'    End If
    
    Call FillDyeAuxFind(IIf(Index = 0, 1, 0))
    
    With grdDyeAux(Index)
    '    .Rows = .Rows + 1
        
        Call ChangeScrollDyeAux(Index)

'        .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
'        .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter
        .SetFocus
        .Select .Rows - 1, 1
    End With
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    With grdDyeAux(Index)
        If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub

        .RemoveItem .Row

        cmdSave.SetFocus
    End With
End Sub

''Private Sub grdDyeAux_Click(Index As Integer)
''    With grdDyeAux(Index)
''        If .MouseRow < .FixedRows Or .MouseRow > .Rows - 1 Or .MouseCol <> 2 Then Exit Sub
''
''        Dim Row%
''        Row = .MouseRow
''        txtTemp = .TextMatrix(Row, 1)
''
''        spDyeAux.Visible = True
''        Call FillDyeAuxFind(Index)
''
'''        If ReturnCode(IIf(Index = 0, LG_DYE, LG_AUX), , False, txtTemp) Then
'''            .TextMatrix(Row, 1) = txtTemp
'''            .TextMatrix(Row, 5) = txtTemp
'''            .TextMatrix(Row, 4) = txtTemp.Tag
'''        End If
''    End With
''End Sub

Private Sub grdDyeAux_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdDyeAux(Index)
        Select Case Col
            Case 2
                Cancel = True
            Case 3
                If Len(.TextMatrix(Row, Col)) = 0 Then .TextMatrix(Row, Col) = ""
                .Cell(flexcpText, Row, Col) = Format(.TextMatrix(Row, Col), "###0.000000")
        End Select
    End With
End Sub

Private Sub grdDyeAux_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col <> 1 Or KeyAscii <> vbKeyReturn Then Exit Sub

    With grdDyeAux(Index)
        txtTemp = .EditText

        If ReturnCode(IIf(Index = 0, LG_DYE, LG_AUX), , False, txtTemp) Then
            .TextMatrix(Row, 1) = txtTemp
            .TextMatrix(Row, 5) = txtTemp
            .TextMatrix(Row, 4) = txtTemp.Tag
        End If
    End With
End Sub

Private Sub grdDyeAux_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    With grdDyeAux(Index)
        If Col = 1 Then
            .Select Row, 3
'            .EditCell
        ElseIf Col = 3 Then
            .Cell(flexcpText, Row, Col) = SetCurrency(.TextMatrix(Row, Col), 6)

            If Row = .Rows - 1 Then
                If QuestionBox(IIf(Index = 0, "염료", "조제") & "를 계속 추가하시겠습니까 ?") Then
                    Call cmdAddNew_Click(Index)
                Else
                    cmdSave.SetFocus
                End If
            Else
                .Row = .Row + 1
            End If
        End If
    End With
End Sub
Sub FillDyeAuxFind(ByVal pIndex As Integer)
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim rs As Recordset, i%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon

    Set rs = oDyeAux.GetDyeAux(CStr(pIndex))

    txtDyeAuxClss = IIf(pIndex = 0, 1, 0)
    
    Set oDyeAux = Nothing

    faDyeAux.Caption = IIf(pIndex = 1, "염료선택", "조제선택")
    
    With grdDyeAuxFind
        .Redraw = False
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & "" & vbTab & rs!DyeAux & vbTab & rs!DyeAuxID & vbTab & "" & vbTab & "0"
            rs.MoveNext
            DoEvents
        Next i

        rs.Close
        Set rs = Nothing


        .Redraw = True
        .Select .FixedRows, 0, .FixedRows, .Cols - 1
    End With
    DoEvents
    Screen.MousePointer = vbDefault

    m_bloading = False

    Exit Sub

ErrHandler:
    m_bloading = False

    Set rs = Nothing
    Set oDyeAux = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox("frmRecipe.FillDyeAuxFind : " & Err.Number, Err.Source, Err.Description)

End Sub

Private Sub tabform_Click(PreviousTab As Integer)
    'If m_bloading Then Exit Sub
    
    If tabForm.Tab = 0 Then
        Call ClearData

        cmdPrint.Visible = True
        cmdSave.Visible = False
        grdColor.Enabled = True
        pnlMsg.Visible = False
        
        Call cmdSearch_Click(0)
    Else
        pnlMsg.Visible = True
        cmdPrint.Visible = False
        cmdSave.Visible = True
    End If
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkSearch(2).Caption = "Order No"
        chkSearchI(2).Caption = "Order No"
        grdRecipe.ColWidth(2) = 1350
        grdRecipe.ColWidth(3) = 0
        grdOrder.ColWidth(1) = 1290
        grdOrder.ColWidth(2) = 0
    Else
        chkSearch(2).Caption = "관리번호"
        chkSearchI(2).Caption = "관리번호"
        grdRecipe.ColWidth(2) = 0
        grdRecipe.ColWidth(3) = 1350
        grdOrder.ColWidth(1) = 0
        grdOrder.ColWidth(2) = 1290
    End If

End Sub

Private Sub cmdSave_Click()
    If SaveData() Then
        Call MessageBox("저장 되었습니다.")

'        tabForm.Tab = 0
'        Call ClearData
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows
        
        txtQty = ""
        m_sFlag = ID_ADDNEW
    End If
    grdColor.Enabled = True

End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrHandler
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim oRecipe As PlusLib2.CRecipe
    Dim sRecipeNO$, sTitle$
    Dim i%, nCnt%

    If grdRecipe.Rows = grdRecipe.FixedRows Then
        Call MessageBox(LoadResString(203))
        'cmdSearch.SetFocus
        Exit Sub
    End If

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    sRecipeNO = Format(grdRecipe.TextMatrix(grdRecipe.Row, 11), "0000000000")
    If Trim(grdRecipe.TextMatrix(grdRecipe.Row, 18)) = "" Then
        sTitle = "작업 처방전"
    Else
        sTitle = "작업 처방전(수정)"
    End If
    
  
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    Set rs = oRecipe.GetRecipeOne(sRecipeNO)
    
    Set oRecipe = Nothing
    
    nCnt = 0
    
    ReDim Preserve sParam(40)
    
    For i = 0 To 40
        sParam(i) = " "
    Next i
    
    ' 염료 처방내역
    With grdShowDyeAux(0)
    
        For i = 1 To .Rows - 1
            sParam(i - 1) = .TextMatrix(i, 1)
            sParam(i + 9) = .TextMatrix(i, 2)
            nCnt = nCnt + 1
        Next i
    
    End With
    
    
    With grdShowDyeAux(1)
        For i = 1 To .Rows - 1
            sParam(i + 19) = .TextMatrix(i, 1)
            sParam(i + 29) = .TextMatrix(i, 2)
            nCnt = nCnt + 1
            
        Next i
    
    End With
       
    sParam(40) = sTitle
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRecipe = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdRecipe
        .Cols = 20
        Call SetVSFlexGrid(grdRecipe)

        .Redraw = flexRDNone

        .TextArray(1) = "거래처":       .ColWidth(1) = 1350:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "Order No":     .ColWidth(2) = 1350:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호":     .ColWidth(3) = 0:               .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "품명":         .ColWidth(4) = 1300:            .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "색상코드":     .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "색상명":       .ColWidth(6) = 1700:            .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "축율":         .ColWidth(7) = 550:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "원단폭":       .ColWidth(8) = 600:             .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "처방" & vbCrLf & "순위":       .ColWidth(9) = 450:             .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "변경" & vbCrLf & "순위":      .ColWidth(10) = 450:            .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(11) = "처방전번호":  .ColWidth(11) = 990:            .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "처방일자":    .ColWidth(12) = 990:            .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "단위" & vbCrLf & "중량":    .ColWidth(13) = 550:            .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "처방자":      .ColWidth(14) = 1350:           .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "처방자":      .ColWidth(15) = 0
        .TextArray(16) = "비고":        .ColWidth(16) = 0
        .TextArray(17) = "축율":        .ColWidth(17) = 0
        .TextArray(18) = "수정처방구분":   .ColWidth(18) = 0
        .TextArray(19) = "수정처방수량":   .ColWidth(19) = 0
        

        .Redraw = flexRDDirect
    End With

    With grdShowDyeAux(0)
        .Cols = 4
        Call SetVSFlexGrid(grdShowDyeAux(0))

        .Redraw = False

        .TextArray(1) = "염료":         .ColWidth(1) = 2000:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "투입비율":     .ColWidth(2) = 900:             .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "염료":         .ColWidth(3) = 0
        
        .ExtendLastCol = True
        .Redraw = True
    End With

    With grdShowDyeAux(1)
        .Cols = 4
        Call SetVSFlexGrid(grdShowDyeAux(1))

        .Redraw = False

        .TextArray(1) = "조제":         .ColWidth(1) = 2000:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "투입비율":     .ColWidth(2) = 900:             .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "조제":         .ColWidth(3) = 0
        
        .ExtendLastCol = True
        .Redraw = True
    End With

    With grdOrder
        .Cols = 7
        Call SetFlexGrid(grdOrder)

        .Redraw = False

        .TextArray(1) = "Order No":     .ColWidth(1) = 1290:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "관리번호":     .ColWidth(2) = 0:               .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "거래처명":     .ColWidth(3) = LIMIT_WIDTH2:    .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처":       .ColWidth(4) = 0
        .TextArray(5) = "품명":         .ColWidth(5) = 0
        .TextArray(6) = "축율":         .ColWidth(6) = 0

        .Redraw = True
    End With

    With grdColor
        .Cols = 5
        Call SetVSFlexGrid(grdColor)

        .Redraw = False

        .TextArray(1) = "색상번호":     .ColWidth(1) = 0:               .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":       .ColWidth(2) = 3600:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "Design No":    .ColWidth(3) = 2600:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "수주수량":     .ColWidth(4) = LIMIT_WIDTH3:    .ColAlignment(4) = flexAlignRightCenter

        .ExtendLastCol = True
        
        .Redraw = True
    End With

    With grdDyeAux(0)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(0))

        .Redraw = flexRDNone

        .TextArray(1) = "염료":         .ColWidth(1) = LIMIT_WIDTH4:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "염료":         .ColWidth(2) = 300:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "염료투입비율": .ColWidth(3) = 1200:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "염료코드":     .ColWidth(4) = 0
        .TextArray(5) = "염료명":       .ColWidth(5) = 0

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With

    With grdDyeAux(1)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(1))

        .Redraw = flexRDNone

        .TextArray(1) = "조제":         .ColWidth(1) = LIMIT_WIDTH4:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "조제":         .ColWidth(2) = 300:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "조제투입비율": .ColWidth(3) = 1200:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "조제":         .ColWidth(4) = 0
        .TextArray(5) = "조제명":       .ColWidth(5) = 0

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With
    
    
    With grdHistory
        .Cols = 6
        Call SetVSFlexGrid(grdHistory)
        .ScrollBars = flexScrollBarBoth
        
        .Redraw = flexRDNone

        .TextArray(1) = "처방일자":                 .ColWidth(1) = 1000:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "처방" & vbCrLf & "순위":   .ColWidth(2) = 470:     .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "변경" & vbCrLf & "순위":   .ColWidth(3) = 470:     .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "처방자":                   .ColWidth(4) = 700:     .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "비고":                     .ColWidth(5) = 5000:    .ColAlignment(5) = flexAlignLeftCenter
        
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
        
        
    With grdDyeAuxFind
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAuxFind)

        .Redraw = flexRDNone

        .TextArray(0) = "":                 .ColWidth(0) = 0
        .TextArray(1) = "선택":             .ColWidth(1) = 600
        .TextArray(2) = "염료(조제)명":     .ColWidth(2) = 2800:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "코드":             .ColWidth(3) = 400:             .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "구분":             .ColWidth(4) = 0:               .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "Seq":              .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignRightCenter
        
        .ColDataType(1) = flexDTBoolean
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With
        
End Sub



Public Sub FillGridRecipe()
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim nChkDate%, sDate$, eDate$
    Dim nChkCustom%, sCustom$
    Dim nChkOrder%, sOrder$
    Dim nChkArticle%, sArticle$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    nChkDate = IIf(chkSearch(3), 1, 0)
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkCustom = IIf(chkSearch(0), 1, 0)
    sCustom = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(1), 1, 0)
    sArticle = txtSearch(1).Tag
    nChkOrder = IIf(chkSearch(2), IIf(optOrder(0), 2, 1), 0)
    sOrder = IIf(optOrder(0), txtSearch(2), Replace(txtSearch(2), "-", ""))
    
    Set rs = oRecipe.GetRecipe(nChkDate, sDate, eDate, nChkCustom, sCustom, nChkOrder, sOrder, nChkArticle, sArticle)
    
    Set oRecipe = Nothing

    With grdRecipe
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!kCustom & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!Article & vbTab & rs!OrderSeq & vbTab & rs!Color & vbTab & SetCurrency(rs!ChunkRate) & vbTab & _
                rs!StuffWidth & vbTab & CStr(rs!RecipeSeq) & vbTab & CStr(rs!ModifySeq) & vbTab & Format(rs!RecipeNO, "####") & vbTab & _
                MakeDate(DF_LONG, rs!RecipeDate) & vbTab & rs!UnitWght & vbTab & CheckNull(rs!Name) & vbTab & CheckNull(rs!PersonID) & vbTab & _
                CheckNull(rs!Remark) & vbTab & rs!ChunkRate & vbTab & rs!ModiClss & vbTab & rs!Qty

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
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
            cmdOperate(ID_UPDATE).Enabled = True
            cmdOperate(ID_DELETE).Enabled = True
            
            If m_bSaved = True Then
                Call FindNewRow

                m_bSaved = False
            End If
        Else
            cmdPrint.Enabled = False
            .HighLight = flexHighlightNever
            cmdOperate(ID_UPDATE).Enabled = False
            cmdOperate(ID_DELETE).Enabled = False
            MsgBox LoadResString(203), vbInformation
        End If

        .SetFocus
    End With

    Screen.MousePointer = vbDefault

    m_bloading = False
    Call grdRecipe_RowColChange
    
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oRecipe = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Public Sub FillGridOrder()
    Dim oOrder As PlusLib2.COrder
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim nChkDate%, sDate$, eDate$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim nChkOrder%, sOrder$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    nChkDate = IIf(chkSearchI(3), 1, 0)
    sDate = MakeDate(DF_SHORT, dtpDateI(0))
    eDate = MakeDate(DF_SHORT, dtpDateI(1))
    nChkCustom = IIf(chkSearchI(0), 1, 0)
    sCustom = txtSearchI(0).Tag
    nChkArticle = IIf(chkSearchI(1), 1, 0)
    sArticle = txtSearchI(1).Tag
    nChkOrder = IIf(chkSearchI(2), 1, 0)
    sOrder = IIf(optOrder(0), txtSearchI(2), MakeOrderID(txtSearchI(2), OM_REDUCE))
    

    m_bloading = True
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    Set rs = oOrder.GetDraftOrder(nChkDate, sDate, eDate, nChkCustom, sCustom, nChkArticle, sArticle, nChkOrder, sOrder, 0, "", 0, "0")

        
    Set oOrder = Nothing

    With grdOrder
        .Redraw = False

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!kCustom & vbTab & rs!CustomID & vbTab & rs!Article & vbTab & rs!ChunkRate

            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)

            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If

            rs.MoveNext
            DoEvents
        Next i

        rs.Close
        Set rs = Nothing

        .Redraw = True
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        .SetFocus
    End With
    DoEvents
    Screen.MousePointer = vbDefault

    Call FillGridColor

    pnlProgress.Visible = False
    m_bloading = False

    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    m_bloading = False

    Set rs = Nothing
    Set oOrder = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGridColor()
    Dim oOrder As PlusLib2.COrder
    Dim rs As Recordset
    Dim i%

    If grdOrder.Rows = grdOrder.FixedRows Then
        grdColor.Rows = grdColor.FixedRows
        grdColor.HighLight = flexHighlightNever
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon

    Set rs = oOrder.GetOrderSub(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE))
    Set oOrder = Nothing

    With grdColor
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!OrderSeq & vbTab & rs!Color & vbTab & CheckNull(rs!DesignNO) & vbTab & SetCurrency(rs!ColorQty)

            rs.MoveNext
        Next i

        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If

    '    Call ChangeScrollColor
        .Redraw = True
    End With

    Screen.MousePointer = vbDefault

    Call ShowSelOrder

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Function IsGetOrder() As Boolean
    IsGetOrder = False

    With grdOrder
        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Function
    End With
    With grdColor
        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Function
    End With

    IsGetOrder = True
End Function

Private Sub ClearData()
    Dim oRecipe As PlusLib2.CRecipe

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    'Call ShowSelOrder
    
    dtpRecipe = Date
    chkRework = vbUnchecked
    chkRework.Tag = "0"
    txtModify = 1

    txtBox(2) = oRecipe.GetMaxRecipeNo
    
    txtBox(0) = ""
    txtBox(1) = ""
    txtBox(3) = ""
    txtBox(3).Tag = ""
    txtBox(4) = ""
    txtBox(5) = 0
    txtBox(6) = 0
    txtRemark = ""
    txtColorQty = ""
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows
    grdOrder.Rows = grdOrder.FixedRows
    grdOrder.HighLight = flexHighlightNever
    grdColor.Rows = grdColor.FixedRows
    grdColor.HighLight = flexHighlightNever
    

    m_sFlag = ID_ADDNEW
    cmdFind(2).Visible = True
    pnlMsg.Caption = LoadResString(121)

    Set oRecipe = Nothing
End Sub

Private Sub ShowSelOrder()
    If IsGetOrder() Then
        txtBox(0) = grdOrder.TextMatrix(grdOrder.Row, 2)
        txtBox(0).Tag = MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE)
        txtBox(1) = grdColor.TextMatrix(grdColor.Row, 2)
        txtBox(1).Tag = grdColor.TextMatrix(grdColor.Row, 1)
        txtBox(4) = grdOrder.TextMatrix(grdOrder.Row, 5)
        txtBox(6) = grdOrder.TextMatrix(grdOrder.Row, 6)
        txtColorQty = grdColor.TextMatrix(grdColor.Row, 4)
    Else
        txtBox(0) = ""
        txtBox(0).Tag = ""
        txtBox(1) = ""
        txtBox(1).Tag = ""
        txtBox(4) = ""
        txtBox(6) = ""
    End If
End Sub

Private Sub ShowData()

    With grdHistory
        If .Rows > .FixedRows Then .Row = .FixedRows
    End With
    
    With grdRecipe
        txtBox(0) = .TextMatrix(.Row, IIf(optOrder(0), 2, 3))
        txtBox(0).Tag = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
        txtBox(1) = .TextMatrix(.Row, 6)
        txtBox(1).Tag = .TextMatrix(.Row, 5)

        chkRework.Tag = CInt(.TextMatrix(.Row, 9))
        txtModify = .TextMatrix(.Row, 10)
        txtBox(3) = .TextMatrix(.Row, 14)
        txtBox(3).Tag = .TextMatrix(.Row, 15)
        txtBox(4) = grdOrder.TextMatrix(grdOrder.Row, 5)
        txtBox(5) = .TextMatrix(.Row, 13)
        txtBox(6) = .TextMatrix(.Row, 17)
        txtRemark = .TextMatrix(.Row, 16)
    End With

    Dim i%
    With grdShowDyeAux(0)
        For i = 0 To .Rows - .FixedRows - 1
            grdDyeAux(0).AddItem CStr(i + 1) & vbTab & .TextMatrix(.FixedRows + i, 1) & vbTab & vbTab & _
                .TextMatrix(.FixedRows + i, 2) & vbTab & .TextMatrix(.FixedRows + i, 3) & vbTab & .TextMatrix(.FixedRows + i, 1)
        Next i
    End With
    With grdShowDyeAux(1)
        For i = 0 To .Rows - .FixedRows - 1
            grdDyeAux(1).AddItem CStr(i + 1) & vbTab & .TextMatrix(.FixedRows + i, 1) & vbTab & vbTab & _
                .TextMatrix(.FixedRows + i, 2) & vbTab & .TextMatrix(.FixedRows + i, 3) & vbTab & .TextMatrix(.FixedRows + i, 1)
        Next i
    End With

    With grdDyeAux(0)
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpPicture, i, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, i, 2) = flexPicAlignCenterCenter
        Next i
    End With
    With grdDyeAux(1)
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpPicture, i, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, i, 2) = flexPicAlignCenterCenter
        Next i
    End With
End Sub


Private Sub ShowDyeAuxHistory(sOrderID As String, nOrderSeq As Integer, nReworkSeq As Integer)
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As Recordset
    Dim i%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
        
    
    Set rs = oRecipe.GetRecipeHistory(sOrderID, nOrderSeq, nReworkSeq)
    
        
    With grdHistory
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & MakeDate(DF_LONG, rs!SetDate) & vbTab & rs!RecipeSeq & vbTab & rs!ModifySeq & vbTab & _
                    rs!Name & vbTab & CheckNull(rs!Remark)
            .RowHeight(i) = 500
            rs.MoveNext
        Next i

        .Redraw = True
        .SetFocus
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
    End With
    rs.Close

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oRecipe = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub ShowDyeAuxData(sOrderID As String, nOrderSeq As Integer, nRecipeSeq As Integer, nModifySeq As Integer)

    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As Recordset
    Dim i%
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    
    Set rs = oRecipe.GetRecipeSub(sOrderID, nOrderSeq, 1, nRecipeSeq, 1, "1", nModifySeq)
    With grdShowDyeAux(0)
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & SetCurrency(rs!DyeAuxRate, 6) & vbTab & rs!DyeAuxID

            rs.MoveNext
        Next i

    
        .Redraw = True
        .SetFocus
    End With
    rs.Close

    Set rs = oRecipe.GetRecipeSub(sOrderID, nOrderSeq, 1, nRecipeSeq, 1, "0", nModifySeq)
    With grdShowDyeAux(1)
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & SetCurrency(rs!DyeAuxRate, 6) & vbTab & rs!DyeAuxID

            rs.MoveNext
        Next i

        .Redraw = True
        .SetFocus
    End With
    rs.Close

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oRecipe = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Function CheckData() As Boolean
    CheckData = False

    If Len(txtBox(0).Tag) <= 0 Then
        Call MessageBox("'ORDER'를 검색 후 선택하십시오.")
        cmdSearch(1).SetFocus
        Exit Function
    End If
    If Len(txtBox(1).Tag) <= 0 Then
        Call MessageBox("'COLOR'를 검색 후 선택하십시오.")
        cmdSearch(1).SetFocus
        Exit Function
    End If
    If Len(txtBox(2)) <> 10 Then
        Call MessageBox("'처방전번호'를 정확히 입력하십시오.")
        txtBox(2).SetFocus
        Exit Function
    End If
    If Len(txtBox(3).Tag) = 0 Then
        Call MessageBox("'처방자'를 입력하십시오.")
        txtBox(3).SetFocus
        Exit Function
    End If

    Dim i%

    With grdDyeAux(0)
        If .Rows = .FixedRows Then
            Call MessageBox("'염료'를 입력하십시오.")
            cmdAddNew(0).SetFocus
            Exit Function
        End If

        For i = .FixedRows To .Rows - 1
            If Len(.TextMatrix(i, 4)) = 0 Then
                Call MessageBox("'염료'를 선택하십시오.")
                .Select i, 1
                Exit Function
            End If
            If Not IsNumeric(.TextMatrix(i, 3)) Then
                Call MessageBox("'염료투입비율'를 정확히 입력하십시오.")
                .Select i, 3
                Exit Function
            End If
        Next i
    End With

    With grdDyeAux(1)
        For i = .FixedRows To .Rows - 1
            If Len(.TextMatrix(i, 4)) = 0 Then
                Call MessageBox("'조제'를 선택하십시오.")
                .Select i, 1
                Exit Function
            End If
            If Not IsNumeric(.TextMatrix(i, 3)) Then
                Call MessageBox("'조제투입비율'를 정확히 입력하십시오.")
                .Select i, 3
                Exit Function
            End If
        Next i
    End With

    CheckData = True
End Function

Private Function SaveData() As Boolean
    Dim TRec      As PlusLib2.TRecipe
    Dim tRecSub() As PlusLib2.TRecipeSub
    Dim oRecipe   As PlusLib2.CRecipe
    Dim i%, nDyeCnt%, nRecSub%
    Dim sOrder$, nOrderSeq%

    SaveData = False
    If Not CheckData Then Exit Function

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName
          
    sOrder = txtBox(0).Tag
    nOrderSeq = txtBox(1).Tag

    If m_sFlag = ID_ADDNEW Then
        If oRecipe.IsExistRecipe(sOrder, nOrderSeq) Then
            If (chkRework <> vbChecked) Then
              
                If MsgBox("선택하신 수주와 색상은 이미 처방되었습니다." & vbCrLf & vbCrLf & "재처방으로 처리하시겠습니까?", vbYesNo) = vbNo Then
                    Screen.MousePointer = vbDefault
                    Set oRecipe = Nothing
                    
                    MsgBox "저장되지 않았습니다"
                    
                    Exit Function
                End If
            End If
        End If
    End If
    
    With TRec
        .OrderID = txtBox(0).Tag
        .OrderSeq = txtBox(1).Tag
        .RecipeSeq = IIf(m_sFlag = ID_ADDNEW, 1, chkRework.Tag)     ' 재처방
        .ModifySeq = IIf(m_sFlag = ID_ADDNEW, 1, 0)     ' 변경순위
        .RecipeNO = txtBox(2)
        .RecipeDate = MakeDate(DF_SHORT, dtpRecipe)
        .PersonID = txtBox(3).Tag
        .UnitWght = IIf(IsNumeric(txtBox(5)), txtBox(5), 0)
        .ChunkRate = IIf(IsNumeric(txtBox(6)), txtBox(6), 0)
        If m_sFlag = ID_ADDNEW Then
            .ModiClss = ""
            .Qty = CLng(txtColorQty.Text)
        ElseIf m_sFlag = ID_UPDATE Then
            If CLng("0" & txtQty.Text) > 0 Then     ' 수정 처방일 경우
                .ModiClss = "*"
                .Qty = CLng(txtQty.Text)
            Else
                .ModiClss = ""
                .Qty = CLng(txtColorQty.Text)
            End If
        Else
            .ModiClss = ""
            .Qty = CLng(txtColorQty.Text)
        End If
        .Remark = txtRemark
    End With
            
    nRecSub = (grdDyeAux(0).Rows - grdDyeAux(0).FixedRows) + (grdDyeAux(1).Rows - grdDyeAux(1).FixedRows) - 1
    
    ReDim tRecSub(nRecSub)
    With grdDyeAux(0)
        For i = 0 To .Rows - .FixedRows - 1
            If .TextMatrix(.FixedRows + i, 1) <> .TextMatrix(.FixedRows + i, 5) Then
                MsgBox "염료명을 정확히 입력해 주십시오"
                
                Exit Function
            End If
        
            tRecSub(i).OrderID = txtBox(0).Tag
            tRecSub(i).OrderSeq = txtBox(1).Tag
            tRecSub(i).ModifySeq = IIf(m_sFlag = ID_ADDNEW, 1, txtModify)
            tRecSub(i).DyeAuxSeq = i + 1
            tRecSub(i).DyeAuxID = .TextMatrix(.FixedRows + i, 4)
            tRecSub(i).DyeAuxRate = CDbl(.TextMatrix(.FixedRows + i, 3))
        Next i
        nDyeCnt = .Rows - .FixedRows
    End With
    With grdDyeAux(1)
        If .Rows > .FixedRows Then
            For i = 0 To .Rows - .FixedRows - 1
                If .TextMatrix(.FixedRows + i, 1) <> .TextMatrix(.FixedRows + i, 5) Then
                    MsgBox "조제명을 정확히 입력해 주십시오"
                    
                    Exit Function
                End If
                
                tRecSub(i + nDyeCnt).OrderID = txtBox(0).Tag
                tRecSub(i + nDyeCnt).OrderSeq = txtBox(1).Tag
                tRecSub(i + nDyeCnt).ModifySeq = IIf(m_sFlag = ID_ADDNEW, 1, txtModify)
                tRecSub(i + nDyeCnt).DyeAuxSeq = i + nDyeCnt + 1
                tRecSub(i + nDyeCnt).DyeAuxID = .TextMatrix(.FixedRows + i, 4)
                tRecSub(i + nDyeCnt).DyeAuxRate = CDbl(.TextMatrix(.FixedRows + i, 3))
            Next i
        End If
    End With

    
    If m_sFlag = ID_ADDNEW Then
        SaveData = oRecipe.AddNewRecipe(TRec, tRecSub)
    Else
        SaveData = oRecipe.UpdateRecipe(TRec, tRecSub)
    End If

    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    m_sOrderID = TRec.OrderID
    m_sColorID = TRec.OrderSeq
    m_nRecipeSeq = TRec.RecipeSeq
    m_nModifySeq = TRec.ModifySeq
    
    m_bSaved = True

    Exit Function

ErrHandler:
    SaveData = False
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function


'Private Sub FindNewRow(sOrderID As String, sColorID As String, nReworkSeq As Integer, nModifySeq As Integer)
Private Sub FindNewRow()
    Dim i%
    
    With grdRecipe
        
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 3) = MakeOrderID(m_sOrderID, OM_EXPAND) Then    ' 관리번호 비교
                If .TextMatrix(i, 5) = m_sColorID Then  ' 색상번호 비교
                    If .TextMatrix(i, 9) = m_nRecipeSeq Then    ' 처방순위 비교
                        If .TextMatrix(i, 10) = m_nModifySeq Then   ' 변경순위 비교
                            .Row = i
                            .TopRow = i
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next i
    
        .Row = .FixedRows
    End With
End Sub


Private Function DeleteData() As Boolean
    Dim oRecipe As PlusLib2.CRecipe
    Dim nUseCount%, sMessage$
    Dim sOrderID$, nOrderSeq%, nRecipeSeq%

    On Error GoTo ErrHandler

    DeleteData = False
    With grdRecipe
        sOrderID = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
        nOrderSeq = .TextMatrix(.Row, 5)
        nRecipeSeq = CInt(.TextMatrix(.Row, 9))
    End With
        

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName
    
    ' 처방전이 본작업 처방전에 사용된경우 삭제안됨
    nUseCount = oRecipe.GetRecipeUseCount(sOrderID, nOrderSeq, nRecipeSeq, nUseCount)

    
    If nUseCount > 0 Then
        sMessage = "처방전 관리번호 : " & MakeOrderID(sOrderID, OM_EXPAND) & vbCrLf & _
                    "처방전 색상번호 : " & nOrderSeq & vbCrLf & _
                    "처방순위 : " & nRecipeSeq & vbCrLf & vbCrLf & _
                    "이 처방전은 본처방 작업에 사용된 처방전입니다." & vbCrLf & "삭제할 수 없습니다"
                    
        Set oRecipe = Nothing
        
        MessageBox sMessage
        DeleteData = False
        Exit Function
    End If


    With grdRecipe
        DeleteData = oRecipe.DeleteRecipe(sOrderID, nOrderSeq, nRecipeSeq)
        
        sMessage = "처방전 관리번호 : " & MakeOrderID(sOrderID, OM_EXPAND) & vbCrLf & _
                    "처방전 색상번호 : " & nOrderSeq & vbCrLf & _
                    "처방순위 : " & nRecipeSeq & vbCrLf & vbCrLf & _
                    "처방전이 삭제되었습니다."
                    
        MessageBox sMessage
    End With

    Set oRecipe = Nothing

    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Set oRecipe = Nothing
End Function

Private Sub ChangeScrollRecipe()
    With grdRecipe
        .ColWidth(6) = IIf(.Rows > LIMIT_ROW1 + .FixedRows, LIMIT_WIDTH1 - 240, LIMIT_WIDTH1)
    End With
End Sub

Private Sub ChangeScrollOrder()
    With grdOrder
        .ColWidth(3) = IIf(.Rows > LIMIT_ROW2 + .FixedRows, LIMIT_WIDTH2 - 240, LIMIT_WIDTH2)
    End With
End Sub

Private Sub ChangeScrollColor()
    With grdColor
        .ColWidth(4) = IIf(.Rows > LIMIT_ROW3 + .FixedRows, LIMIT_WIDTH3 - 240, LIMIT_WIDTH3)
    End With
End Sub

Private Sub ChangeScrollDyeAux(Index As Integer)
    With grdDyeAux(Index)
        .ColWidth(1) = IIf(.Rows > LIMIT_ROW4 + .FixedRows, LIMIT_WIDTH4 - 240, LIMIT_WIDTH4)
    End With
End Sub



Private Sub GetRecipeOne(sRecipeNO As String)
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim sRemark$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    Set rs = oRecipe.GetRecipeOne(Format(sRecipeNO, "0000000000"))
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        Call MessageBox("'" & sRecipeNO & "' 번호의 처방전이 존재하지 않습니다.")

        rs.Close
        Set rs = Nothing
        Set oRecipe = Nothing

        Exit Sub
    End If

    Dim sOrderID$, sColorID%, nReworkSeq%, nModifySeq%
    Dim i%
    
    sRemark = txtRemark
'    Call ClearData
    txtRemark = sRemark

    txtBox(3) = CheckNull(rs!Name)
    txtBox(3).Tag = rs!PersonID
    sOrderID = rs!OrderID
    sColorID = rs!OrderSeq
    nReworkSeq = rs!RecipeSeq
    nModifySeq = CheckNull(rs!ModifySeq)
    rs.Close

    oRecipe.Connection = g_adoCon
    Set rs = oRecipe.GetRecipeSub(sOrderID, sColorID, 1, nReworkSeq, 1, "1", nModifySeq)
    With grdDyeAux(0)
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & vbTab & SetCurrency(rs!DyeAuxRate, 6) & vbTab & rs!DyeAuxID & vbTab & rs!DyeAux

            .Cell(flexcpPicture, i, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, i, 2) = flexPicAlignCenterCenter

            rs.MoveNext
        Next i

        Call ChangeScrollDyeAux(0)

        .Redraw = flexRDDirect
    End With
    rs.Close

    oRecipe.Connection = g_adoCon
    Set rs = oRecipe.GetRecipeSub(sOrderID, sColorID, 1, nReworkSeq, 1, "0", nModifySeq)
    With grdDyeAux(1)
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & vbTab & SetCurrency(rs!DyeAuxRate, 6) & vbTab & rs!DyeAuxID & vbTab & rs!DyeAux

            .Cell(flexcpPicture, i, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, i, 2) = flexPicAlignCenterCenter

            rs.MoveNext
        Next i

        Call ChangeScrollDyeAux(1)

        .Redraw = flexRDDirect
    End With
    rs.Close

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    dtpRecipe.SetFocus

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault
End Sub



