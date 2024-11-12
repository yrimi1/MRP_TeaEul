VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRecipeCalc 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9390
   ClientLeft      =   1635
   ClientTop       =   1545
   ClientWidth     =   15405
   Icon            =   "frmRecipeCalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   15405
   Begin TabDlg.SSTab sbTab 
      Height          =   9375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   16536
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmRecipeCalc.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdDelete"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAddRecipe"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClose"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraSearch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdExit"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdRecipeCal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "pnlRapidOrder"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "stTab"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "pnlProgress"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmRecipeCalc.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPrint"
      Tab(1).Control(1)=   "cmdCancel"
      Tab(1).Control(2)=   "cmdSave"
      Tab(1).Control(3)=   "Text2"
      Tab(1).Control(4)=   "SSPanel1"
      Tab(1).Control(5)=   "fraRecipe"
      Tab(1).Control(6)=   "fraMatch"
      Tab(1).Control(7)=   "SSPanel4"
      Tab(1).ControlCount=   8
      Begin Threed.SSPanel pnlProgress 
         Height          =   870
         Left            =   1740
         TabIndex        =   110
         Top             =   2550
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
            TabIndex        =   111
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
            TabIndex        =   112
            Top             =   120
            Width           =   270
         End
      End
      Begin TabDlg.SSTab stTab 
         Height          =   4095
         Left            =   30
         TabIndex        =   129
         Top             =   990
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   7223
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "공정별 카드내역"
         TabPicture(0)   =   "frmRecipeCalc.frx":0044
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdData"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "평량지시 내역"
         TabPicture(1)   =   "frmRecipeCalc.frx":0060
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdRecipeCalc"
         Tab(1).ControlCount=   1
         Begin VSFlex7LCtl.VSFlexGrid grdRecipeCalc 
            Height          =   3675
            Left            =   -74940
            TabIndex        =   132
            Top             =   360
            Width           =   15135
            _cx             =   26696
            _cy             =   6482
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
            Height          =   3675
            Left            =   30
            TabIndex        =   130
            Top             =   360
            Width           =   15195
            _cx             =   26802
            _cy             =   6482
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
      Begin Threed.SSPanel pnlRapidOrder 
         Height          =   2835
         Left            =   60
         TabIndex        =   113
         Top             =   5670
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   5001
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
            Index           =   0
            Left            =   9300
            TabIndex        =   126
            Tag             =   "염색호기"
            Top             =   390
            Width           =   825
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
            Left            =   10140
            TabIndex        =   125
            Tag             =   "염색패턴"
            Top             =   390
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
            Index           =   2
            Left            =   13470
            TabIndex        =   124
            Tag             =   "염색구분"
            Top             =   390
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
            Index           =   4
            Left            =   12120
            TabIndex        =   123
            Tag             =   "염색구분"
            Top             =   390
            Width           =   1335
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
            Index           =   3
            Left            =   14310
            TabIndex        =   122
            Tag             =   "작업자"
            Top             =   2220
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtRemark1 
            Height          =   315
            Left            =   1380
            TabIndex        =   114
            Top             =   2460
            Width           =   7845
         End
         Begin VSFlex7LCtl.VSFlexGrid grdCardList 
            Height          =   2400
            Left            =   30
            TabIndex        =   115
            Top             =   15
            Width           =   9240
            _cx             =   16298
            _cy             =   4233
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
            Index           =   5
            Left            =   9300
            TabIndex        =   116
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   635
            _Version        =   196609
            Caption         =   "염색호기"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   360
            Index           =   6
            Left            =   10140
            TabIndex        =   117
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
            Index           =   7
            Left            =   13470
            TabIndex        =   118
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
            Index           =   8
            Left            =   14340
            TabIndex        =   119
            Top             =   1860
            Visible         =   0   'False
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
            Index           =   9
            Left            =   30
            TabIndex        =   120
            Top             =   2460
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
            Index           =   10
            Left            =   12150
            TabIndex        =   121
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
      Begin Threed.SSCommand cmdRecipeCal 
         Height          =   510
         Left            =   90
         TabIndex        =   108
         Top             =   5130
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "카드선택"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   690
         Left            =   13605
         TabIndex        =   109
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "      닫기(&X)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   690
         Left            =   -63060
         TabIndex        =   82
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "      인쇄(&P)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   690
         Left            =   -61365
         TabIndex        =   83
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "      취소(&C)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   690
         Left            =   -65130
         TabIndex        =   84
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "      저장(&M)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '없음
         Height          =   630
         Left            =   -74910
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   81
         Text            =   "frmRecipeCalc.frx":007C
         Top             =   8610
         Width           =   4785
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2745
         Left            =   -74940
         TabIndex        =   4
         Top             =   60
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   4842
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtINQty 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   2850
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1275
            Width           =   930
         End
         Begin VB.TextBox txtRoll 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   975
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1275
            Width           =   930
         End
         Begin VB.TextBox txtRemark 
            BackColor       =   &H00FFFFFF&
            Height          =   630
            Left            =   4770
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   945
            Width           =   3045
         End
         Begin VB.TextBox txtColor 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4770
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   270
            Width           =   3060
         End
         Begin VB.TextBox txtWorkClss 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4770
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   615
            Width           =   3060
         End
         Begin VB.TextBox txtArticle 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   975
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   615
            Width           =   2805
         End
         Begin VB.TextBox txtCustom 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   975
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   285
            Width           =   2805
         End
         Begin VB.TextBox txtDyeID 
            Height          =   285
            Left            =   225
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1755
            Visible         =   0   'False
            Width           =   660
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   11
            Left            =   15
            TabIndex        =   13
            Top             =   945
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "염색 호기"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   2
            Left            =   15
            TabIndex        =   14
            Top             =   1275
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "투입  절수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   0
            Left            =   1890
            TabIndex        =   15
            Top             =   1275
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "투입   수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtPattern 
            Height          =   300
            Left            =   2850
            TabIndex        =   16
            Top             =   945
            Width           =   930
            _ExtentX        =   1640
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
            BackColor       =   16777215
            IMEMode         =   10
         End
         Begin MRPPlus2.WizText txtMachine 
            Height          =   300
            Left            =   975
            TabIndex        =   17
            Top             =   945
            Width           =   930
            _ExtentX        =   1640
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
            BackColor       =   16777215
            IMEMode         =   10
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   12
            Left            =   1890
            TabIndex        =   18
            Top             =   945
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "Pattern No"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   630
            Index           =   16
            Left            =   3795
            TabIndex        =   19
            Top             =   945
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   1111
            _Version        =   196609
            Caption         =   "비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   4
            Left            =   3795
            TabIndex        =   20
            Top             =   270
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "색상"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   300
            Left            =   3795
            TabIndex        =   21
            Top             =   615
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "작업구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   15
            TabIndex        =   22
            Top             =   615
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "품       명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Left            =   15
            TabIndex        =   23
            Top             =   285
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "거  래  처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid grdCard 
            Height          =   1140
            Left            =   0
            TabIndex        =   24
            Top             =   1590
            Width           =   7815
            _cx             =   13785
            _cy             =   2011
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
            AutoSearch      =   1
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   300
            Left            =   0
            TabIndex        =   25
            Top             =   -15
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "   염색  스케쥴"
            Alignment       =   1
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSFrame fraRecipe 
         Height          =   1455
         Left            =   -67080
         TabIndex        =   26
         Top             =   30
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   2566
         _Version        =   196609
         Begin VB.TextBox txtRecipeSeq 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3165
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   615
            Width           =   645
         End
         Begin VB.TextBox txtRecipeNO 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   300
            Width           =   1155
         End
         Begin VB.TextBox txtModifySeq 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4695
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   615
            Width           =   645
         End
         Begin VB.TextBox txtOrderID 
            Height          =   270
            Left            =   -180
            TabIndex        =   31
            Top             =   -60
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtRecipePerson 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3165
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   300
            Width           =   1140
         End
         Begin VB.TextBox txtWght 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "12"
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtRecipeRemark 
            BackColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   1035
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   900
            Width           =   6315
         End
         Begin Threed.SSCommand cmdRecipe 
            Height          =   375
            Left            =   5580
            TabIndex        =   28
            Top             =   315
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   661
            _Version        =   196609
            Caption         =   "처방전 변경"
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   3
            Left            =   2205
            TabIndex        =   35
            Top             =   300
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "처방자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   6
            Left            =   60
            TabIndex        =   36
            Top             =   300
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "처방번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   7
            Left            =   2205
            TabIndex        =   37
            Top             =   615
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "처방순위"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   8
            Left            =   3825
            TabIndex        =   38
            Top             =   615
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "변경순위"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   9
            Left            =   45
            TabIndex        =   39
            Top             =   615
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "단위 중량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   540
            Index           =   1
            Left            =   45
            TabIndex        =   40
            Top             =   915
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   953
            _Version        =   196609
            Caption         =   "특기사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   300
            Left            =   30
            TabIndex        =   41
            Top             =   0
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "   처   방   전"
            Alignment       =   1
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSFrame fraMatch 
         Height          =   1290
         Left            =   -67050
         TabIndex        =   42
         Top             =   1500
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   2275
         _Version        =   196609
         Begin VB.TextBox txtRPCalcRemark 
            BackColor       =   &H00FFFFFF&
            Height          =   570
            Left            =   1050
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   675
            Width           =   6285
         End
         Begin MRPPlus2.WizText txtPerson 
            Height          =   300
            Left            =   1065
            TabIndex        =   44
            Top             =   345
            Width           =   975
            _ExtentX        =   1720
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
            BackColor       =   16777152
            IMEMode         =   10
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   10
            Left            =   45
            TabIndex        =   45
            Top             =   345
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "평량작성자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   0
            Left            =   2070
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   345
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
         Begin Threed.SSPanel pnlName 
            Height          =   555
            Index           =   5
            Left            =   30
            TabIndex        =   47
            Top             =   690
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   979
            _Version        =   196609
            Caption         =   "비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   300
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "   평 량 지 시"
            Alignment       =   1
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdRemarkCopy 
            Height          =   375
            Left            =   5565
            TabIndex        =   49
            Top             =   300
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   661
            _Version        =   196609
            Caption         =   "특기사항 복사"
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   5715
         Left            =   -74940
         TabIndex        =   50
         Top             =   2820
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   10081
         _Version        =   196609
         BackColor       =   12632256
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel pnlCalc 
            Height          =   5700
            Index           =   1
            Left            =   7755
            TabIndex        =   51
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   10054
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin TabDlg.SSTab tabDye 
               Height          =   3015
               Left            =   15
               TabIndex        =   52
               Top             =   30
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   5318
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "추가 1"
               TabPicture(0)   =   "frmRecipeCalc.frx":00E3
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "cmdDyeDel(0)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "grdDye(0)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "cmdDyeAdd(0)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "txtDyeTemp"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).ControlCount=   4
               TabCaption(1)   =   "추가 2"
               TabPicture(1)   =   "frmRecipeCalc.frx":00FF
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "cmdDyeDel(1)"
               Tab(1).Control(1)=   "cmdDyeAdd(1)"
               Tab(1).Control(2)=   "grdDye(1)"
               Tab(1).ControlCount=   3
               TabCaption(2)   =   "추가 3"
               TabPicture(2)   =   "frmRecipeCalc.frx":011B
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "cmdDyeDel(2)"
               Tab(2).Control(1)=   "cmdDyeAdd(2)"
               Tab(2).Control(2)=   "grdDye(2)"
               Tab(2).ControlCount=   3
               Begin VB.TextBox txtDyeTemp 
                  Height          =   270
                  Left            =   4710
                  TabIndex        =   53
                  Top             =   1605
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin Threed.SSCommand cmdDyeAdd 
                  Height          =   705
                  Index           =   0
                  Left            =   6660
                  TabIndex        =   54
                  Top             =   360
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "추가"
               End
               Begin VSFlex7LCtl.VSFlexGrid grdDye 
                  Height          =   2625
                  Index           =   0
                  Left            =   45
                  TabIndex        =   55
                  Top             =   360
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   4630
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
               Begin VSFlex7LCtl.VSFlexGrid grdDye 
                  Height          =   2625
                  Index           =   1
                  Left            =   -74955
                  TabIndex        =   56
                  Top             =   360
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   4630
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
               Begin VSFlex7LCtl.VSFlexGrid grdDye 
                  Height          =   2625
                  Index           =   2
                  Left            =   -74955
                  TabIndex        =   57
                  Top             =   360
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   4630
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
               Begin Threed.SSCommand cmdDyeDel 
                  Height          =   705
                  Index           =   0
                  Left            =   6660
                  TabIndex        =   58
                  Top             =   1095
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "제거"
               End
               Begin Threed.SSCommand cmdDyeAdd 
                  Height          =   705
                  Index           =   1
                  Left            =   -68340
                  TabIndex        =   59
                  Top             =   360
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "추가"
               End
               Begin Threed.SSCommand cmdDyeDel 
                  Height          =   705
                  Index           =   1
                  Left            =   -68340
                  TabIndex        =   60
                  Top             =   1095
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "제거"
               End
               Begin Threed.SSCommand cmdDyeAdd 
                  Height          =   705
                  Index           =   2
                  Left            =   -68340
                  TabIndex        =   61
                  Top             =   360
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "추가"
               End
               Begin Threed.SSCommand cmdDyeDel 
                  Height          =   705
                  Index           =   2
                  Left            =   -68340
                  TabIndex        =   62
                  Top             =   1095
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "제거"
               End
            End
            Begin TabDlg.SSTab tabAux 
               Height          =   2310
               Left            =   30
               TabIndex        =   63
               Top             =   3045
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   4075
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "추가 1"
               TabPicture(0)   =   "frmRecipeCalc.frx":0137
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "cmdAuxDel(0)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "cmdAuxAdd(0)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "grdAux(0)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "txtAuxTemp"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).ControlCount=   4
               TabCaption(1)   =   "추가 2"
               TabPicture(1)   =   "frmRecipeCalc.frx":0153
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "cmdAuxDel(1)"
               Tab(1).Control(1)=   "cmdAuxAdd(1)"
               Tab(1).Control(2)=   "grdAux(1)"
               Tab(1).ControlCount=   3
               TabCaption(2)   =   "추가 3"
               TabPicture(2)   =   "frmRecipeCalc.frx":016F
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "cmdAuxDel(2)"
               Tab(2).Control(1)=   "cmdAuxAdd(2)"
               Tab(2).Control(2)=   "grdAux(2)"
               Tab(2).ControlCount=   3
               Begin VB.TextBox txtAuxTemp 
                  Height          =   270
                  Left            =   4770
                  TabIndex        =   64
                  Top             =   1650
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VSFlex7LCtl.VSFlexGrid grdAux 
                  Height          =   1905
                  Index           =   0
                  Left            =   45
                  TabIndex        =   65
                  Top             =   360
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   3360
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
               Begin VSFlex7LCtl.VSFlexGrid grdAux 
                  Height          =   1890
                  Index           =   1
                  Left            =   -74955
                  TabIndex        =   66
                  Top             =   360
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   3334
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
               Begin VSFlex7LCtl.VSFlexGrid grdAux 
                  Height          =   1890
                  Index           =   2
                  Left            =   -74955
                  TabIndex        =   67
                  Top             =   360
                  Width           =   6600
                  _cx             =   11642
                  _cy             =   3334
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
               Begin Threed.SSCommand cmdAuxAdd 
                  Height          =   705
                  Index           =   0
                  Left            =   6660
                  TabIndex        =   68
                  Top             =   360
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "추가"
               End
               Begin Threed.SSCommand cmdAuxDel 
                  Height          =   705
                  Index           =   0
                  Left            =   6660
                  TabIndex        =   69
                  Top             =   1095
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "제거"
               End
               Begin Threed.SSCommand cmdAuxAdd 
                  Height          =   705
                  Index           =   1
                  Left            =   -68340
                  TabIndex        =   70
                  Top             =   360
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "추가"
               End
               Begin Threed.SSCommand cmdAuxDel 
                  Height          =   705
                  Index           =   1
                  Left            =   -68340
                  TabIndex        =   71
                  Top             =   1095
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "제거"
               End
               Begin Threed.SSCommand cmdAuxAdd 
                  Height          =   705
                  Index           =   2
                  Left            =   -68340
                  TabIndex        =   72
                  Top             =   360
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "추가"
                  PictureAlignment=   6
               End
               Begin Threed.SSCommand cmdAuxDel 
                  Height          =   705
                  Index           =   2
                  Left            =   -68340
                  TabIndex        =   73
                  Top             =   1095
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   1244
                  _Version        =   196609
                  Caption         =   "제거"
               End
            End
         End
         Begin Threed.SSPanel pnlTitle 
            Height          =   360
            Left            =   -15
            TabIndex        =   74
            Top             =   0
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   635
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "   염조제 투입량"
            Alignment       =   1
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCalc 
            Height          =   5355
            Index           =   0
            Left            =   30
            TabIndex        =   75
            Top             =   360
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   9446
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSPanel SSPanel6 
               Height          =   360
               Left            =   45
               TabIndex        =   76
               Top             =   30
               Width           =   7680
               _ExtentX        =   13547
               _ExtentY        =   635
               _Version        =   196609
               BackColor       =   -2147483638
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "염료 투입량"
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
            End
            Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
               Height          =   2595
               Index           =   0
               Left            =   45
               TabIndex        =   77
               Top             =   390
               Width           =   7680
               _cx             =   13547
               _cy             =   4577
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
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
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmRecipeCalc.frx":018B
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
            Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
               Height          =   1905
               Index           =   1
               Left            =   45
               TabIndex        =   78
               Top             =   3405
               Width           =   7680
               _cx             =   13547
               _cy             =   3360
               _ConvInfo       =   1
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   360
               Left            =   45
               TabIndex        =   79
               Top             =   3060
               Width           =   7680
               _ExtentX        =   13547
               _ExtentY        =   635
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
               Caption         =   "조제 투입량"
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               FloodShowPct    =   -1  'True
               Begin VB.TextBox txtWaterRate 
                  Height          =   270
                  Left            =   5055
                  TabIndex        =   80
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   1065
               End
            End
         End
      End
      Begin Threed.SSFrame fraSearch 
         Height          =   915
         Left            =   60
         TabIndex        =   85
         Top             =   60
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   1614
         _Version        =   196609
         Begin VB.ComboBox cboProcess 
            Height          =   300
            Left            =   9990
            Style           =   2  '드롭다운 목록
            TabIndex        =   92
            Top             =   75
            Width           =   1335
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "검색(&F)"
            Height          =   780
            Left            =   14370
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   91
            ToolTipText     =   "자료 저장"
            Top             =   60
            Width           =   780
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   3
            Left            =   6600
            TabIndex        =   90
            Top             =   75
            Width           =   1905
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   2
            Left            =   2820
            TabIndex        =   89
            Top             =   495
            Width           =   1905
         End
         Begin VB.TextBox txtSearch 
            Height          =   300
            Index           =   1
            Left            =   2850
            TabIndex        =   88
            Top             =   75
            Width           =   1905
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   4
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   87
            Top             =   495
            Width           =   1185
         End
         Begin VB.TextBox txtSearch 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   5
            Left            =   7830
            MaxLength       =   4
            TabIndex        =   86
            Top             =   495
            Width           =   675
         End
         Begin Threed.SSPanel pnlOrder 
            Height          =   795
            Left            =   60
            TabIndex        =   93
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   1402
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.OptionButton optOrder 
               Caption         =   "Order No."
               Height          =   180
               Index           =   0
               Left            =   60
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   120
               Width           =   1200
            End
            Begin VB.OptionButton optOrder 
               Caption         =   "관리 번호"
               Height          =   180
               Index           =   1
               Left            =   60
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   480
               Value           =   -1  'True
               Width           =   1200
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   96
            Top             =   75
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
               TabIndex        =   97
               Top             =   45
               Width           =   975
            End
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   1
            Left            =   4785
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   75
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
            Left            =   1440
            TabIndex        =   99
            Top             =   495
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
               TabIndex        =   100
               Top             =   60
               Width           =   975
            End
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   2
            Left            =   4770
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   495
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
            Index           =   0
            Left            =   5220
            TabIndex        =   102
            Top             =   75
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
               Caption         =   "관리번호"
               Height          =   180
               Index           =   3
               Left            =   60
               TabIndex        =   103
               Top             =   60
               Width           =   1185
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   5220
            TabIndex        =   104
            Top             =   495
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
               Caption         =   "카드번호"
               Height          =   180
               Index           =   4
               Left            =   60
               TabIndex        =   105
               Top             =   60
               Width           =   1185
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   8610
            TabIndex        =   106
            Top             =   60
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
               Caption         =   "대기공정"
               Height          =   180
               Index           =   5
               Left            =   60
               TabIndex        =   107
               Top             =   60
               Width           =   1185
            End
         End
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   510
         Left            =   12600
         TabIndex        =   127
         Top             =   5130
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "평량지시작성"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdClose 
         Height          =   510
         Left            =   13980
         TabIndex        =   128
         Top             =   5130
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "취소"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdAddRecipe 
         Height          =   510
         Left            =   11220
         TabIndex        =   131
         Top             =   5130
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "추가작업"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   510
         Left            =   9840
         TabIndex        =   133
         Top             =   5130
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "평량지시삭제"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel pnlRecipe 
      Height          =   5820
      Left            =   10230
      TabIndex        =   0
      Top             =   10305
      Visible         =   0   'False
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   10266
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdSelect 
         Height          =   690
         Left            =   5040
         TabIndex        =   2
         Top             =   5040
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "      선택"
         PictureAlignment=   1
      End
      Begin VSFlex7LCtl.VSFlexGrid grdRecipe 
         Height          =   4830
         Left            =   150
         TabIndex        =   1
         Top             =   135
         Width           =   6510
         _cx             =   11483
         _cy             =   8520
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
   Begin Crystal.CrystalReport cryReport 
      Left            =   3660
      Top             =   8865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRecipeCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\RecipeCalc.rpt"

Private Const LIMIT_ROW1 = 6
Private Const LIMIT_ROW2 = 6
Private Const LIMIT_ROW3 = 23
Private Const LIMIT_WIDTH = 2710
Private Const LIMIT_WIDTH3 = 1250

Private m_iFlag     As Integer   ' 현재 상태 (추가/수정/삭제/검색)
Private m_bLoading  As Boolean
Private m_bLoading1 As Boolean
Private m_bLoading2 As Boolean
Private m_bSaved   As Boolean '저장후 처방전 출력을 하기 위한 조회 Flag
Private m_bModify   As Boolean  ' 수정작업 중인지 여부

Private m_nDyeID    As Long   ' 스케쥴 번호
Private m_nDyeSeq   As Integer  ' 염색 순위
Private m_nWorkClss As Integer  ' 작업 구분 : 얼룩수정, 주름수정, 오염수정, 색수정....

' 염색 스케쥴 번호
Public Property Let DyeID(nID As Long)
    m_nDyeID = nID
End Property

' 염색 순번
Public Property Let DyeSeq(nSeq As Integer)
    m_nDyeSeq = nSeq
    
End Property



Private Sub chkSearch_Click(Index As Integer)
    If Index >= 1 And Index <= 3 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    ElseIf Index = 4 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(4).Enabled = True
            txtSearch(5).Enabled = True
            txtSearch(4).SetFocus
        Else
            txtSearch(4).Enabled = False
            txtSearch(5).Enabled = False
        End If
    Else
        If chkSearch(Index).Value = vbChecked Then
            cboProcess.Enabled = True
            cboProcess.SetFocus
        Else
            cboProcess.Enabled = False
        End If
    End If
End Sub

Private Sub cmdAddRecipe_Click()
Dim oRapid As PlusLib2.CRapid
    Dim nSchID As Long
    Dim nSeq As Integer
    Dim nNewDyeSeq As Integer

    If grdCardList.Rows = grdCardList.FixedRows Then Exit Sub
    
    If grdCardList.TextMatrix(1, 11) <> "작업" Then
        MsgBox "추가작업은 염색작업중인 카드만가능합니다." & vbCrLf & "본작업에 대한 평량지시를 먼저 내려주십시오.", vbCritical, "작성 오류"
        Exit Sub
    End If
    
    If MsgBox("정말로 추가작업을 하시겠습니까?", vbYesNo + vbQuestion, "추가작업") = vbYes Then
        Set oRapid = New PlusLib2.CRapid
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
        
        With grdCardList
            nSchID = .TextMatrix(1, 23)
            nSeq = .TextMatrix(1, 24)
        End With
                
        If oRapid.AddDyeWorkRapid(nSchID, nSeq, Format(Now, "YYYYMMDD"), Format(time, "HHMM"), nNewDyeSeq) Then
            Set oRapid = Nothing
            MsgBox "염색 추가작업이 정상적으로 처리되었습니다" & vbCrLf & _
                   "추가작업에 대한 평량지시를 내린후 작업을 진행시켜야 합니다", vbInformation, "추가작업"
                   
            Call FillGridRecipeCalc
            sbTab.Tab = 1
            
            Call SetInstruction(nSchID, nNewDyeSeq)
        Else
            Set oRapid = Nothing
        End If
    End If

End Sub

' 평량 추가작업 지시. 조제 추가
Private Sub cmdAuxAdd_Click(Index As Integer)

    If Index <> m_nDyeSeq - 2 Then Exit Sub
    
    With grdAux(Index)
    
        .AddItem CStr(.Rows) & vbTab & " " & vbTab & 0 & vbTab & 0
    
        If ReturnCode(LG_AUX, , False, txtAuxTemp) = True Then
    
            .TextMatrix(.Rows - 1, 1) = txtAuxTemp
            .TextMatrix(.Rows - 1, 4) = txtAuxTemp.Tag
            
            txtAuxTemp = ""
            txtAuxTemp.Tag = 0
            
            .Row = .Row - 1
        Else
            .RemoveItem .Rows - 1
            
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "cmdDyeAdd_Click", Err.Description)

End Sub


' 평량 추가작업 지시. 조제 삭제
Private Sub cmdAuxDel_Click(Index As Integer)
    If Index <> m_nDyeSeq - 2 Then Exit Sub
    
    With grdAux(Index)
    
        If .Rows = .FixedRows Then Exit Sub
               
        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
            .RemoveItem .Row
            
        End If

    End With
    
End Sub


Private Sub cmdCancel_Click()
    sbTab.Tab = 0
    grdCardList.Rows = grdCardList.FixedRows
    
    If stTab.Tab = 0 Then
        Call FillGridData
    Else
        Call FillGridRecipeCalc
    End If
End Sub

Private Sub cmdClose_Click()
    grdCardList.Rows = grdCardList.FixedRows
End Sub

Private Sub cmdDelete_Click()

    If grdRecipeCalc.TextMatrix(grdRecipeCalc.Row, 13) = "작업" Then
        MsgBox "염색작업중인 카드입니다." & vbCrLf & "작업진행중인 카드는 평량지시내역을 삭제할 수 없습니다.", vbCritical
        Exit Sub
    End If
    
    If MsgBox("평량지시내역을 삭제하시겠습니까?", vbYesNo) = vbNo Then Exit Sub
    
    If DeleteData Then
        Call FillGridRecipeCalc
    End If

End Sub

' 평량 추가작업 지시. 염료 추가
Private Sub cmdDyeAdd_Click(Index As Integer)
    
    If Index <> m_nDyeSeq - 2 Then Exit Sub
    
    With grdDye(Index)
    
        .AddItem CStr(.Rows) & vbTab & " " & vbTab & 0 & vbTab & 0
    
        If ReturnCode(LG_DYE, , False, txtDyeTemp) = True Then
    
            .TextMatrix(.Rows - 1, 1) = txtDyeTemp
            .TextMatrix(.Rows - 1, 4) = txtDyeTemp.Tag
            
            txtDyeTemp = ""
            txtDyeTemp.Tag = 0
            
            .Row = .Rows - 1
        Else
            .RemoveItem .Rows - 1
            
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "cmdDyeAdd_Click", Err.Description)
    
End Sub


' 평량 추가작업 지시. 염료삭제
Private Sub cmdDyeDel_Click(Index As Integer)
    
    If Index <> m_nDyeSeq - 2 Then Exit Sub
    
    With grdDye(Index)
    
        If .Rows = .FixedRows Then Exit Sub
               
        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
            .RemoveItem .Row
            
            .Row = .Rows - 1
        End If

    End With
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdFind_Click(Index As Integer)
    Dim sOrderID$
    
    If Index = 0 Then
        Call ReturnCode(LG_PERSON, , False, txtPerson)
    ElseIf Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If

End Sub


Private Sub cmdOK_Click()
    Dim nDyeID&, nDyeSeq%
    
    With grdCardList
        If .Rows = .FixedRows Then Exit Sub
        
        If Not CheckRapidData Then Exit Sub
        
        nDyeID = CheckNum(.TextMatrix(.FixedRows, 23))
        nDyeSeq = CheckNum(.TextMatrix(.FixedRows, 24))
    End With
    
    sbTab.Tab = 1
    Call SetInstruction(nDyeID, nDyeSeq)

End Sub

Private Sub cmdPrint_Click()
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim i%, nPos%, j%, k%
    Dim sDye() As String
    Dim sAux() As String
    Dim nDyeCnt%, nAuxCnt%
    Dim bFind As Boolean
    Dim sCard$
    
    On Error GoTo ErrHandler
    
    ' 평량 지시 내역이 저장되지 않았다면 함수 종료
'    If m_bSaved = False Then
'        MsgBox "먼저 저장한 후 출력하십시오"
'        Exit Sub
'    End If
    
    If grdCard.Rows = grdCard.FixedRows Then
        MessageBox "공정카드가 설정되지 않았습니다"
        Exit Sub
    End If
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    ' Printing
    Screen.MousePointer = vbHourglass
    
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    Set rs = oRecipe.GetDyeCommandOne(m_nDyeID, m_nDyeSeq)

    Set oRecipe = Nothing
    
    ReDim sParam(130)
    ReDim sDye(10)
    ReDim sAux(10)
    
    ' Parameters
    ' 0~9 : 염료명,             10~19: 염료비율,            20~29: 염료사용량,
    ' 30~39: 추가1회 사용량,    40~49: 추가2회 사용량,      50~59: 추가3회 사용량
    ' 60~69: 조제명,            70~79: 조제비율,            80~89: 조제사용량
    ' 90~99: 추가1회 사용량,    100~109: 추가2회 사용량,    110~119: 추가3회 사용량
    ' 120~129: 카드내역
    
    With grdDyeAux(0)
        For i = 0 To .Rows - 2
            sParam(i) = .TextMatrix(i + 1, 1)   ' 염료명
            sDye(i) = .TextMatrix(i + 1, 5)     ' 염료코드 배열(염료, 추가제거분에 대해 위치 찾기위함)
            nDyeCnt = nDyeCnt + 1
            
            sParam(i + 10) = Format(.TextMatrix(i + 1, 2), "#0.000000")  ' 투입비율
            sParam(i + 20) = Format(.TextMatrix(i + 1, 4), "#####0.00")  ' 염료 투입량
        Next i
    End With
    
    With grdDyeAux(1)
        For i = 0 To .Rows - 2
            sParam(i + 60) = .TextMatrix(i + 1, 1)  ' 조제명
            sAux(i) = .TextMatrix(i + 1, 5)         ' 조제코드 배열(염료, 추가제거분에 대해 위치 찾기위함)
            nAuxCnt = nAuxCnt + 1
            
            sParam(i + 70) = Format(.TextMatrix(i + 1, 2), "#0.000000")  ' 투입비율
            sParam(i + 80) = Format(.TextMatrix(i + 1, 4), "#####0.00")  ' 조제 투입량
        Next i
    End With
    
    If m_nDyeSeq > 1 Then
        For i = 2 To m_nDyeSeq
            ' 추가작업 염료 투입량
            With grdDye(i - 2)
                ' 추가작업 염료 그리드 항목 Loop
                For j = 1 To .Rows - 1
                    ' 동일 염료의 출력위치 확인
                    For k = 0 To 9
                        ' 기존 염료 내역중에서 현재 염료 항목의 위치를 찾음 - 출력위치 지정
                        ' 해당 출력 위치에 염료 투입량 입력
                        If sDye(k) = .TextMatrix(j, 4) Then
                            sParam(i * 10 + 10 + k) = Format(.TextMatrix(j, 3), "#####0.00")
                            bFind = True    ' 기존 염료내역 중에서 현재 염료항목을 찾았음을 의미
                        End If
                    Next k
                    
                    ' 기존 염료 내역에서 현재 염료 항목을 찾지 못함
                    ' 기존 염료 내역에 없는 염료항목이 추가되어있을 경우
                    If bFind = False Then
                        ' 새로운 염료 항목을 출력할 염료 항목에 추가
                        sParam(nDyeCnt) = .TextMatrix(j, 1)     ' 염료명
                        sParam(i * 10 + 10 + nDyeCnt) = Format(.TextMatrix(j, 3), "#####0.00")  ' 투입량
                        nDyeCnt = nDyeCnt + 1
                    Else
                        bFind = False
                    End If
                    
                Next j ' 추가작업 염료 그리드 항목 Loop
            
            End With
            
            
            ' 추가작업 조제량 투입량
            With grdAux(i - 2)
                ' 추가작업 조제 그리드 항목 Loop
                For j = 1 To .Rows - 1
                    ' 동일 조제의 출력위치 확인
                    For k = 0 To 9
                        ' 기존 조제 내역중에서 현재 조제 항목의 위치를 찾음 - 출력위치 지정
                        ' 해당 출력 위치에 염료 투입량 입력
                        If sAux(k) = .TextMatrix(j, 4) Then
                            sParam(i * 10 + 70 + k) = Format(.TextMatrix(j, 3), "#####0.00")
                            bFind = True    ' 기존 조제내역 중에서 현재 조제항목을 찾았음을 의미
                        End If
                    Next k
                    
                    ' 기존 조제 내역에서 현재 염료 항목을 찾지 못함
                    ' 기존 조제 내역에 없는 염료항목이 추가되어있을 경우
                    If bFind = False Then
                        ' 새로운 조제 항목을 출력할 조제 항목에 추가
                        sParam(nAuxCnt + 60) = .TextMatrix(j, 1)    ' 조제명
                        sParam(i * 10 + 70 + nAuxCnt) = Format(.TextMatrix(j, 3), "#####0.00")  ' 투입량
                        nAuxCnt = nAuxCnt + 1
                    Else
                        bFind = False
                    End If
                    
                Next j ' 추가작업 조제 그리드 항목 Loop
            
            End With
        
        Next i  ' 염색 차수 m_nDyeSeq
        
    End If
    
    With grdCard
        For i = 1 To .Rows - 1
            sCard = .TextMatrix(i, 1)
            sCard = sCard & IIf(Len(Trim(.TextMatrix(i, 2))) = 0, "", "(" & Trim(.TextMatrix(i, 2)) & ")")
            sParam(119 + i) = sCard
        Next i
    
    End With
    
    sParam(130) = txtRPCalcRemark.Text
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "cmdPrint_Click", Err.Description)
End Sub

' 처방전 변경
Private Sub cmdRecipe_Click()
    Dim nRecipeCnt%

    ' 최초작업시에만 처방전 변경 가능
    If m_nDyeSeq > 1 Then
        MessageBox "최초작업시에만 처방전 변경 가능합니다"
        Exit Sub
    End If
    nRecipeCnt = GetRecipeCount  ' 처방전 갯수 파악.
            
    If nRecipeCnt = 0 Then
        MessageBox "처방전이 내려지지 않았습니다. 확인해 주십시오"
    
        Exit Sub
    End If
    
    If nRecipeCnt > 1 Then
        Call GetRecipeDataAll
    End If
    
End Sub


Private Sub cmdRecipeCal_Click()
    Dim i%, nSeq%, nRoll%, nQty%
    Dim nDyeSchID%, nDyeSeq%
    
    With grdData
        .Redraw = flexRDNone
        grdCardList.Rows = grdCardList.FixedRows
        If .Rows = .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                .Cell(flexcpChecked, i, 1) = flexUnchecked
                
                grdCardList.AddItem ""
                grdCardList.TextMatrix(grdCardList.Rows - 1, 3) = .TextMatrix(i, 23)    '밧자번호
                grdCardList.TextMatrix(grdCardList.Rows - 1, 4) = nSeq + 1
                grdCardList.TextMatrix(grdCardList.Rows - 1, 5) = .TextMatrix(i, 2)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 6) = .TextMatrix(i, 3)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 7) = .TextMatrix(i, 8)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 8) = .TextMatrix(i, 4)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 9) = .TextMatrix(i, 6)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 10) = .TextMatrix(i, 7)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 11) = .TextMatrix(i, 13)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 12) = .TextMatrix(i, 9)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 13) = .TextMatrix(i, 10)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 17) = MakeCardID(.TextMatrix(i, 6), OM_REDUCE)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 18) = .TextMatrix(i, 7)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 20) = MakeOrderID(.TextMatrix(i, 4), OM_REDUCE)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 21) = .TextMatrix(i, 15)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 23) = .TextMatrix(i, 24)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 24) = .TextMatrix(i, 25)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 25) = .TextMatrix(i, 26)
                
                nRoll = nRoll + CheckNum(.TextMatrix(i, 9))
                nQty = nQty + CheckNum(.TextMatrix(i, 10))
                nSeq = nSeq + 1
            End If
        Next i
        .Redraw = flexRDDirect
    End With

    grdCardList.Rows = grdCardList.Rows + 1
    grdCardList.RowHeight(grdCardList.Rows - 1) = 300
    grdCardList.Cell(flexcpText, grdCardList.Rows - 1, 0, grdCardList.Rows - 1, 11) = "선택되어진 카드 총 합계"
    grdCardList.Cell(flexcpFontBold, grdCardList.Rows - 1, 0, grdCardList.Rows - 1, grdCardList.Cols - 1) = True
    grdCardList.TextMatrix(grdCardList.Rows - 1, 12) = Format(nRoll, "#,##0")
    grdCardList.TextMatrix(grdCardList.Rows - 1, 13) = Format(nQty, "#,###,##0")
    grdCardList.MergeCells = flexMergeRestrictRows
    grdCardList.MergeRow(grdCardList.Rows - 1) = True
    grdCardList.Row = grdCardList.FixedRows
    
    If nSeq = 0 Then
        MsgBox "평량지시를 내릴 카드를 선택한 후 평량지시 버턴을 내려주십시오", vbInformation
    End If
End Sub

Private Sub cmdRemarkCopy_Click()
    txtRPCalcRemark = txtRecipeRemark
End Sub

Private Sub cmdSearch_Click()
    If stTab.Tab = 0 Then
        Call FillGridData
    Else
        Call FillGridRecipeCalc
    End If
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
    Unload Me
End Sub

Private Sub grdAux_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i%, nValue!, nSetQty!, nRate!
    Dim nAuxQty As Single
    Dim nDyeRate As Single, nWaterRate As Single
    Dim nDyeAuxQty As Single, nDyeAuxRate As Single
    
    nValue = CSng(IIf(IsNumeric(txtWght), txtWght, "0"))    ' 단위중량
    nSetQty = CSng(txtINQty)                                ' 투입수량
    nWaterRate = CSng(txtWaterRate)
    
    If Col < 2 Or Col > 3 Then

    Else ' 염료 추가
        
        With grdAux(Index)
            If Not IsNumeric(.TextMatrix(Row, Col)) Then Exit Sub
    
            ' 초기 투입비율 입력
            If Col = 2 Then
                ' 투입비율
                ' 투입비율 * 호기별 액비
                nAuxQty = CSng(.TextMatrix(Row, 2)) * nWaterRate
                
                .TextMatrix(Row, Col + 1) = SetCurrency(nAuxQty, 2)
                
            ' 투입량 입력
            ' 비율 = 투입량  / 액비
            ElseIf Col = 3 Then
                nDyeAuxRate = CSng(.TextMatrix(Row, 3)) / nWaterRate
                
                .TextMatrix(Row, Col - 1) = SetCurrency(nDyeAuxRate, 6)
            End If
            
        End With
    End If

End Sub

Private Sub grdAux_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Index <> m_nDyeSeq - 2 Then
    
        Cancel = True
    End If
    
    If Col < 2 Or Col > 3 Then
        Cancel = True
    End If
    
End Sub

Private Sub grdData_Click()
    With grdData
        If .Row < .FixedRows Or .Col <> 1 Then Exit Sub
        
        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, 1) = flexChecked
        End If
    End With
End Sub

Private Sub grdDye_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i%, nValue!, nSetQty!, nRate!
    Dim nDyeQty As Single
    Dim nDyeRate As Single, nWaterRate As Single
    
    nValue = CSng(IIf(IsNumeric(txtWght), txtWght, "0"))    ' 단위중량
    nSetQty = CSng(txtINQty)                                ' 투입수량
    nWaterRate = CSng(txtWaterRate)
    
    If Col < 2 Or Col > 3 Then

    Else ' 염료 추가
        
        With grdDye(Index)
            If Not IsNumeric(.TextMatrix(Row, Col)) Then Exit Sub

            ' 투입비율 입력
            If Col = 2 Then
                ' 투입비율
                ' 투입비율 * 단위중량 * 투입수량 / 100
                nDyeQty = CSng(.TextMatrix(Row, 2)) * nValue * nSetQty / 100
                
                .TextMatrix(Row, Col + 1) = SetCurrency(nDyeQty, 2)
                
            ' 투입량 입력
            ' 비율 = 투입량 * 100 / (단위중량 * 수량)
            ElseIf Col = 3 Then
                nDyeRate = (CSng(.TextMatrix(Row, 3)) * 100) / (nValue * nSetQty)
                
                .TextMatrix(Row, Col - 1) = SetCurrency(nDyeRate, 4)
            End If
            
        End With
    End If

End Sub

Private Sub grdDye_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Index <> m_nDyeSeq - 2 Then
    
        Cancel = True
    End If
    
    If Col < 2 Or Col > 3 Then
        Cancel = True
    End If
    
End Sub



Private Sub grdDyeAux_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i%, nValue!, nSetQty!, nRate!
    Dim nDyeAuxQty As Single
    Dim nDyeAuxRate As Single, nWaterRate As Single
    
    nValue = CSng(IIf(IsNumeric(txtWght), txtWght, "0"))    ' 단위중량
    nSetQty = CSng(txtINQty)                                ' 투입수량
    nWaterRate = CSng(txtWaterRate)
    
    If Col < 3 Or Col > 4 Then

    Else ' 염료 추가
        If Index = 0 Then
            With grdDyeAux(0)
                If Not IsNumeric(.TextMatrix(Row, Col)) Then Exit Sub
    
                ' 초기 투입비율 입력
                If Col = 3 Then
                    ' 투입비율
                    ' 투입비율 * 단위중량 * 투입수량 / 100
                    nDyeAuxQty = CSng(.TextMatrix(Row, 3)) * nValue * nSetQty / 100
                    
                    .TextMatrix(Row, Col + 1) = SetCurrency(nDyeAuxQty, 2)
                    
                ' 투입량 입력
                ' 비율 = 투입량 * 100 / (단위중량 * 수량)
                ElseIf Col = 4 Then
                    nDyeAuxRate = (CSng(.TextMatrix(Row, 4)) * 100) / (nValue * nSetQty)
                    
                    .TextMatrix(Row, Col - 1) = SetCurrency(nDyeAuxRate, 6)
                End If
                
                .Row = Row
                .Col = 3
                .CellBackColor = IIf(CSng(.TextMatrix(Row, 2)) <> CSng(.TextMatrix(Row, 3)), vbRed, vbWhite)
            
            End With
        
        Else    ' 조제
            With grdDyeAux(1)
                If Not IsNumeric(.TextMatrix(Row, Col)) Then Exit Sub
    
                ' 초기 투입비율 입력
                If Col = 3 Then
                    ' 투입비율
                    ' 투입비율 * 단위중량 * 투입수량 / 100
                    nDyeAuxQty = CSng(.TextMatrix(Row, 3)) * nWaterRate
                    
                    .TextMatrix(Row, Col + 1) = SetCurrency(nDyeAuxQty, 2)
                    
                ' 투입량 입력
                ' 비율 = 투입량 * 100 / (단위중량 * 수량)
                ElseIf Col = 4 Then
                    nDyeAuxRate = CSng(.TextMatrix(Row, 4)) / nWaterRate
                    
                    .TextMatrix(Row, Col - 1) = SetCurrency(nDyeAuxRate, 6)
                End If
                
                .Row = Row
                .Col = 3
                .CellBackColor = IIf(CSng(.TextMatrix(Row, 2)) <> CSng(.TextMatrix(Row, 3)), vbRed, vbWhite)
            
            End With
        End If
    End If
End Sub


Private Sub grdDyeAux_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' 실제 투입비율및 조제 투입비율만 수정 가능
    If Col < 3 Or Col > 4 Then
        Cancel = True
    End If
          
End Sub




Private Sub grdRecipeCalc_RowColChange()
    If m_bLoading1 Then Exit Sub
    Call ShowData
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
        If optOrder(0).Value Then
            .ColWidth(5) = 1350
            .ColWidth(4) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(5) = 0
            .ColWidth(4) = 1350
            chkSearch(3).Caption = "관리번호"
        End If
    End With
End Sub

Private Sub stTab_Click(PreviousTab As Integer)
    If stTab.Tab = 0 Then
        cmdOK.Caption = "평량지시작성"
        cmdOK.Tag = "작성"
        cmdAddRecipe.Visible = False
        cmdDelete.Visible = False
        cmdRecipeCal.Visible = True
        
        grdCardList.Rows = grdCardList.FixedRows
        
        Call FillGridData

    Else
        cmdOK.Caption = "평량지시수정"
        cmdOK.Tag = "수정"
        cmdAddRecipe.Visible = True
        cmdDelete.Visible = True
        cmdRecipeCal.Visible = False
        
        grdCardList.Rows = grdCardList.FixedRows
        
        Call FillGridRecipeCalc
    End If
End Sub

Private Sub txtPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        Call ReturnCode(LG_PERSON, , False, txtPerson)

        txtWght.SetFocus
    End If
End Sub


Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15460, 9740
    
    Call InitGrid
    Call MakeProcessCombo
    Call SetOperate(Me)
    Call AddLstBox

    cmdSelect.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdSave.Picture = LoadResPicture("SELECT", vbResIcon)
    For i = 0 To 2
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
    Next i
    cmdFind(0).Enabled = True
    
    For i = 1 To 5
        txtSearch(i).Enabled = False
    Next i
    
    cmdAddRecipe.Visible = False
    cmdDelete.Visible = False
    cboProcess.Enabled = False
    
    pnlProgress.Visible = False
    
    Call ClearData

    grdDyeAux(0).Editable = flexEDKbdMouse
    grdDyeAux(1).Editable = flexEDKbdMouse
    
    m_bSaved = False
    m_iFlag = 3
      
    Me.Show
End Sub


Public Sub SetInstruction(Optional nDyeID As Long, Optional nDyeSeq As Integer)
    Dim nRecipeCnt%
    Dim sTitle$
    
    m_nDyeID = nDyeID
    m_nDyeSeq = nDyeSeq
    
    txtDyeID = Format(m_nDyeID, "000000000") & " (" & CStr(m_nDyeSeq) & ")"
    
    If m_nDyeSeq <= 1 Then
        pnlCalc(0).Enabled = True
        pnlCalc(1).Enabled = False
        tabDye.Tab = 0
        tabAux.Tab = 0
    Else
        pnlCalc(0).Enabled = False
        pnlCalc(1).Enabled = True
        
    End If
    
    If m_nDyeID = 0 Then
        Call ShowDyeList
    Else
        ' 염색 지시내역
        If ShowDyeCommand(m_nDyeID, m_nDyeSeq) = True Then
            ' 염색지시 카드내역
            Call ShowCardList(m_nDyeID, m_nDyeSeq)
            
            ' 기존에 평량지시가 내려져 있지 않은경우
            ' 평량지시 신규작성
            If m_bModify = False Then
            
                '추가작업의 경우 신규작성이라도
                ' 이전 작업의 평량지시내역 출력
                If m_nDyeSeq > 1 Then
                    Call ShowMatchData(m_nDyeID, m_nDyeSeq)
                Else
                    nRecipeCnt = GetRecipeCount  ' 처방전 갯수 파악.
                    
                    If nRecipeCnt = 0 Then
                        MessageBox "처방전이 내려지지 않았습니다. 확인해 주십시오"
                    
                        Exit Sub
                    End If
                    
                    If nRecipeCnt > 1 Then
                        Call GetRecipeDataAll
                    Else
                        Call GetRecipeData(0)
                    End If
                End If
                
            Else
                ' 기존에 평량 지시가 내려진 경우
                ' 기존 평량지시 변경
                Call ShowMatchData(m_nDyeID, m_nDyeSeq)
                
            End If
            
            If m_nDyeSeq = 1 Then
                sTitle = "염조제 투입량  - 본작업 평량지시"
                tabDye.Tab = 0
                tabAux.Tab = 0
            ElseIf m_nDyeSeq = 2 Then
                sTitle = "염조제 투입량  - 추가 1회 평량지시"
                tabDye.Tab = m_nDyeSeq - 2
                tabAux.Tab = m_nDyeSeq - 2
            ElseIf m_nDyeSeq = 3 Then
                sTitle = "염조제 투입량  - 추가 2회 평량지시"
                tabDye.Tab = m_nDyeSeq - 2
                tabAux.Tab = m_nDyeSeq - 2
            ElseIf m_nDyeSeq = 4 Then
                sTitle = "염조제 투입량  - 추가 3회 평량지시"
                tabDye.Tab = m_nDyeSeq - 2
                tabAux.Tab = m_nDyeSeq - 2
            End If
            
            pnlTitle.Caption = "   " & sTitle & IIf(m_bModify = True, " 수정", " 작성")
            
        Else
            Call ClearData
        End If
    End If
End Sub



Private Sub ClearData()
    Dim i%
    
    grdCard.Rows = grdCard.FixedRows
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows
    
    For i = 0 To 2
        grdDye(i).Rows = grdDye(i).FixedRows
        grdAux(i).Rows = grdAux(i).FixedRows

    Next i

    txtCustom = ""
    txtCustom.Tag = ""
    txtArticle = ""
    txtArticle.Tag = ""
    txtRemark = ""
    txtRecipeRemark = ""
    txtRPCalcRemark = ""
    txtWorkClss = ""

    txtColor = ""
    txtColor.Tag = ""
    txtMachine = ""
    txtRoll = 0
    txtINQty = 0
    txtWght = 0
    txtPerson = ""
    txtPerson.Tag = ""

    txtRecipePerson = ""
    txtRecipePerson.Tag = ""
    txtRecipeNO = ""
    txtRecipeSeq = ""
    txtModifySeq = ""

End Sub


Private Sub InitGrid()
    Dim i%

    With grdData
        .Redraw = flexRDNone
        .Cols = 27
        
        Call SetVSFlexGrid(grdData)
        .Rows = 1
        .RowHeightMin = 390
        
        .TextArray(0) = " ":
        .TextArray(1) = " ":            .ColWidth(1) = 250:     .ColDataType(1) = flexDTBoolean
        .TextArray(2) = "거래처":       .ColWidth(2) = 1000:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":         .ColWidth(3) = 1800:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":     .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNo":      .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "카드번호":     .ColWidth(6) = 1000:               .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "분할" & vbCrLf & "번호":     .ColWidth(7) = 500:            .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":         .ColWidth(8) = 1300:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "절수":         .ColWidth(9) = 500:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "수량":         .ColWidth(10) = 600:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "완료공정":    .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "대기공정":    .ColWidth(12) = 900:           .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "카드상태":    .ColWidth(13) = 900:           .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "계획공정":    .ColWidth(14) = 7000:             .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "색상코드":    .ColHidden(15) = True   '.ColWidth(15) = 0
        .TextArray(16) = "재투입구분":  .ColHidden(16) = True '.ColWidth(16) = 0
        .TextArray(17) = "긴급구분":    .ColHidden(17) = True '.ColWidth(17) = 0
        .TextArray(18) = "공정패턴":    .ColHidden(18) = True
        .TextArray(19) = "사종":        .ColHidden(19) = True
        .TextArray(20) = "제직처":      .ColHidden(20) = True
        .TextArray(21) = "생지폭":      .ColHidden(21) = True
        .TextArray(22) = "생지밀도":    .ColHidden(22) = True
        .TextArray(23) = "밧자번호":    .ColHidden(23) = True
        .TextArray(24) = "스케쥴번호":  .ColHidden(24) = True
        .TextArray(25) = "차수":        .ColHidden(25) = True
        .TextArray(26) = "단위":        .ColHidden(26) = True
        
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With

    With grdRecipeCalc
        .Redraw = flexRDNone
        .Cols = 32
        
        Call SetVSFlexGrid(grdRecipeCalc)
        .Rows = 1
        .RowHeightMin = 390
        
        .TextArray(0) = " ":             .ColWidth(0) = 0
        .TextArray(1) = " ":            .ColWidth(1) = 250 ':             .ColDataType(1) = flexDTBoolean
        .TextArray(2) = "거래처":       .ColWidth(2) = 1000:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":         .ColWidth(3) = 1800:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":     .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNo":      .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "카드번호":     .ColWidth(6) = 1000:               .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "분할" & vbCrLf & "번호":     .ColWidth(7) = 500:            .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":         .ColWidth(8) = 1300:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "절수":         .ColWidth(9) = 500:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "수량":         .ColWidth(10) = 600:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "완료공정":    .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "대기공정":    .ColWidth(12) = 900:           .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "카드상태":    .ColWidth(13) = 900:           .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "계획공정":    .ColWidth(14) = 7000:             .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "색상코드":    .ColHidden(15) = True   '.ColWidth(15) = 0
        .TextArray(16) = "재투입구분":  .ColHidden(16) = True '.ColWidth(16) = 0
        .TextArray(17) = "긴급구분":    .ColHidden(17) = True '.ColWidth(17) = 0
        .TextArray(18) = "공정패턴":    .ColHidden(18) = True
        .TextArray(19) = "사종":        .ColHidden(19) = True
        .TextArray(20) = "제직처":      .ColHidden(20) = True
        .TextArray(21) = "생지폭":      .ColHidden(21) = True
        .TextArray(22) = "생지밀도":    .ColHidden(22) = True
        .TextArray(23) = "밧자번호":    .ColHidden(23) = True
        .TextArray(24) = "스케쥴번호":  .ColHidden(24) = True
        .TextArray(25) = "차수":        .ColHidden(25) = True
        .TextArray(26) = "단위":        .ColHidden(26) = True
        .TextArray(27) = "카드갯수":    .ColHidden(27) = True
        .TextArray(28) = "염색호기":    .ColHidden(28) = True
        .TextArray(29) = "작업구분":    .ColHidden(29) = True
        .TextArray(30) = "염색구분":    .ColHidden(30) = True
        .TextArray(31) = "염색패턴":    .ColHidden(31) = True
        
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With


    With grdCardList
        .Redraw = flexRDNone
        Call SetVSFlexGrid(grdCardList)
        
        .WordWrap = False

        .Rows = 1:      .Cols = 26

        .FixedRows = 1:     .FixedCols = 0
        
        .TextArray(0) = "":                     .ColWidth(0) = 0:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "밧자번호":             .ColWidth(1) = 0:           .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "밧자순위":             .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "밧자":                 .ColWidth(3) = 500:         .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "No":                   .ColWidth(4) = 300:         .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "거래처":               .ColWidth(5) = 1100:        .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "품명":                 .ColWidth(6) = 2300:        .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "색상":                 .ColWidth(7) = 1800:        .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "관리번호":             .ColWidth(8) = 0:           .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "카드번호":             .ColWidth(9) = 1000:        .ColAlignment(9) = flexAlignLeftCenter
        .TextArray(10) = "분할":                .ColWidth(10) = 500:        .ColAlignment(10) = flexAlignLeftCenter
        .TextArray(11) = "대기":                .ColWidth(11) = 0:        .ColAlignment(11) = flexAlignLeftCenter
        .TextArray(12) = "절수":                .ColWidth(12) = 600:        .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "수량":                .ColWidth(13) = 800:        .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "거래처코드":          .ColWidth(14) = 0:          .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "품명코드":            .ColWidth(15) = 0:          .ColAlignment(15) = flexAlignLeftCenter
        .TextArray(16) = "색상코드":            .ColWidth(16) = 0:          .ColAlignment(16) = flexAlignLeftCenter
        .TextArray(17) = "카드번호":            .ColWidth(17) = 0:          .ColAlignment(17) = flexAlignCenterCenter
        .TextArray(18) = "분할":                .ColWidth(18) = 0:          .ColAlignment(18) = flexAlignLeftCenter
        .TextArray(19) = "제직처":              .ColWidth(19) = 0:          .ColAlignment(19) = flexAlignLeftCenter
        .TextArray(20) = "관리번호":            .ColWidth(20) = 0:          .ColAlignment(20) = flexAlignLeftCenter
        .TextArray(21) = "OrderSeq":            .ColWidth(21) = 0:        .ColAlignment(21) = flexAlignLeftCenter
        .TextArray(22) = "계획 후공정":         .ColWidth(22) = 0:       .ColAlignment(22) = flexAlignLeftCenter
        .TextArray(23) = "스케쥴번호":          .ColWidth(23) = 0:          .ColAlignment(23) = flexAlignLeftCenter
        .TextArray(24) = "차수":                .ColWidth(24) = 0:          .ColAlignment(24) = flexAlignLeftCenter
        .TextArray(25) = "단위":                .ColWidth(25) = 0:          .ColAlignment(25) = flexAlignLeftCenter
        
        .ColFormat(12) = "#,###"
        .ColFormat(13) = "#,###"
        
        .Redraw = flexRDDirect
    End With
    
    With grdCard
        .Cols = 10

        Call SetVSFlexGrid(grdCard)

        .Redraw = flexRDNone

        .TextArray(1) = "카드번호":                 .ColWidth(1) = 1500:      .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "분할" & vbCrLf & "번호":   .ColWidth(2) = 600:       .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "관리번호":                 .ColWidth(3) = 1300:      .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "Order NO":                 .ColWidth(4) = 0
        .TextArray(5) = "투입" & vbCrLf & "절수":   .ColWidth(5) = 600
        .TextArray(6) = "투입" & vbCrLf & "수량":   .ColWidth(6) = 800
        .TextArray(7) = "단위":                     .ColWidth(7) = 600
        .TextArray(8) = "UnitClss":                 .ColWidth(8) = 0
        .TextArray(9) = "색상명":                   .ColWidth(9) = 1000

        .ExtendLastCol = True
        .WordWrap = False

        .Redraw = flexRDDirect
    End With

    With grdDyeAux(0)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(0))

        .Redraw = flexRDNone

        .TextArray(1) = "염료명":                           .ColWidth(1) = 3000:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "실험실" & vbCrLf & "투입비율":     .ColWidth(2) = 1400:     .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "투입비율" & vbCrLf & "(%)":        .ColWidth(3) = 1400:     .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "염료투입량":                       .ColWidth(4) = 1010:    .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "염료코드":                         .ColWidth(5) = 0

        .ColFormat(2) = "#,##0.000000"
        .ColFormat(3) = "#,##0.000000"
        .ColFormat(4) = "#,##0.00"
                
        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusHeavy
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False

        .WordWrap = False

        .Redraw = flexRDDirect
    End With

    With grdDyeAux(1)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(1))

        .Redraw = flexRDNone

        .TextArray(1) = "조제명":                           .ColWidth(1) = 3000:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "실험실" & vbCrLf & "투입비율":     .ColWidth(2) = 1400:         .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "투입비율" & vbCrLf & "(g/ℓ)":     .ColWidth(3) = 1400:         .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "조제투입량":                       .ColWidth(4) = 1010:        .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "조제코드":                             .ColWidth(5) = 0

        .ColFormat(2) = "#,##0.000000"
        .ColFormat(3) = "#,##0.000000"
        .ColFormat(4) = "#,##0.00"
        
        .Editable = flexEDKbdMouse
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusHeavy
        
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False


        .WordWrap = False

        .Redraw = flexRDDirect
    End With


    With grdRecipe
        .Cols = 6
        Call SetVSFlexGrid(grdRecipe)

        .Redraw = flexRDNone

        .TextArray(0) = "처방순위":     .ColWidth(0) = 900:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "염료순위":     .ColWidth(1) = 950:         .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "염조제명":     .ColWidth(2) = 1500:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "투입비율":     .ColWidth(3) = 1000:        .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "비고사항":     .ColWidth(4) = 1000:        .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "단위중량":     .ColWidth(5) = 0

        .ColFormat(3) = "#,##0.000000"


        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone

        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .MergeCol(4) = True

        .WordWrap = False

        .Redraw = flexRDDirect
    End With


    For i = 0 To 2
        
        With grdDye(i)
            .Cols = 5
            Call SetVSFlexGrid(grdDye(i))
    
            .Redraw = flexRDNone
    
            .TextArray(1) = "염료명":                           .ColWidth(1) = 3000:        .ColAlignment(1) = flexAlignLeftCenter
            .TextArray(2) = "투입비율" & vbCrLf & "(%)":        .ColWidth(2) = 1500:         .ColAlignment(2) = flexAlignRightCenter
            .TextArray(3) = "염료투입량":                       .ColWidth(3) = 1010:        .ColAlignment(3) = flexAlignRightCenter
            .TextArray(4) = "염료코드":                         .ColWidth(4) = 0
    
            .ColFormat(2) = "#,##0.000000"
            .ColFormat(3) = "#,##0.00"
                        
            .ExtendLastCol = True
            .Editable = flexEDKbdMouse
            .HighLight = flexHighlightWithFocus
            .FocusRect = flexFocusHeavy
            
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    
    
            .WordWrap = False
    
            .Redraw = flexRDDirect
        End With
        
        With grdAux(i)
            .Cols = 5
            Call SetVSFlexGrid(grdAux(i))
    
            .Redraw = flexRDNone
    
            .TextArray(1) = "조제명":                           .ColWidth(1) = 3000:        .ColAlignment(1) = flexAlignLeftCenter
            .TextArray(2) = "투입비율" & vbCrLf & "(%)":        .ColWidth(2) = 1500:         .ColAlignment(2) = flexAlignRightCenter
            .TextArray(3) = "조제투입량":                       .ColWidth(3) = 1010:        .ColAlignment(3) = flexAlignRightCenter
            .TextArray(4) = "조제코드":                         .ColWidth(4) = 0
    
            .ColFormat(2) = "#,##0.000000"
            .ColFormat(3) = "#,##0.00"
                        
            .ExtendLastCol = True
            .Editable = flexEDKbdMouse
            .HighLight = flexHighlightWithFocus
            .FocusRect = flexFocusHeavy
            
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
        
            .WordWrap = False
    
            .Redraw = flexRDDirect
        End With
        
    Next i


End Sub

Private Sub MakeProcessCombo()
    Dim oCard As PlusLib2.CCard
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bLoading1 = True
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon

    Set rs = oCard.GetProcess(1)
    Set oCard = Nothing

    With cboProcess
        .Clear

        Do Until rs.EOF
            .AddItem CStr(rs!Process)
            .ItemData(.NewIndex) = CLng(Left(rs!processid, 2))
            
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing

        If .ListCount > 0 Then .ListIndex = 0
    End With

    m_bLoading1 = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    m_bLoading1 = False
    Call ErrorBox(Err.Number, "frmRecipeCalc.MakeProcessCombo", Err.Description)
End Sub


' 처방전의 갯수 확인
Private Function GetRecipeCount() As Integer
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim sOrder$, nOrderSeq%

    GetRecipeCount = 0

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    If grdCard.Rows = grdCard.FixedRows Then Exit Function

    'If Len(grdCard.TextMatrix(grdCard.Row, 11)) = 0 Then Exit Function

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    With grdCard
        sOrder = MakeOrderID(.TextMatrix(.FixedRows, 3), OM_REDUCE)
        nOrderSeq = txtColor.Tag
    End With

    Set rs = oRecipe.GetRecipeCount(sOrder, nOrderSeq)
            
    GetRecipeCount = rs.RecordCount

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set oRecipe = Nothing
    Call ErrorBox(Err.Number, "frmRecipeCalc.GetRecipeDataAll", Err.Description)

End Function



' 모든 처방내역 불러오기
Private Function GetRecipeDataAll() As Boolean
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim sOrderID$, nOrderSeq%

    GetRecipeDataAll = False

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    With grdCard
        sOrderID = MakeOrderID(.TextMatrix(.FixedRows, 3), OM_REDUCE)
        nOrderSeq = txtColor.Tag
    End With

    Dim i%

    Set rs = oRecipe.GetRecipeSubAll(sOrderID, nOrderSeq)
    With grdRecipe
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem rs!RecipeSeq & vbTab & rs!DyeAuxSeq & vbTab & rs!DyeAux & vbTab & SetCurrency(rs!DyeAuxRate, 6) & vbTab & _
                    IIf(IsNull(rs!Remark) Or Len(Trim(rs!Remark)) = 0, " ", rs!Remark) & vbTab & rs!UnitWght

            rs.MoveNext
        Next i
        rs.Close

        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If

        .Redraw = flexRDDirect
    End With

    pnlRecipe.Move 2565, 1980
    pnlRecipe.Visible = True

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    GetRecipeDataAll = True

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set oRecipe = Nothing
    Call ErrorBox(Err.Number, "frmRecipeCalc.GetRecipeDataAll", Err.Description)

End Function



' 처방전 내역 불러오기
Private Function GetRecipeData(nChkReworkSeq As Integer) As Boolean
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim rsSub   As Recordset
    Dim sOrder$, nOrderSeq%
    Dim i%
    
    GetRecipeData = False

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    If grdCard.Rows = grdCard.FixedRows Then Exit Function

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    With grdCard
        sOrder = MakeOrderID(.TextMatrix(.FixedRows, 3), OM_REDUCE)
        nOrderSeq = txtColor.Tag
    End With

    ' 염색 본작업
    If m_nDyeSeq <= 1 Then
        ' 처방전 여러종류일 경우 변경순위 받음
        ' 해당 처방전의 처방자, 처방번호, 처방순위등 받음..
        Set rs = oRecipe.GetRecipeByColor(sOrder, nOrderSeq, nChkReworkSeq, IIf(nChkReworkSeq = 0, 0, txtRecipeSeq))
    
        If rs.EOF Then
            Call ClearGrid
    
            Screen.MousePointer = vbDefault
            rs.Close
            Set rs = Nothing
            Set oRecipe = Nothing
    
            MessageBox "처방전이 내려지지 않았습니다"
    
            Exit Function
        End If
    
        txtRecipePerson = CheckNull(rs!Person)
        txtRecipePerson.Tag = CheckNull(rs!PersonID)
        txtRecipeSeq = CheckNull(rs!RecipeSeq)
        txtRecipeNO = CheckNull(rs!RecipeNO)
        txtModifySeq = CheckNull(rs!ModifySeq)
        txtOrderID = CheckNull(rs!OrderID)
        txtWght = rs!UnitWght
        txtRecipeRemark = CheckNull(rs!Remark)

        
        If rs!UnitWght = 0 Then
            Call MessageBox("단위중량이 설정되지 않았습니다" & vbCrLf & vbCrLf & "실험실로 문의하시기 바랍니다", vbInformation)
        
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
            
        End If
        
        rs.Close
    
        Set rs = Nothing
        
        
        '처방전의 상세내역
       ' sOrderID , sColorID, nReworkSeq, 1, "1"
        Set rsSub = oRecipe.GetRecipeSubByRecipeSeq(sOrder, nOrderSeq, nChkReworkSeq, IIf(nChkReworkSeq = 0, 0, txtRecipeSeq), 1, "1")
        With grdDyeAux(0)
            .Redraw = flexRDNone
    
            .Rows = .FixedRows
            For i = 1 To rsSub.RecordCount
                ' 염료 투입량 = 염료 투입비율 * 단위중량 * 투입수량 / 100
                .AddItem CStr(i) & vbTab & rsSub!DyeAux & vbTab & SetCurrency(rsSub!DyeAuxRate, 6) & vbTab & SetCurrency(rsSub!DyeAuxRate, 6) & vbTab & _
                    SetCurrency(rsSub!DyeAuxRate * CSng(txtWght) * CSng(txtINQty) / 100, 2) & vbTab & rsSub!DyeAuxID
    
                rsSub.MoveNext
            Next i
            rsSub.Close
    
            .Redraw = flexRDDirect
        End With
    
        Set rsSub = oRecipe.GetRecipeSubByRecipeSeq(sOrder, nOrderSeq, nChkReworkSeq, IIf(nChkReworkSeq = 0, 0, txtRecipeSeq), 1, "0")
        With grdDyeAux(1)
            .Redraw = flexRDNone
    
            .Rows = .FixedRows
            For i = 1 To rsSub.RecordCount
                ' 조제 투입량 = 투입비율 * 염색 호기별 액비
                .AddItem CStr(i) & vbTab & rsSub!DyeAux & vbTab & SetCurrency(rsSub!DyeAuxRate, 6) & vbTab & SetCurrency(rsSub!DyeAuxRate, 6) & vbTab & _
                    SetCurrency(rsSub!DyeAuxRate * CSng(txtWaterRate), 2) & vbTab & rsSub!DyeAuxID
    
                rsSub.MoveNext
            Next i
            rsSub.Close
    
            .Redraw = flexRDDirect
        End With
        
        
    ' 추가작업
    Else
        '처방전의 상세내역
       ' sOrderID , sColorID, nReworkSeq, 1, "1"
        Set rsSub = oRecipe.GetRecipeSubByRecipeSeq(sOrder, nOrderSeq, nChkReworkSeq, IIf(nChkReworkSeq = 0, 0, txtRecipeSeq), 1, "1")
        With grdDye(m_nDyeSeq - 2)
            .Redraw = flexRDNone
    
            .Rows = .FixedRows
            For i = 1 To rsSub.RecordCount
                ' 염료 투입량 = 염료 투입비율 * 단위중량 * 투입수량 / 100
                .AddItem CStr(i) & vbTab & rsSub!DyeAux & vbTab & SetCurrency(rsSub!DyeAuxRate, 6) & vbTab & _
                    SetCurrency(rsSub!DyeAuxRate * CSng(txtWght) * CSng(txtINQty) / 100, 2) & vbTab & rsSub!DyeAuxID
    
                rsSub.MoveNext
            Next i
            rsSub.Close
    
            .Redraw = flexRDDirect
        End With
    
        Set rsSub = oRecipe.GetRecipeSubByRecipeSeq(sOrder, nOrderSeq, nChkReworkSeq, IIf(nChkReworkSeq = 0, 0, txtRecipeSeq), 1, "0")
        With grdAux(m_nDyeSeq - 2)
            .Redraw = flexRDNone
    
            .Rows = .FixedRows
            For i = 1 To rsSub.RecordCount
                ' 조제 투입량 = 투입비율 * 염색 호기별 액비
                .AddItem CStr(i) & vbTab & rsSub!DyeAux & vbTab & SetCurrency(rsSub!DyeAuxRate, 6) & vbTab & _
                    SetCurrency(rsSub!DyeAuxRate * CSng(txtWaterRate), 2) & vbTab & rsSub!DyeAuxID
    
                rsSub.MoveNext
            Next i
            rsSub.Close
    
            .Redraw = flexRDDirect
        End With
    
    End If
    
    
    Set rsSub = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbDefault

    GetRecipeData = True

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set oRecipe = Nothing
    Set rsSub = Nothing

    GetRecipeData = False
    Call ErrorBox(Err.Number, "frmRecipeCalc.GetRecipeData", Err.Description)
End Function


Private Sub ClearGrid()
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows

End Sub


Private Sub cmdSelect_Click()
    If grdRecipe.Row = 0 Then Exit Sub
    
    txtRecipeSeq = grdRecipe.TextMatrix(grdRecipe.Row, 0)

    pnlRecipe.Visible = False
    Call GetRecipeData(1)
End Sub


Private Sub grdRecipe_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub ShowCardList(nDyeID As Long, nSeq As Integer)
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim sCardID$, sColorID$, sCard$
    Dim nRoll%, nQty&

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bLoading = True

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    Set rs = oRecipe.GetRapidCommandSub(nDyeID, nSeq)
    
    Set oRecipe = Nothing

    With grdCard
        .Redraw = flexRDNone
        .Rows = .FixedRows

        Do Until rs.EOF
            
            .AddItem CStr(.Rows) & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    rs!OrderNo & vbTab & rs!Roll & vbTab & SetCurrency(rs!Qty) & vbTab & IIf(rs!UnitClss = 0, "Y", "M") & vbTab & _
                    rs!UnitClss & vbTab & CheckNull(rs!Color)
                    
            rs.MoveNext
        Loop

        .Redraw = flexRDDirect

        m_bLoading = False

        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If

    End With

    rs.Close
    
    m_bLoading = False

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set oRecipe = Nothing
    m_bLoading = False

    Call ErrorBox(Err.Number, "frmRecipeCalc.ShowCardList", Err.Description)


End Sub


' ******************************************************************
' *
' *     염색 지시내역 불러오기
' *
' *     2003-12-02
' *
' *******************************************************************

Private Function ShowDyeCommand(nID As Long, nSeq As Integer) As Boolean
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As ADODB.Recordset
    Dim sMessage$

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    m_bLoading = True

    Set rs = oRecipe.GetDyeCommandOne(nID, nSeq)

    Set oRecipe = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        Exit Function
    End If
    
    txtRemark = CheckNull(rs!Remark)
    txtMachine = rs!wiMachID
    txtPattern = rs!PatternID
    txtWght = rs!UnitWght
    txtRoll = rs!wiRoll
    txtINQty = rs!wiQty
    txtColor = rs!Color
    txtColor.Tag = rs!OrderSeq
    txtWorkClss = rs!RapidClss
    txtCustom = rs!kCustom
    txtArticle = rs!Article
    txtWaterRate = rs!WaterRate
    
    ' 평량지시 변경모드
    ' 기존에 평량지시가 내려져 있을경우 해당 내용을 수정하도록 함
    If CheckNull(rs!instclss) = "*" Then
        m_bModify = True
    Else
        m_bModify = False
    End If
    
    If CheckNull(rs!Complitclss) = "*" Then
'        sMessage = "이미 완료된 작업입니다" & vbCrLf & "평량지시를 내릴 수 없습니다"
'        Call MessageBox(sMessage, vbCritical)
        
        cmdSave.Enabled = False
        cmdRemarkCopy.Enabled = False
        cmdRecipe.Enabled = False
        
        Dim iCount As Integer
        For iCount = 0 To 2
            cmdDyeAdd(iCount).Enabled = False
            cmdDyeDel(iCount).Enabled = False
            cmdAuxAdd(iCount).Enabled = False
            cmdAuxDel(iCount).Enabled = False
        Next iCount

'        ShowDyeCommand = False
'
'        cmdSave.Enabled = False
'
'        rs.Close
'        Set rs = Nothing
'
'        Exit Function
    End If
    
    rs.Close
    Set rs = Nothing
 
    ShowDyeCommand = True
    Exit Function

ErrHandler:
    Screen.MousePointer = vbArrow
    Set rs = Nothing
    Set oRecipe = Nothing
    
    ShowDyeCommand = False
    
    Call ErrorBox(Err.Number, "frmRecipeCalc.ShowDyeCommand", Err.Description)

End Function



' 이전 처방내역 출력
Private Sub ShowMatchData(nDyeID As Long, nDyeSeq As Integer)
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs      As Recordset
    Dim nRow%, i%, nSeq%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    Set rs = oRecipe.GetMatch(nDyeID, 1)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Screen.MousePointer = vbArrow
        Exit Sub
    End If
    
    txtRecipePerson = CheckNull(rs!RecipePerson)        ' 처방자
    txtRecipePerson.Tag = CheckNull(rs!RecipePersonID)
    txtOrderID = rs!OrderID                       ' 처방전 관리번호
    txtRecipeNO = CheckNull(rs!RecipeNO)
    txtRecipeSeq = rs!RecipeSeq
    txtModifySeq = rs!ModifySeq
    txtWght = rs!UnitWght
    txtRecipeRemark = CheckNull(rs!Remark)
    
    txtPerson = CheckNull(rs!RPRatePerson)        ' 평량 처방자
    txtPerson.Tag = CheckNull(rs!RPRatePersonID)
    txtRPCalcRemark = CheckNull(rs!RPRateRemark)
    
    rs.Close
    
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows

    ' 실험실에서 처방된 투입비율을 표기
    'Set rs = oRecipe.GetRecipeByColor(txtOrderID, txtColor.Tag, txtRecipeSeq, txtModifySeq)
    Set rs = oRecipe.GetRecipeSub(txtOrderID, txtColor.Tag, 1, txtRecipeSeq, 0, "0", txtModifySeq)

    Do Until rs.EOF
        If Left(rs!DyeAuxID, 1) = "1" Then
            With grdDyeAux(0)
                .AddItem CStr(.Rows) & vbTab & Trim(rs!DyeAux) & vbTab & rs!DyeAuxRate
            End With
        Else
            With grdDyeAux(1)
                .AddItem CStr(.Rows) & vbTab & Trim(rs!DyeAux) & vbTab & rs!DyeAuxRate
            End With
        End If

        rs.MoveNext
    Loop

    rs.Close

    
    
    '**************************************************************************
    '*
    '* 평량지시 세부내역 확인 (2003-12-02)
    '*
    '*  - 본작업 평량지시 내역
    '* Author : 최승백
    '**************************************************************************
    
    Set rs = oRecipe.GetMatchSub(nDyeID, 1, "1")

    
    For i = 1 To rs.RecordCount

        With grdDyeAux(0)
            .TextMatrix(i, 3) = rs!DyeAuxRate
            .TextMatrix(i, 4) = rs!DyeAuxQty
            .TextMatrix(i, 5) = rs!DyeAuxID
            
            .Col = 3
            .Row = i
            .CellBackColor = IIf(CSng(.TextMatrix(i, 2)) <> CSng(.TextMatrix(i, 3)), vbRed, vbWhite)

        End With
            
        rs.MoveNext
    Next i

    Set rs = oRecipe.GetMatchSub(nDyeID, 1, "0")

    ' 본작업시 처방된 투입 비율및 투입수량 출력
    For i = 1 To rs.RecordCount
    
        With grdDyeAux(1)
            .TextMatrix(i, 3) = rs!DyeAuxRate
            .TextMatrix(i, 4) = rs!DyeAuxQty
            .TextMatrix(i, 5) = rs!DyeAuxID
            
            .Col = 3
            .Row = i
            .CellBackColor = IIf(CSng(.TextMatrix(i, 2)) <> CSng(.TextMatrix(i, 3)), vbRed, vbWhite)

        End With

        rs.MoveNext
    Next i
    '''''
    ' 본작업 평량 지시내역 출력 끝

    
    '**************************************************************************
    '*  - 재작업 평량지시 내역
    '*
    '* Author : 최승백
    '**************************************************************************
    If nDyeSeq > 1 Then
        For nSeq = 2 To IIf(m_bModify = True, nDyeSeq, nDyeSeq - 1)
            ' 평량 새로 작성시에는 바로 이전 단계의 재작업 내역 출력
            ' 변경시에는 재작업 평량 지시내역 모두 출력
            Set rs = oRecipe.GetMatchSub(nDyeID, nSeq, "1")
             
            ' 염료 - 추가작업 투입량
            For i = 1 To rs.RecordCount
                With grdDye(nSeq - 2)
                
                    .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & rs!DyeAuxRate & vbTab & rs!DyeAuxQty & vbTab & rs!DyeAuxID
                                        
                End With
                    
                rs.MoveNext
            Next i
        
            Set rs = oRecipe.GetMatchSub(nDyeID, nSeq, "0")
        
            ' 조제 - 추가작업 투입량
            For i = 1 To rs.RecordCount
                With grdAux(nSeq - 2)
                    .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & rs!DyeAuxRate & vbTab & rs!DyeAuxQty & vbTab & rs!DyeAuxID

                End With
        
                rs.MoveNext
            Next i
        Next nSeq
    
    End If
    
    ' 새로 입력할 추가 작업 평량지시 위해 처방전 불러옴
    If m_bModify = False Then
        Call GetRecipeData(1)
    End If
    ''
    ' 추가작업 평량지시내역 출력
    

    Set rs = Nothing
    Set oRecipe = Nothing
    Screen.MousePointer = vbArrow

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbArrow
    Set rs = Nothing
    Set oRecipe = Nothing
    Call ErrorBox(Err.Number, "frmRecipeCalc.ShowMatchData", Err.Description)
End Sub


Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            Call ReturnCode(LG_PERSON, , False, txtPerson)
        ElseIf Index = 1 Then
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    ElseIf KeyAscii = vbKeyReturn And Index >= 3 Then
        Call NextFocus
    End If
End Sub

Private Function DeleteData() As Boolean
    Dim oRecipe  As PlusLib2.CRecipe
    Dim nDyeSchID&, nDyeSeq%
    
    If grdData.Rows = grdData.FixedRows Then Exit Function
   
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName
    
    With grdRecipeCalc
        nDyeSchID = .TextMatrix(.Row, 24)
        nDyeSeq = .TextMatrix(.Row, 25)
    End With
    
    DeleteData = oRecipe.DeleteMatchData(nDyeSchID, nDyeSeq)
    
    Set oRecipe = Nothing
    
    Exit Function

ErrHandler:
    Set oRecipe = Nothing
    DeleteData = False

    Call ErrorBox(Err.Number, "frmRecipeCalc.DeleteData", Err.Description)
End Function



Private Sub cmdSave_Click()
    On Error GoTo ErrHandler

    If CheckData Then
        If m_nDyeSeq = 0 Then
            If Not SaveRapidSch(CheckNum(txtRoll), CheckNum(txtINQty)) Then Exit Sub
        End If
        Call SaveMatchData
    
    End If
    
    Exit Sub

ErrHandler:

    Call ErrorBox(Err.Number, "frmRecipeCalc.cmdSave_Click", Err.Description)

End Sub


Private Function CheckData() As Boolean

    If grdCard.Rows = grdCard.FixedRows Then
        MessageBox "공정카드가 설정되지 않았습니다"
        CheckData = False
        Exit Function
    End If

    If Len(txtPerson.Tag) = 0 Then
        MessageBox "작업자를 입력하십시오"
        CheckData = False
        txtPerson.SetFocus
        Exit Function
    End If

    CheckData = True

End Function


Private Function SaveMatchData() As Boolean
    Dim oRecipe As PlusLib2.CRecipe
    Dim tData As PlusLib2.TMatch
    Dim tDataSub() As PlusLib2.TMatchSub
    Dim i%, j%, nWorkClss%, nCnt%
    Dim nRateCol%

    On Error GoTo ErrHandler

    ' 평량지시 기본 내역
    With tData
        .DyeID = m_nDyeID
        .DyeSeq = m_nDyeSeq
        .RecipeOrderID = txtOrderID
        .RecipeOrderSeq = txtColor.Tag
        .RecipeSeq = txtRecipeSeq
        .RecipeModifySeq = txtModifySeq
        .PersonID = txtPerson.Tag
        .Remark = txtRPCalcRemark
    End With

    nCnt = 0

    If m_nDyeSeq = 1 Then
        ' 본작업 평량지시 염조제 내역
        For i = 0 To 1
            With grdDyeAux(i)
                For j = 1 To .Rows - .FixedRows
                    ReDim Preserve tDataSub(nCnt)
                    tDataSub(nCnt).DyeID = m_nDyeID
                    tDataSub(nCnt).DyeSeq = m_nDyeSeq
                    tDataSub(nCnt).DyeAuxSeq = nCnt + 1
                    tDataSub(nCnt).DyeAuxID = .TextMatrix(j, 5)     ' 염조제 I
                    tDataSub(nCnt).DyeAuxRate = CDbl(.TextMatrix(j, 3))   ' 실제 투입비율
                    tDataSub(nCnt).DyeAuxQty = CSng(.TextMatrix(j, 4))    ' 투입량
                    tDataSub(nCnt).DyeAuxRQty = 0
                    
                    nCnt = nCnt + 1
                Next j
            End With
        Next i
        ' 본작업 평량지시 염조제 내역 작성 끝.
        
    Else
        ' 추가작업 평량지시
        ' 염료
        With grdDye(m_nDyeSeq - 2)
            For j = 1 To .Rows - .FixedRows
                ReDim Preserve tDataSub(nCnt)
                tDataSub(nCnt).DyeID = m_nDyeID
                tDataSub(nCnt).DyeSeq = m_nDyeSeq
                tDataSub(nCnt).DyeAuxSeq = nCnt + 1                             '순위
                tDataSub(nCnt).DyeAuxID = .TextMatrix(j, 4)                 ' 염료 ID
                tDataSub(nCnt).DyeAuxRate = CDbl(CheckNum(.TextMatrix(j, 2)))     ' 처방 비율
                tDataSub(nCnt).DyeAuxQty = CheckNum(.TextMatrix(j, 3))      ' 투입량
                
                If tDataSub(nCnt).DyeAuxQty = 0 Then
                    MessageBox "염료 투입량이 입력되지 않았습니다"
                    
                    Exit Function
                End If
                
                tDataSub(nCnt).DyeAuxRQty = 0
            
                nCnt = nCnt + 1
            Next j
        
        End With
        
        '조제
        With grdAux(m_nDyeSeq - 2)
            For j = 1 To .Rows - .FixedRows
                ReDim Preserve tDataSub(nCnt)
                tDataSub(nCnt).DyeID = m_nDyeID
                tDataSub(nCnt).DyeSeq = m_nDyeSeq
                tDataSub(nCnt).DyeAuxSeq = nCnt + 1                     ' 순위
                tDataSub(nCnt).DyeAuxID = .TextMatrix(j, 4)             ' 조제 ID
                tDataSub(nCnt).DyeAuxRate = CDbl(CheckNum(.TextMatrix(j, 2))) ' 처방 비율
                tDataSub(nCnt).DyeAuxQty = CheckNum(.TextMatrix(j, 3))  ' 처방 투입량
                
                If tDataSub(nCnt).DyeAuxQty = 0 Then
                    MessageBox "조제 투입량이 입력되지 않았습니다"
                    
                    Exit Function
                End If
                
                tDataSub(nCnt).DyeAuxRQty = 0
            
                nCnt = nCnt + 1
            Next j
        
        End With
    
    End If

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    oRecipe.UserName = g_sUserName

    If m_bModify = False Then
        SaveMatchData = oRecipe.AddNewMatchData(tData, tDataSub)
    Else
        ' 처방전 수정
        SaveMatchData = oRecipe.UpdateMatchData(tData, tDataSub)
    End If
    
    Set oRecipe = Nothing
    
    MessageBox "평량지시 내역이 저장되었습니다"

    m_bSaved = True
    cmdPrint.Enabled = True

    Exit Function

ErrHandler:
    Set oRecipe = Nothing

    Call ErrorBox(Err.Number, "frmRecipeCalc.SaveMatchData", Err.Description)

End Function


Private Sub FillGridData()
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    m_bLoading1 = True
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    
    Set rs = oRecipe.GetCardList(IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), txtSearch(4), txtSearch(5), _
                                 IIf(chkSearch(5) = vbChecked, 1, 0), Format(Left(cboProcess.ItemData(cboProcess.ListIndex), 2), "00"), sbTab.Tab)
    Set oRecipe = Nothing
        
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & False & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & _
                    rs!Color & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                    rs!UseClss & vbTab & CheckNull(rs!AfterProc) & vbTab & rs!OrderSeq & vbTab & rs!ReWorkClss & vbTab & _
                    rs!EmerClss & vbTab & rs!PatternID & vbTab & rs!ThreadName & vbTab & rs!StuffCustom & vbTab & _
                    rs!StuffWidth & vbTab & rs!StuffDensity & vbTab & rs!BatJaNO & vbTab & rs!DyeSchID & vbTab & _
                    rs!DyeSeq & vbTab & rs!UnitClss
            
            If rs!UseClss = "보류" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbRed
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            ElseIf rs!UseClss = "작업" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbBlue
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            End If
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
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
            
            Call ClearData
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bLoading1 = False
    Exit Sub

ErrHandler:
    Set oRecipe = Nothing
    Set rs = Nothing
    pnlProgress.Visible = False
    m_bLoading1 = False
    Call ErrorBox(Err.Number, "frmRecipeCalc.FillGridData", Err.Description)
End Sub

Private Sub FillGridRecipeCalc()
    Dim oRecipe As PlusLib2.CRecipe
    Dim rs As ADODB.Recordset
    Dim i%, nTop%, nDyeSchID%
    
    On Error GoTo ErrHandler
    
    m_bLoading1 = True
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon
    
    Set rs = oRecipe.GetCardList(IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), txtSearch(4), txtSearch(5), _
                                 IIf(chkSearch(5) = vbChecked, 1, 0), Format(Left(cboProcess.ItemData(cboProcess.ListIndex), 2), "00"), stTab.Tab)
    Set oRecipe = Nothing
        
    With grdRecipeCalc
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            If nDyeSchID <> rs!DyeSchID Then
                .AddItem "" & vbTab & "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                        rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & _
                        rs!Color & vbTab & rs!wiRoll & vbTab & rs!wiQty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                        rs!UseClss & vbTab & CheckNull(rs!AfterProc) & vbTab & rs!OrderSeq & vbTab & rs!ReWorkClss & vbTab & _
                        rs!EmerClss & vbTab & rs!PatternID & vbTab & rs!ThreadName & vbTab & rs!StuffCustom & vbTab & _
                        rs!StuffWidth & vbTab & rs!StuffDensity & vbTab & rs!BatJaNO & vbTab & rs!DyeSchID & vbTab & _
                        rs!DyeSeq & vbTab & rs!UnitClss & vbTab & rs!MaxCardSeq & vbTab & rs!wiMachID & vbTab & _
                        rs!WorkClss & vbTab & rs!RapidClss & vbTab & rs!DyePatternID

                        Call DoFlexGridGroup(grdRecipeCalc, .Rows - 1, 1)
                        Call GridCollapse(grdRecipeCalc, nTop)
                        nTop = .Rows - 1
            End If
            
            If rs!MaxCardSeq > 1 Then
                .AddItem "" & vbTab & "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                        rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & _
                        rs!Color & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                        rs!UseClss & vbTab & CheckNull(rs!AfterProc) & vbTab & rs!OrderSeq & vbTab & rs!ReWorkClss & vbTab & _
                        rs!EmerClss & vbTab & rs!PatternID & vbTab & rs!ThreadName & vbTab & rs!StuffCustom & vbTab & _
                        rs!StuffWidth & vbTab & rs!StuffDensity & vbTab & rs!BatJaNO & vbTab & rs!DyeSchID & vbTab & _
                        rs!DyeSeq & vbTab & rs!UnitClss & vbTab & rs!MaxCardSeq & vbTab & rs!wiMachID & vbTab & _
                        rs!WorkClss & vbTab & rs!RapidClss & vbTab & rs!DyePatternID

            End If
            
            If rs!UseClss = "보류" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbRed
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            ElseIf rs!UseClss = "작업" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbBlue
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            End If
            nDyeSchID = rs!DyeSchID
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
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
                        
            Call ShowData
        Else
            .HighLight = flexHighlightNever
            
            Call ClearData
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bLoading1 = False
    Exit Sub

ErrHandler:
    Set oRecipe = Nothing
    Set rs = Nothing
    pnlProgress.Visible = False
    m_bLoading1 = False
    Call ErrorBox(Err.Number, "frmRecipeCalc.FillRecipeCalc", Err.Description)
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
                rs.MoveNext
            Next iCount
            rs.Close
            Set rs = Nothing
        End If
    End If
End Sub

Private Sub lstArray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 2:
            lstArray(4).Selected(0) = True
        Case 4:
            lstArray(2).ListIndex = -1
    End Select

End Sub

Private Sub AddLstBox()
    Dim oRapid As PlusLib2.CRapid
    Dim oPerson  As PlusLib2.CPerson
    Dim rs As Recordset
    Dim iCount%, i%, iSeq%
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    txtRemark1 = ""
    For i = 0 To lstArray.Count - 1
        lstArray(i).Clear
    Next i
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
                
    Set rs = oRapid.GetMachineNo("4301")
    For iCount = 1 To rs.RecordCount
        lstArray(0).AddItem Format(rs!MachineNO, "00") & " 호기" & "       " & rs!WaterRate
        
        rs.MoveNext
    Next iCount
    rs.Close
    Set rs = Nothing
    
    Set oRapid = Nothing
    
' 진호염직의 염색구분 목록
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
    lstArray(2).ListIndex = 0
' 진호염직의 작업구분
    lstArray(4).AddItem "염색"
    lstArray(4).AddItem "BOX 탈색"
    lstArray(4).AddItem "BOX R/C"
    lstArray(4).AddItem "도포 Washing"
    lstArray(4).AddItem "Soaping"
    lstArray(4).AddItem "기계수리"
    lstArray(4).ListIndex = 0
        
    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    oPerson.UserName = g_sUserName
    Set rs = oPerson.GetWorkerList("05")     '염색 부서
    For iCount = 1 To rs.RecordCount
        lstArray(3).AddItem rs!Name & "             " & Format(rs!PersonID, "00000000")
        rs.MoveNext
    Next iCount
    lstArray(3).ListIndex = 0
    rs.Close
    Set rs = Nothing
        
    Screen.MousePointer = vbDefault


    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    Set oPerson = Nothing
    
    Call ErrorBox(Err.Number, "frmRecipeCalc.AddLstBox", Err.Description)
End Sub

Private Function CheckRapidData() As Boolean
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
    If lstArray(4).SelCount = 0 Then
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(4).ListIndex > 0 Then
        If lstArray(2).SelCount > 0 Then
            MsgBox "염색구분이 선택되면 안됩니다", vbCritical, "작성 오류"
            Exit Function
        End If
    ElseIf lstArray(4).ListIndex = 0 Then
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
            
    If grdCardList.TextMatrix(1, 11) = "작업" Then
        MsgBox "염색작업중인 카드입니다." & vbCrLf & "작업진행중인 카드는 추가작업만 할수 있습니다.", vbCritical, "작성 오류"
        Exit Function
    End If
            
    CheckRapidData = True
End Function

Private Sub ShowDyeList()
    Dim i%, nRecipeCnt%
    
    txtRemark = txtRemark1
    txtMachine = Left(lstArray(0).Text, 2)
    txtPattern = Left(lstArray(1).Text, 2)
    txtRoll = grdCardList.TextMatrix(grdCardList.Rows - 1, 12)
    txtINQty = grdCardList.TextMatrix(grdCardList.Rows - 1, 13)
    txtColor = grdCardList.TextMatrix(grdCardList.FixedRows, 7)
    txtColor.Tag = grdCardList.TextMatrix(grdCardList.FixedRows, 21)
    txtWorkClss = lstArray(2).Text
    txtCustom = grdCardList.TextMatrix(grdCardList.FixedRows, 5)
    txtArticle = grdCardList.TextMatrix(grdCardList.FixedRows, 6)
    txtWaterRate = Trim(Right(lstArray(0).Text, 5))

    m_bModify = False
      
    With grdCardList
        grdCard.Rows = grdCard.FixedRows
        For i = 1 To .Rows - 2
            grdCard.AddItem grdCard.Rows
            grdCard.TextMatrix(grdCard.Rows - 1, 1) = .TextMatrix(i, 9)
            grdCard.TextMatrix(grdCard.Rows - 1, 2) = .TextMatrix(i, 10)
            grdCard.TextMatrix(grdCard.Rows - 1, 3) = .TextMatrix(i, 8)
            grdCard.TextMatrix(grdCard.Rows - 1, 5) = .TextMatrix(i, 12)
            grdCard.TextMatrix(grdCard.Rows - 1, 6) = .TextMatrix(i, 13)
            grdCard.TextMatrix(grdCard.Rows - 1, 7) = IIf(.TextMatrix(i, 25) = 0, "Y", "M") '단위
            grdCard.TextMatrix(grdCard.Rows - 1, 8) = .TextMatrix(i, 25) '단위
            grdCard.TextMatrix(grdCard.Rows - 1, 9) = .TextMatrix(i, 7)
        Next i
    End With
    
    nRecipeCnt = GetRecipeCount  ' 처방전 갯수 파악.
    
    If nRecipeCnt = 0 Then
        MessageBox "처방전이 내려지지 않았습니다. 확인해 주십시오"
    
        Exit Sub
    End If
    
    If nRecipeCnt > 1 Then
        Call GetRecipeDataAll
    Else
        Call GetRecipeData(0)
    End If
End Sub

Private Function SaveRapidSch(TotRoll As Long, TotQty As Long) As Boolean
    Dim oRapid As PlusLib2.CRapid
    Dim tCardList() As PlusLib2.tRapidCard
    Dim i%, iCount%, iCntChk%, iCol%, iRow%, iSeq%
    Dim nDyeSchID&, nDyeSeq%
    
    Screen.MousePointer = vbHourglass
    SaveRapidSch = False

    On Error GoTo ErrHandler

    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

    With grdCardList
        ReDim tCardList(.Rows - 3)
        iCount = 0
        For i = .FixedRows To .Rows - 2
            tCardList(iCount).sCardID = Trim(.TextMatrix(i, 17))
            tCardList(iCount).sSplitID = IIf(Trim(.TextMatrix(i, 18)) = "", " ", Trim(.TextMatrix(i, 18)))
            If lstArray(2).Text = "추가" Then
                tCardList(iCount).lDyeSchID = CLng(.TextMatrix(i, 23))
            Else
                tCardList(iCount).lDyeSchID = 0
            End If
            iCount = iCount + 1
        Next i
    End With
        
    If Not oRapid.AddNewwiRapidItem(tCardList(), CLng(tCardList(0).lDyeSchID), "4301", Left(lstArray(0).Text, 2), _
        0, lstArray(4).Text, lstArray(2).Text, Format(CInt("0" & Left(lstArray(1).Text, 2)), "000"), 0, TotRoll, _
        TotQty, " ", Right(lstArray(3).Text, 8), CheckNull(txtRemark), nDyeSchID, nDyeSeq) Then
        Set oRapid = Nothing
        SaveRapidSch = False
        Exit Function
    End If
    m_nDyeID = nDyeSchID
    m_nDyeSeq = nDyeSeq
    
    SaveRapidSch = True
    
    Set oRapid = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    SaveRapidSch = False

    Set oRapid = Nothing
    Call ErrorBox(Err.Number, "frmRecipeCalc.SaveRapidSch", Err.Description)
End Function

Private Sub ShowData()
    Dim i%, nSeq%, nDyeSchID&, nDyeSeq%
    Dim nRoll%, nQty%
    Dim sMachID$, sWorkClss$, sRapidClss$, sPatternID$
    
    With grdRecipeCalc
        .Redraw = flexRDNone
        grdCardList.Rows = grdCardList.FixedRows
        If .Rows = .FixedRows Then Exit Sub
        nDyeSchID = .TextMatrix(.Row, 24)
        nDyeSeq = .TextMatrix(.Row, 25)
        For i = .FixedRows To .Rows - 1
            If nDyeSchID = .TextMatrix(i, 24) And nDyeSeq = .TextMatrix(.Row, 25) And (.IsSubtotal(i) = False Or .TextMatrix(i, 27) = 1) Then
                grdCardList.AddItem ""
                grdCardList.TextMatrix(grdCardList.Rows - 1, 3) = .TextMatrix(i, 23)    '밧자번호
                grdCardList.TextMatrix(grdCardList.Rows - 1, 4) = nSeq + 1
                grdCardList.TextMatrix(grdCardList.Rows - 1, 5) = .TextMatrix(i, 2)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 6) = .TextMatrix(i, 3)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 7) = .TextMatrix(i, 8)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 8) = .TextMatrix(i, 4)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 9) = .TextMatrix(i, 6)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 10) = .TextMatrix(i, 7)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 11) = .TextMatrix(i, 13)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 12) = .TextMatrix(i, 9)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 13) = .TextMatrix(i, 10)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 17) = MakeCardID(.TextMatrix(i, 6), OM_REDUCE)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 18) = .TextMatrix(i, 7)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 20) = MakeOrderID(.TextMatrix(i, 4), OM_REDUCE)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 21) = .TextMatrix(i, 15)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 23) = .TextMatrix(i, 24)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 24) = .TextMatrix(i, 25)
                grdCardList.TextMatrix(grdCardList.Rows - 1, 25) = .TextMatrix(i, 26)
                
                nRoll = nRoll + CheckNum(.TextMatrix(i, 9))
                nQty = nQty + CheckNum(.TextMatrix(i, 10))
                nSeq = nSeq + 1
                
                sMachID = .TextMatrix(i, 28)
                sWorkClss = .TextMatrix(i, 29)
                sRapidClss = .TextMatrix(i, 30)
                sPatternID = .TextMatrix(i, 31)
                
            End If
            
        Next i
        .Redraw = flexRDDirect
        
    End With
    
    grdCardList.Rows = grdCardList.Rows + 1
    grdCardList.RowHeight(grdCardList.Rows - 1) = 300
    grdCardList.Cell(flexcpText, grdCardList.Rows - 1, 0, grdCardList.Rows - 1, 11) = "선택되어진 카드 총 합계"
    grdCardList.Cell(flexcpFontBold, grdCardList.Rows - 1, 0, grdCardList.Rows - 1, grdCardList.Cols - 1) = True
    grdCardList.TextMatrix(grdCardList.Rows - 1, 12) = Format(nRoll, "#,##0")
    grdCardList.TextMatrix(grdCardList.Rows - 1, 13) = Format(nQty, "#,###,##0")
    grdCardList.MergeCells = flexMergeRestrictRows
    grdCardList.MergeRow(grdCardList.Rows - 1) = True
    grdCardList.Row = grdCardList.FixedRows
    
    '염색호기 0 패턴 1 작업구분4 염색구분 2
        
    For i = 0 To lstArray(0).ListCount - 1
        If Left(lstArray(0).List(i), 2) = Format(sMachID, "00") Then
            lstArray(0).Selected(i) = True
            Exit For
        End If
    Next i
        
    For i = 0 To lstArray(1).ListCount - 1
        If Left(lstArray(1).List(i), 2) = Format(sPatternID, "00") Then
            lstArray(1).Selected(i) = True
            Exit For
        End If
    Next i
        
    ' 작업구분
    For i = 0 To lstArray(4).ListCount - 1
        If lstArray(4).List(i) = sWorkClss Then
            lstArray(4).Selected(i) = True
            Exit For
        End If
    Next i
    ' 염색구분
    For i = 0 To lstArray(2).ListCount - 1
        If lstArray(2).List(i) = sRapidClss Then
            lstArray(2).Selected(i) = True
            Exit For
        End If
    Next i
End Sub

