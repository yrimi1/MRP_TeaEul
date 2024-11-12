VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInspectCode 
   Caption         =   "검사관련 코드관리"
   ClientHeight    =   6795
   ClientLeft      =   2595
   ClientTop       =   1155
   ClientWidth     =   10290
   Icon            =   "frmInspectCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10290
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   5805
      TabIndex        =   19
      Top             =   5100
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   900
      _Version        =   196609
      BackColor       =   65535
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   8550
      TabIndex        =   18
      Top             =   5985
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab tabForm 
      Height          =   5340
      Left            =   30
      TabIndex        =   16
      Top             =   960
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   9419
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   679
      TabCaption(0)   =   "불량 관리 "
      TabPicture(0)   =   "frmInspectCode.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pnlEdit(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdData(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "검사기준 관리"
      TabPicture(1)   =   "frmInspectCode.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pnlEdit(1)"
      Tab(1).Control(1)=   "grdData(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "등급관리"
      TabPicture(2)   =   "frmInspectCode.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pnlEdit(2)"
      Tab(2).Control(1)=   "grdData(2)"
      Tab(2).ControlCount=   2
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4815
         Index           =   0
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   5625
         _cx             =   9922
         _cy             =   8493
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
      Begin Threed.SSPanel pnlEdit 
         Height          =   3315
         Index           =   0
         Left            =   5730
         TabIndex        =   26
         Top             =   15
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   5847
         _Version        =   196609
         Enabled         =   0   'False
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   1
            Left            =   1560
            TabIndex        =   45
            Top             =   2970
            Width           =   2775
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   3
            Left            =   1560
            TabIndex        =   2
            Top             =   840
            Width           =   2775
            _ExtentX        =   4895
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
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   0
            Left            =   1560
            TabIndex        =   7
            Top             =   2610
            Width           =   2775
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   2
            Left            =   1545
            TabIndex        =   1
            Top             =   450
            Width           =   2760
            _ExtentX        =   4868
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
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   1
            Left            =   1560
            TabIndex        =   12
            Top             =   90
            Width           =   1110
            _ExtentX        =   1958
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
            MaxLength       =   3
            BackColor       =   12648384
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   27
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "코     드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   90
            TabIndex        =   28
            Top             =   450
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "단말기 Display1"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   90
            TabIndex        =   29
            Top             =   810
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "단말기 Display2"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   90
            TabIndex        =   30
            Top             =   1170
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "단말기 Display3"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   90
            TabIndex        =   31
            Top             =   2250
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "Tag Name"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   4
            Left            =   1545
            TabIndex        =   3
            Top             =   1170
            Width           =   2775
            _ExtentX        =   4895
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
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   5
            Left            =   1545
            TabIndex        =   4
            Top             =   1530
            Width           =   2775
            _ExtentX        =   4895
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
            MaxLength       =   25
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   33
            Top             =   1530
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "불량명 (한글)"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   6
            Left            =   90
            TabIndex        =   34
            Top             =   1890
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "불량명 (영문)"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   6
            Left            =   1545
            TabIndex        =   5
            Top             =   1890
            Width           =   2775
            _ExtentX        =   4895
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
            MaxLength       =   25
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   7
            Left            =   1545
            TabIndex        =   6
            Top             =   2250
            Width           =   2775
            _ExtentX        =   4895
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
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   8
            Left            =   90
            TabIndex        =   35
            Top             =   2610
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "불량 종류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   13
            Left            =   90
            TabIndex        =   44
            Top             =   2970
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "세부 불량 종류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4815
         Index           =   1
         Left            =   -74940
         TabIndex        =   32
         Top             =   30
         Width           =   5580
         _cx             =   9842
         _cy             =   8493
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
      Begin Threed.SSPanel pnlEdit 
         Height          =   855
         Index           =   1
         Left            =   -69270
         TabIndex        =   11
         Top             =   45
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1508
         _Version        =   196609
         Enabled         =   0   'False
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   10
            Left            =   1545
            TabIndex        =   8
            Top             =   450
            Width           =   2760
            _ExtentX        =   4868
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
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   9
            Left            =   1560
            TabIndex        =   15
            Top             =   90
            Width           =   1110
            _ExtentX        =   1958
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
            MaxLength       =   2
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   36
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "코       드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   10
            Left            =   90
            TabIndex        =   37
            Top             =   450
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "검사 기준"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4815
         Index           =   2
         Left            =   -74940
         TabIndex        =   38
         Top             =   75
         Width           =   5580
         _cx             =   9842
         _cy             =   8493
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
      Begin Threed.SSPanel pnlEdit 
         Height          =   840
         Index           =   2
         Left            =   -69270
         TabIndex        =   39
         Top             =   90
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1482
         _Version        =   196609
         Enabled         =   0   'False
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   12
            Left            =   1545
            TabIndex        =   40
            Top             =   450
            Width           =   2760
            _ExtentX        =   4868
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
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Index           =   11
            Left            =   1545
            TabIndex        =   41
            Top             =   90
            Width           =   1110
            _ExtentX        =   1958
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
            MaxLength       =   1
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   11
            Left            =   90
            TabIndex        =   42
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "코       드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   12
            Left            =   90
            TabIndex        =   43
            Top             =   450
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "등       급"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   915
      Left            =   5745
      TabIndex        =   17
      Top             =   30
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   1350
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   9
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   2940
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   13
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   3735
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   14
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   2145
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   555
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   10
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   915
      Left            =   30
      TabIndex        =   20
      Top             =   30
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   21
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코드명 검색"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   2160
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   450
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         _Version        =   196609
         MousePointer    =   99
         CaptionStyle    =   1
         PictureAnimationEnabled=   0   'False
         Alignment       =   6
         PictureAlignment=   0
         BevelWidth      =   1
         ShapeSize       =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "검색건수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   24
      Top             =   6450
      Width           =   945
   End
End
Attribute VB_Name = "frmInspectCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH As Integer = 3140
Private Const LIMIT_ROW = 16

Private m_sOperate     As String * 1
Private m_bSortForward As Boolean
Private m_bloading  As Boolean

Private Sub Form_Load()
    Me.Move 0, 0, 10410, 7140

    Call SetOperate(Me)

    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)

    Call InitGrid
    'Call MakeCodeCombo(cboName(0), CD_PROCESS)
    With cboName(0)
        .AddItem "가공불량"
        .AddItem "제직불량"
    End With
    With cboName(1)
        .AddItem "염색불량"
        .AddItem "가공불량"
        .AddItem "제직불량"
    End With

    Call FillGrid
End Sub

Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
    Dim i%, iCount%, iNowRow%

    On Error GoTo ErrHandler

    If Len(Trim(txtSearch)) > 0 Then
        With grdData(tabForm.Tab)
            .Redraw = flexRDNone

            For i = .FixedRows To .Rows - .FixedRows
                If InStr(UCase(.TextArray(i * .Cols + 2)), UCase(txtSearch)) = 0 Then
                    .RowHidden(i) = True
                    iCount = iCount + 1
                Else
                    .RowHidden(i) = False
                    iNowRow = i
                End If
            Next i

            If iNowRow > .FixedRows Then
                .Row = iNowRow
                .Col = .FixedCols
                .ColSel = .Cols - 1
            End If

            .Redraw = flexRDDirect
        End With
    Else
        Call cmdAll_Click
    End If

    If iCount > 0 Then
        cmdAll.Visible = True
    Else
        cmdAll.Visible = False
    End If

    Call ChangeScroll

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "txtSearch.Change", Err.Description)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then grdData(tabForm.Tab).SetFocus
End Sub

Private Sub cmdAll_Click()
    Dim i%

    With grdData(tabForm.Tab)
        .Redraw = flexRDNone

        For i = .FixedRows To .Rows - .FixedRows
            .RowHidden(i) = False
        Next i

        .Redraw = flexRDDirect
    End With

    txtSearch.Text = ""
    cmdAll.Visible = False
End Sub

Private Sub grdData_DblClick(Index As Integer)
    With grdData(tabForm.Tab)
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub

        Call cmdOperate_Click(ID_UPDATE)
    End With
End Sub

Private Sub grdData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOperate_Click(ID_UPDATE)
End Sub

Private Sub grdData_RowColChange(Index As Integer)
    Call ShowData
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim bResult As Boolean

    If grdData(1).Rows > 9 Then
        Call MessageBox("등급 종류는 9개를 넘을 수 없습니다.")
        Exit Sub
    End If

    On Error GoTo ErrHandler

    Select Case Index
    Case ID_ADDNEW
        m_sOperate = ID_ADDNEW
        Call ChangeMode(Me, False)
        Call ClearData
        pnlMsg.Caption = LoadResString(302)

        pnlEdit(tabForm.Tab).Enabled = True

        Select Case tabForm.Tab
        Case 0 '불량관리
            tabForm.TabEnabled(1) = False
            tabForm.TabEnabled(2) = False
            txtName(2).SetFocus
        Case 1 '검사기준관리
            tabForm.TabEnabled(0) = False
            tabForm.TabEnabled(2) = False
            txtName(10).SetFocus
        Case 2 '등급관리
            tabForm.TabEnabled(0) = False
            tabForm.TabEnabled(1) = False
            txtName(12).SetFocus
        End Select
    Case ID_UPDATE '[2] 수정
        m_sOperate = ID_UPDATE
        Call ChangeMode(Me, False)

        pnlMsg.Caption = LoadResString(303)

        pnlEdit(tabForm.Tab).Enabled = True
        Select Case tabForm.Tab
        Case 0 '불량 관리
            tabForm.TabEnabled(1) = False
            tabForm.TabEnabled(2) = False
            txtName(1).Locked = True
            txtName(2).SetFocus
        Case 1 '검사기준
            tabForm.TabEnabled(0) = False
            tabForm.TabEnabled(2) = False
            
            txtName(9).Locked = True
            txtName(10).SetFocus
        
        Case 2 '등급 관리
            tabForm.TabEnabled(0) = False
            tabForm.TabEnabled(1) = False
            
            txtName(11).Locked = True
            txtName(12).SetFocus
        End Select
    Case ID_DELETE '[3] 삭제
        If grdData(tabForm.Tab).Rows = grdData(tabForm.Tab).FixedRows Then Exit Sub

        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
            m_sOperate = ID_DELETE

            If SaveData() Then Call FillGrid
        End If
    Case ID_SAVE  '[4] 저장
        If SaveData() Then
            Call FillGrid
            Call ChangeMode(Me, True)

            m_sOperate = ""
            tabForm.TabEnabled(0) = True
            tabForm.TabEnabled(1) = True
            tabForm.TabEnabled(2) = True

            pnlEdit(tabForm.Tab).Enabled = False

            Select Case tabForm.Tab
            Case 0 '불량 관리
                txtName(1).Locked = False
            Case 1 '검사기준 관리
                txtName(9).Locked = False
            Case 2 '등급관리
                txtName(11).Locked = False
            End Select
        End If
        grdData(tabForm.Tab).SetFocus
    Case ID_CANCEL
        m_sOperate = ""
        If grdData(tabForm.Tab).Rows > 1 Then
            Call ShowData
        Else
            Call ClearData
        End If
        Call ChangeMode(Me, True)
        
        tabForm.TabEnabled(0) = True
        tabForm.TabEnabled(1) = True
        tabForm.TabEnabled(2) = True
         
        pnlEdit(tabForm.Tab).Enabled = False

        Select Case tabForm.Tab
        Case 0 '불량 관리
            txtName(1).Locked = False
        Case 1 '검사기준 관리
            txtName(9).Locked = False
        Case 2 '등급 관리
            txtName(11).Locked = False
        End Select
        grdData(tabForm.Tab).SetFocus
    End Select

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 12 Or Index = 14 Or Index = 16 Then
        If KeyCode = vbKeyReturn Then cmdOperate(4).SetFocus
    End If
End Sub

Private Sub cboName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyReturn Then cboName(1).SetFocus
    Else
        If KeyCode = vbKeyReturn Then cmdOperate(3).SetFocus
    End If
End Sub

Private Sub tabform_Click(PreviousTab As Integer)
    Dim sMenuID As String

    pnlCaption(0).Caption = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "") & "검색"

    Call PlusMDI.RunForm(1610 + (10 * tabForm.Tab))

    Call FillGrid

    txtSearch.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub ClearData()
    Dim i%

    Select Case tabForm.Tab
        Case 0
            For i = 1 To 7
                txtName(i) = ""
            Next i
            cboName(0).ListIndex = 0
            cboName(1).ListIndex = 0
        Case 1
            For i = 9 To 10
                txtName(i) = ""
            Next i
        Case 2
            txtName(11) = ""
            txtName(12) = ""
    End Select
End Sub

Private Sub ShowData()
    Dim i As Integer

    If m_bloading = True Then Exit Sub
    
    With grdData(tabForm.Tab)
        Select Case tabForm.Tab
            Case 0
                For i = 1 To 7
                txtName(i) = .TextMatrix(.Row, i)
                Next i
                cboName(0).ListIndex = .TextMatrix(.Row, 8) - 1
                cboName(1).ListIndex = .TextMatrix(.Row, 9)
            Case 1  ' 검사기준
                For i = 9 To 10
                txtName(i) = .TextMatrix(.Row, i - 8)
                Next i
            Case 2
                txtName(11) = .TextMatrix(.Row, 1)
                txtName(12) = .TextMatrix(.Row, 2)
         End Select
    End With
End Sub

Private Sub InitGrid()
    With grdData(0)
        .Cols = 10
        Call SetVSFlexGrid(grdData(0))

        .Redraw = flexRDNone
        .Rows = 1

        .TextArray(0) = ""
        .TextArray(1) = "코드":                         .ColWidth(1) = 450:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "단말기" & vbCrLf & "DISPLAY1": .ColWidth(2) = 0:       .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "단말기" & vbCrLf & "DISPLAY2": .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "단말기" & vbCrLf & "DISPLAY3": .ColWidth(4) = 0:       .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "불량명" & vbCrLf & "(한글)":   .ColWidth(5) = 2100:    .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "불량명" & vbCrLf & "(영문)":   .ColWidth(6) = 2000:    .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "Tag" & vbCrLf & "Name":        .ColWidth(7) = 600:     .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "불량" & vbCrLf & "종류":       .ColWidth(8) = 0:       .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "불량" & vbCrLf & "종류":       .ColWidth(9) = 0:       .ColAlignment(9) = flexAlignLeftCenter

        .Redraw = flexRDDirect
    End With

    With grdData(1)
        .Cols = 3
        Call SetVSFlexGrid(grdData(1))

        .Redraw = flexRDNone
        .Rows = 1

        .TextArray(0) = ""
        .TextArray(1) = "코    드":         .ColWidth(1) = 600:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "검사기준":         .ColWidth(2) = 1000:    .ColAlignment(2) = flexAlignLeftCenter

        .Redraw = flexRDDirect
    End With
    
    With grdData(2)
        .Cols = 3
        Call SetVSFlexGrid(grdData(2))

        .Redraw = flexRDNone
        .Rows = 1

        .TextArray(0) = ""
        .TextArray(1) = "코  드":         .ColWidth(1) = 600:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "등  급":         .ColWidth(2) = 1000:    .ColAlignment(2) = flexAlignLeftCenter

        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub FillGrid()
    Dim oDefect  As PlusLib2.CDefect
    Dim oCode    As PlusLib2.CCode
    Dim oGrade   As PlusLib2.CGrade
    Dim rs      As ADODB.Recordset
    Dim lNowRow&

    On Error GoTo ErrHandler
    m_bloading = True
    
    Select Case tabForm.Tab
    Case 0 ' [0] 불량 관리
        Set oDefect = New PlusLib2.CDefect
        oDefect.Connection = g_adoCon

        Set rs = oDefect.GetDefect("%")
        Set oDefect = Nothing
    Case 1 ' [1] 검사기준 관리
        Set oCode = New PlusLib2.CCode
        oCode.CodeType = CD_BASIS
        oCode.Connection = g_adoCon
        
        Set rs = oCode.Getcode
        Set oCode = Nothing
    Case 2 ' 등급관리
        Set oGrade = New PlusLib2.CGrade
        oGrade.Connection = g_adoCon
        
        Set rs = oGrade.GetGrade
        Set oGrade = Nothing
    End Select

    With grdData(tabForm.Tab)
        .Redraw = flexRDNone

        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        Select Case tabForm.Tab
        Case 0
            Do Until rs.EOF
                .AddItem CStr(.Rows) & vbTab & rs!DefectID & vbTab & rs!Display1 & vbTab & rs!Display2 & vbTab & _
                    rs!Display3 & vbTab & rs!KDefect & vbTab & rs!EDefect & vbTab & rs!TagName & vbTab & rs!DefectClss & vbTab & rs!DefectClssSub

                rs.MoveNext
            Loop
        Case 1
            Do Until rs.EOF
                .AddItem CStr(.Rows) & vbTab & rs(0) & vbTab & rs(1)

                rs.MoveNext
            Loop
        Case 2
            Do Until rs.EOF
                .AddItem CStr(.Rows) & vbTab & rs(0) & vbTab & rs(1)

                rs.MoveNext
            Loop
        End Select
        
        rs.Close
        Set rs = Nothing

        lblCount.Caption = LoadResString(250) & CStr(grdData(tabForm.Tab).Rows - grdData(tabForm.Tab).FixedRows) & " 건"

        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
           .TopRow = lNowRow

           .Col = .FixedCols
           .ColSel = .Cols - 1

            
        Else
            .HighLight = flexHighlightNever

            Call ClearData
        End If

        .Redraw = flexRDDirect
    End With

    m_bloading = False
    Call ShowData
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oDefect = Nothing
    Set oGrade = Nothing
    m_bloading = False

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Function SaveData() As Boolean
    Dim oDefect   As PlusLib2.CDefect
    Dim oCode     As PlusLib2.CCode
    Dim oGrade    As PlusLib2.CGrade
    Dim tDef      As PlusLib2.TDefect
    Dim tCode     As PlusLib2.tCode
    Dim TGrade    As PlusLib2.TGrade

    On Error GoTo ErrHandler

    Select Case tabForm.Tab
    Case 0  '[0] 불량 관리
        If cboName(0).ListIndex = 0 Then
            If cboName(1).ListIndex = 2 Then
                MsgBox "세부 불량종류가 잘못되어 있습니다", vbCritical, "세부 불량종류 선택 오류"
                Exit Function
            End If
        Else
            If cboName(1).ListIndex < 2 Then
                MsgBox "세부 불량종류가 잘못되어 있습니다", vbCritical, "세부 불량종류 선택 오류"
                Exit Function
            End If
        End If
        
        Set oDefect = New PlusLib2.CDefect
        oDefect.Connection = g_adoCon
        oDefect.UserName = g_sUserName

        tDef.DefectID = IIf(m_sOperate = ID_ADDNEW, "1", txtName(1))
        tDef.Display1 = txtName(2)
        tDef.Display2 = txtName(3)
        tDef.Display3 = txtName(4)
        tDef.KDefect = txtName(5)
        tDef.EDefect = txtName(6)
        tDef.TagName = txtName(7)
        tDef.KindID = Format(cboName(0).ListIndex + 1, "0")
        tDef.KindIDSub = Format(cboName(1).ListIndex, "0")

        If m_sOperate = ID_ADDNEW Then
            SaveData = oDefect.AddNewDefect(tDef)
        ElseIf m_sOperate = ID_UPDATE Then
            SaveData = oDefect.UpdateDefect(tDef)
        ElseIf m_sOperate = ID_DELETE Then
            SaveData = oDefect.DeleteDefect(tDef.DefectID)
        End If

        Set oDefect = Nothing
    Case 1  '[1] 검사기준 관리
        Set oCode = New PlusLib2.CCode
        oCode.CodeType = CD_BASIS
        oCode.Connection = g_adoCon
        oCode.UserName = g_sUserName
        
        tCode.sCodeID = txtName(9)
        tCode.scode = txtName(10)
        
        If m_sOperate = ID_ADDNEW Then
            SaveData = oCode.AddNewCode(tCode)
        ElseIf m_sOperate = ID_UPDATE Then
            SaveData = oCode.UpdateCode(tCode)
        ElseIf m_sOperate = ID_DELETE Then
            SaveData = oCode.DeleteCode(tCode.sCodeID)
        End If

        Set oCode = Nothing
    
    Case 2  ' 등급 관리
        Set oGrade = New PlusLib2.CGrade
        oGrade.Connection = g_adoCon
        oGrade.UserName = g_sUserName
        
        TGrade.GradeID = txtName(11)
        TGrade.Grade = txtName(12)
        
        If m_sOperate = ID_ADDNEW Then
            SaveData = oGrade.AddNewGrade(TGrade)
        ElseIf m_sOperate = ID_UPDATE Then
            SaveData = oGrade.UpdateGrade(TGrade)
        ElseIf m_sOperate = ID_DELETE Then
            SaveData = oGrade.DeleteGrade(TGrade.GradeID)
        End If
    
    End Select

    Exit Function

ErrHandler:
    
    Set oDefect = Nothing
    Set oCode = Nothing
    Set oGrade = Nothing
    
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Function

Private Sub ChangeScroll()
    With grdData(tabForm.Tab)
        If tabForm.Tab = 0 Then
            .ColWidth(5) = .ColWidth(5) - IIf(.Rows > LIMIT_ROW, 100, 0)
        Else
        
            .ColWidth(2) = .ColWidth(2) - IIf(.Rows > LIMIT_ROW, 100, 0)
        End If
    End With
End Sub
