VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBTCalc 
   ClientHeight    =   9285
   ClientLeft      =   2445
   ClientTop       =   1515
   ClientWidth     =   11865
   Icon            =   "frmBTCalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   11865
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   8655
      TabIndex        =   63
      Top             =   450
      Width           =   1425
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   2
      Left            =   8655
      TabIndex        =   56
      Top             =   105
      Width           =   1425
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   10185
      TabIndex        =   2
      Top             =   8550
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   315
      Left            =   10620
      TabIndex        =   32
      Top             =   8745
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   556
      _Version        =   196609
      Caption         =   "SSCommand1"
   End
   Begin VB.TextBox txtTemp 
      Height          =   270
      Left            =   8010
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   8205
      Visible         =   0   'False
      Width           =   750
   End
   Begin VSFlex7LCtl.VSFlexGrid grdBt 
      Height          =   7545
      Left            =   30
      TabIndex        =   12
      Top             =   930
      Width           =   3495
      _cx             =   6165
      _cy             =   13309
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   30
      TabIndex        =   17
      Top             =   8595
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1164
      _Version        =   196609
      Begin VB.OptionButton optID 
         Caption         =   "B/T번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optID 
         Caption         =   "의뢰번호"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   90
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   7560
      Left            =   3570
      TabIndex        =   15
      Top             =   930
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   13335
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSFrame frmBT 
         Height          =   4200
         Left            =   45
         TabIndex        =   34
         Top             =   3300
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   7408
         _Version        =   196609
         Caption         =   " B/T 처방내역 "
         Begin VB.CommandButton cmdoperate 
            Caption         =   "추가(&A)"
            Height          =   690
            Index           =   0
            Left            =   1635
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   70
            ToolTipText     =   "자료 추가"
            Top             =   195
            Width           =   765
         End
         Begin VB.CommandButton cmdoperate 
            Caption         =   "수정(&U)"
            Height          =   690
            Index           =   1
            Left            =   2400
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   69
            ToolTipText     =   "자료 수정"
            Top             =   195
            Width           =   765
         End
         Begin VB.CommandButton cmdoperate 
            Caption         =   "삭제(&D)"
            Height          =   690
            Index           =   2
            Left            =   3165
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   68
            ToolTipText     =   "자료 삭제"
            Top             =   195
            Width           =   765
         End
         Begin VB.CommandButton cmdoperate 
            Caption         =   "저장(&S)"
            Height          =   690
            Index           =   3
            Left            =   105
            Style           =   1  '그래픽
            TabIndex        =   67
            ToolTipText     =   "자료저장"
            Top             =   195
            Width           =   765
         End
         Begin VB.CommandButton cmdoperate 
            Caption         =   "취소(&X)"
            Height          =   690
            Index           =   4
            Left            =   870
            Style           =   1  '그래픽
            TabIndex        =   66
            ToolTipText     =   "입력취소"
            Top             =   195
            Width           =   765
         End
         Begin Threed.SSCommand cmdLotDown 
            Height          =   420
            Left            =   1680
            TabIndex        =   44
            Top             =   930
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   741
            _Version        =   196609
         End
         Begin Threed.SSCommand cmdLotUP 
            Height          =   420
            Left            =   1215
            TabIndex        =   43
            Top             =   930
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   741
            _Version        =   196609
         End
         Begin VB.ComboBox cboLot 
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
            Left            =   90
            Style           =   1  '단순 콤보
            TabIndex        =   38
            Text            =   "cboLot"
            Top             =   1005
            Width           =   1095
         End
         Begin Threed.SSCommand cmdAddNew 
            Height          =   450
            Left            =   2355
            TabIndex        =   39
            Top             =   900
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   794
            _Version        =   196609
            Caption         =   "염료추가"
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   450
            Left            =   3195
            TabIndex        =   40
            Top             =   900
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   794
            _Version        =   196609
            Caption         =   "염료삭제"
         End
         Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
            Height          =   2745
            Left            =   75
            TabIndex        =   41
            Top             =   1410
            Width           =   3900
            _cx             =   6879
            _cy             =   4842
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
      Begin Threed.SSFrame frmConfirm 
         Height          =   4185
         Left            =   4155
         TabIndex        =   33
         Top             =   3300
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   7382
         _Version        =   196609
         Caption         =   " B/T 승인 내역 "
         Begin VB.TextBox txtConfirmDate 
            Height          =   315
            Left            =   2355
            TabIndex        =   35
            Top             =   930
            Width           =   1650
         End
         Begin VSFlex7LCtl.VSFlexGrid grdConfirm 
            Height          =   2745
            Left            =   75
            TabIndex        =   42
            Top             =   1395
            Width           =   3900
            _cx             =   6879
            _cy             =   4842
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
         Begin Threed.SSCommand cmdConfirm 
            Height          =   570
            Left            =   465
            TabIndex        =   71
            Top             =   255
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   1005
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "       승 인"
            PictureAlignment=   1
         End
         Begin Threed.SSCommand cmdUnConfirm 
            Height          =   570
            Left            =   2220
            TabIndex        =   72
            Top             =   240
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   1005
            _Version        =   196609
            BackColor       =   12648384
            Caption         =   "       승인취소"
            PictureAlignment=   1
         End
         Begin VB.Label lblLot 
            Caption         =   "Lot"
            Height          =   270
            Left            =   105
            TabIndex        =   37
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "승인일자 :"
            Height          =   315
            Left            =   1395
            TabIndex        =   36
            Top             =   1005
            Width           =   855
         End
      End
      Begin Threed.SSPanel pnlData 
         Height          =   2880
         Left            =   4110
         TabIndex        =   20
         Top             =   75
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5080
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   8
            Left            =   2250
            TabIndex        =   60
            Top             =   1470
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "차 수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   21
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "B/T 번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   22
            Top             =   435
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "거  래  처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   780
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "의뢰번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   24
            Top             =   1125
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "품      명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   25
            Top             =   1470
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "색      수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   26
            Top             =   1815
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "실 험 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtBT 
            Height          =   315
            Index           =   0
            Left            =   1275
            TabIndex        =   27
            Top             =   2160
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65863681
            CurrentDate     =   37575
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   28
            Top             =   2160
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "접수일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   29
            Top             =   2505
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "발송일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtBT 
            Height          =   315
            Index           =   1
            Left            =   1275
            TabIndex        =   30
            Top             =   2505
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65863681
            CurrentDate     =   37575
         End
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   0
            Left            =   1275
            TabIndex        =   45
            Top             =   90
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
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
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   1
            Left            =   1275
            TabIndex        =   46
            Top             =   435
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
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
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   2
            Left            =   1275
            TabIndex        =   47
            Top             =   780
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
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
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   3
            Left            =   1275
            TabIndex        =   48
            Top             =   1125
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
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
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   4
            Left            =   1275
            TabIndex        =   49
            Top             =   1470
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
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
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   5
            Left            =   1275
            TabIndex        =   50
            Top             =   1815
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
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
         Begin MRPPlus2.WizText txtData 
            Height          =   315
            Index           =   6
            Left            =   3075
            TabIndex        =   61
            Top             =   1470
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
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
      End
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   2880
         Left            =   60
         TabIndex        =   16
         Top             =   90
         Width           =   3975
         _cx             =   7011
         _cy             =   5080
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   690
      Left            =   10770
      TabIndex        =   14
      Top             =   45
      Width           =   795
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   5160
      TabIndex        =   4
      Top             =   450
      Width           =   1425
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1425
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   345
      Width           =   570
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   570
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   7395
      Top             =   8670
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   4
      Left            =   3615
      TabIndex        =   5
      Top             =   450
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   1410
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   0
      Left            =   6615
      TabIndex        =   7
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
      Index           =   9
      Left            =   3615
      TabIndex        =   8
      Top             =   105
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
         TabIndex        =   9
         Top             =   60
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   2235
      TabIndex        =   10
      Top             =   -15
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Format          =   65863681
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   2235
      TabIndex        =   11
      Top             =   285
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Format          =   65863681
      CurrentDate     =   36871
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   13
      Top             =   0
      Width           =   0
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   675
      TabIndex        =   51
      Top             =   15
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T 접수일자"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   52
         Top             =   60
         Value           =   1  '확인
         Width           =   1410
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   645
      TabIndex        =   53
      Top             =   585
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T 발송일자"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   54
         Top             =   60
         Width           =   1425
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   2
      Left            =   2235
      TabIndex        =   55
      Top             =   585
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   65863681
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   12
      Left            =   7110
      TabIndex        =   57
      Top             =   105
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "처 방 자"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   58
         Top             =   60
         Width           =   990
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   6615
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   450
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   2
      Left            =   10110
      TabIndex        =   62
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
      Index           =   6
      Left            =   7110
      TabIndex        =   64
      Top             =   450
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T NO"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   65
         Top             =   60
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmBTCalc"
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
Private Const LIMIT_ROW5 = 4
Private Const LIMIT_WIDTH1 = 1380
Private Const LIMIT_WIDTH2 = 1635
Private Const LIMIT_WIDTH3 = 1965
Private Const LIMIT_WIDTH4 = 2085
Private Const LIMIT_WIDTH5 = 1890
Private Const LIMIT_WIDTH6 = 1000

Private m_sFlag        As String
Private m_nSelected    As Integer
Private m_bloading     As Boolean
Private m_bLoadingColor As Boolean
Private m_bSortForward As Boolean

Private m_tPrevDyeAux() As PlusLib2.tBtPrevDyeAux


Private Sub cboLot_Click()
    If grdBt.Rows > grdBt.FixedRows Then
        Call ShowDyeAux
    End If
End Sub


Private Sub cmdConfirm_Click()
    Dim sBTID$, nBTSeq%, nColorSeq%
    Dim nConfirmLot%, sDate$
    Dim oBt As PlusLib2.CBt
    Dim bConfirm As Boolean
    
    On Error GoTo ErrHandler
    
    nConfirmLot = cboLot.ItemData(cboLot.ListIndex)
    
    If grdDyeAux.Rows = grdDyeAux.FixedRows Then
        Exit Sub
    End If
    
    If grdColor.TextMatrix(grdColor.Row, 3) = "1" Then
        If grdColor.TextMatrix(grdColor.Row, 4) = nConfirmLot Then
            MsgBox ("이미 " & nConfirmLot & "번 Lot 로 승인되어 있습니다")
            
            Exit Sub
        End If
    
        If MsgBox("이미 승인받은 Color입니다." & vbCrLf & vbCrLf & "새로운 Lot로 변경하시겠습니까?", vbOKCancel) = vbCancel Then
        
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    sBTID = MakeBTID(grdBt.TextMatrix(grdBt.Row, 2), OM_REDUCE)
    nBTSeq = grdBt.TextMatrix(grdBt.Row, 4)
    nColorSeq = grdColor.TextMatrix(grdColor.Row, 2)
    nConfirmLot = cboLot.ItemData(cboLot.ListIndex)
    sDate = MakeDate(DF_SHORT, Date)
    
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    oBt.UserName = g_sUserName
    
    bConfirm = oBt.UpdateBTConfirm(sBTID, nBTSeq, nColorSeq, "1", nConfirmLot, sDate)

    Set oBt = Nothing
    
    Screen.MousePointer = vbDefault
    
    Call ShowData
    
    Exit Sub
    
ErrHandler:
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 0 Then
        Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
    ElseIf Index = 1 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_PERSON, , False, txtSearch(2))
    End If
End Sub

Private Sub cmdLotDown_Click()
    If cboLot.ListIndex < 14 Then
        cboLot.ListIndex = cboLot.ListIndex + 1
    End If
End Sub

Private Sub cmdLotUP_Click()
    If cboLot.ListIndex > 0 Then
        cboLot.ListIndex = cboLot.ListIndex - 1
    End If
    
End Sub

Private Sub cmdUnConfirm_Click()
    Dim sBTID$, nBTSeq%, nColorSeq%, nRecipeSeq%, sDate$
    Dim oBt As PlusLib2.CBt
    Dim bConfirm As Boolean
    
    On Error GoTo ErrHandler
    
    nRecipeSeq = cboLot.ItemData(cboLot.ListIndex)

    If grdColor.TextMatrix(grdColor.Row, 3) = "1" Then
        
        If MsgBox("승인을 취소하시겠습니까?", vbOKCancel) = vbCancel Then
        
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
        
    sBTID = MakeBTID(grdBt.TextMatrix(grdBt.Row, 2), OM_REDUCE)
    nBTSeq = grdBt.TextMatrix(grdBt.Row, 4)
    nColorSeq = grdColor.TextMatrix(grdColor.Row, 2)
        
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    oBt.UserName = g_sUserName
    
    bConfirm = oBt.UpdateBTConfirm(sBTID, nBTSeq, nColorSeq, "0", 0, "0")

    Set oBt = Nothing
    
    Screen.MousePointer = vbDefault
    
    Call ShowData
    
    Exit Sub
    
ErrHandler:
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)

    dtpDate(0) = Now
    dtpDate(1) = Now
    dtpDate(2) = Now

    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdConfirm.Picture = LoadResPicture("ORDER", vbResIcon)
    cmdUnConfirm.Picture = LoadResPicture("SHUT", vbResIcon)
    cmdLotUP.Picture = LoadResPicture("BACK", vbResIcon)
    cmdLotDown.Picture = LoadResPicture("FRONT", vbResIcon)
    
    optID(0).Value = True
    pnlData.Enabled = False
    cmdAddNew.Enabled = False
    cmdDelete.Enabled = False

    Call InitGrid
    
    ReDim m_tPrevDyeAux(0)

    grdDyeAux.Editable = flexEDNone
    
    With cboLot
        For i = 1 To 15
            .AddItem "Lot " & i
            .ItemData(.NewIndex) = i
        Next i
        .ListIndex = 0
    End With
    
    Show

    txtSearch(0).Enabled = False
    txtSearch(1).Enabled = False
    cmdFind(0).Enabled = False
    cmdFind(1).Enabled = False
    cmdFind(2).Enabled = False
    
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 3 Or Index = 4 Then
    
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
        End If
    
    Else
    
        If Index = 5 Then
            
            If chkSearch(5).Value = vbChecked Then
                txtSearch(3).Enabled = True
                txtSearch(3).SetFocus
            Else
                txtSearch(3).Enabled = False
                
            End If
            
        Else
            If chkSearch(Index) Then
                txtSearch(Index).Enabled = True
                txtSearch(Index).SetFocus
                cmdFind(Index).Enabled = True
                
            Else
                txtSearch(Index).Enabled = False
                cmdFind(Index).Enabled = False
                cmdSearch.SetFocus
                
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




Private Sub cmdSearch_Click()
    Call FillGridBt
End Sub


Private Sub grdBt_RowColChange()
    If m_bloading Then Exit Sub

    With grdBt
        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Sub

        Call ShowData

        .SetFocus
    End With
End Sub

Private Sub ShowData()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim sBTID$, nBTSeq%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    With grdBt
        txtData(0) = .TextMatrix(.Row, 3)       ' B/T NO
        txtData(1) = .TextMatrix(.Row, 1)       ' 거래처
        txtData(2) = .TextMatrix(.Row, 2)       ' 접수번호
        txtData(3) = .TextMatrix(.Row, 5)       ' 품명
        txtData(4) = .TextMatrix(.Row, 6)       ' 색상수
        txtData(5) = .TextMatrix(.Row, 7)       ' 실험자
        txtData(6) = .TextMatrix(.Row, 4)       ' 차수
        dtBT(0) = MakeDate(DF_FULL, .TextMatrix(.Row, 11))
        dtBT(1) = IIf(Len(.TextMatrix(.Row, 12)) = 0, Now, MakeDate(DF_FULL, .TextMatrix(.Row, 12)))
    
        sBTID = .TextMatrix(.Row, 9)
        nBTSeq = .TextMatrix(.Row, 4)
    End With
    

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    Set rs = oBt.GetBtSub(sBTID, nBTSeq)
    Set oBt = Nothing


    m_bLoadingColor = True
    
    With grdColor
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!Color & vbTab & rs!ColorSeq & vbTab & _
                CheckNull(rs!ConfClss) & vbTab & CheckNull(rs!ConfLotNO) & vbTab & CheckNull(rs!ConfDate)
            
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

        If .Rows > .FixedRows Then
            cmdPrint.Enabled = True
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
            cmdOperate(ID_UPDATE).Enabled = True
            cmdOperate(ID_DELETE).Enabled = True
        Else
            cmdPrint.Enabled = False
            .HighLight = flexHighlightNever
            cmdOperate(ID_UPDATE).Enabled = False
            cmdOperate(ID_DELETE).Enabled = False
        End If

        .Redraw = flexRDDirect
        .SetFocus
    End With
    

    Screen.MousePointer = vbDefault

    m_bLoadingColor = False
    
    Call ShowDyeAux

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oBt = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub



Private Sub ShowConfirmData()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim sBTID$, nBTSeq%, nLot%, nColorSeq%
    Dim sDate$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    sBTID = MakeBTID(grdBt.TextMatrix(grdBt.Row, 2), OM_REDUCE)
    nBTSeq = grdBt.TextMatrix(grdBt.Row, 4)
    nColorSeq = grdColor.TextMatrix(grdColor.Row, 2)
    nLot = grdColor.TextMatrix(grdColor.Row, 4)      'lot
    sDate = MakeDate(DF_FULL, grdColor.TextMatrix(grdColor.Row, 5))
    
    txtConfirmDate = sDate
    lblLot.Caption = "LotNO : " & nLot
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    Set rs = oBt.GetBtDyeAux(sBTID, nBTSeq, nColorSeq, nLot)
    Set oBt = Nothing

    With grdConfirm
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!DyeAux & vbTab & rs!DyeAuxRate & vbTab & rs!DyeAuxID
            
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

        If .Rows > .FixedRows Then
            cmdPrint.Enabled = True
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            
        End If

'        Call ChangeScrollColor

        .Redraw = flexRDDirect
        .SetFocus
    End With

    Screen.MousePointer = vbDefault
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oBt = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub



Private Sub ShowDyeAux(Optional nShowOption As Integer = 0)
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim sBTID$, nBTSeq%, nColorSeq%, nLot%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    If m_bLoadingColor Then Exit Sub
    
    m_bloading = True
    
        
    With grdBt
        sBTID = .TextMatrix(.Row, 9)
        nBTSeq = .TextMatrix(.Row, 4)
    End With
    nColorSeq = grdColor.TextMatrix(grdColor.Row, 2)
    
    nLot = cboLot.ItemData(cboLot.ListIndex)
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    Set rs = oBt.GetBtDyeAux(sBTID, nBTSeq, nColorSeq, nLot)
    Set oBt = Nothing

    With grdDyeAux
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        ' 염료 항목 저장... 새로 추가시 사용
        If rs.RecordCount > 0 Then
            ReDim m_tPrevDyeAux(rs.RecordCount)
        End If
        
        If rs.RecordCount > 0 Then
        
            For i = 1 To rs.RecordCount
                .AddItem CStr(i) & vbTab & CheckNull(rs!DyeAux) & vbTab & vbTab & rs!DyeAuxRate & vbTab & rs!DyeAuxID & vbTab & CheckNull(rs!DyeAux)
                
                If (i Mod 2) = 0 Then
                    .Row = .FixedRows + i - 1
                    .Col = .FixedCols
                    .ColSel = .Cols - 1
                    .CellBackColor = &HE0E0E0
                End If
                
                With m_tPrevDyeAux(i)
                    .sDyeAux = CheckNull(rs!DyeAux)
                    .sDyeAuxID = rs!DyeAuxID
                    .nRecipeQty = rs!DyeAuxRate
                End With
                
                .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
                .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter
    
                rs.MoveNext
            Next i
        End If

        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            cmdPrint.Enabled = True
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
            cmdOperate(ID_UPDATE).Enabled = True
            cmdOperate(ID_DELETE).Enabled = True
            cmdOperate(ID_ADDNEW).Enabled = False
        Else
            .HighLight = flexHighlightNever
            cmdOperate(ID_ADDNEW).Enabled = True
            cmdOperate(ID_UPDATE).Enabled = False
            cmdOperate(ID_DELETE).Enabled = False
        End If

        Call ChangeScrollColor

        .Redraw = flexRDDirect
        .SetFocus
    End With

    Screen.MousePointer = vbDefault

    m_bloading = False

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oBt = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub ShowPrevDyeAux()
    Dim i%
    
   With grdDyeAux
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        ' 염료 항목 저장... 새로 추가시 사용
        For i = 1 To UBound(m_tPrevDyeAux)
            .AddItem CStr(i) & vbTab & m_tPrevDyeAux(i).sDyeAux & vbTab & vbTab & "0" & vbTab & m_tPrevDyeAux(i).sDyeAuxID
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If
                            
            .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter

        Next i

        .Redraw = flexRDDirect
        .SetFocus
        
    End With

End Sub

Private Sub grdColor_RowColChange()
    cboLot.ListIndex = 0
    
    ReDim m_tPrevDyeAux(0)
    
    Call ShowDyeAux
    
    If grdColor.TextMatrix(grdColor.Row, 3) = "1" Then
        Call ShowConfirmData
    Else
        grdConfirm.Rows = grdConfirm.FixedRows
        lblLot.Caption = "LotNO : "
        txtConfirmDate = " "
    End If

End Sub


Private Sub grdDyeAux_DblClick()
    With grdDyeAux
        .EditCell
    End With
End Sub


Private Sub optID_Click(Index As Integer)
    With grdBt
        If optID(0).Value Then
            .ColHidden(2) = False
            .ColHidden(3) = True
        Else
            .ColHidden(2) = True
            .ColHidden(3) = False
        End If
        
    End With
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
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(0))
        ElseIf Index = 3 Then
            Call ReturnCode(LG_PERSON, 8, False, txtSearch(3))
        End If
        
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim i%
    
    Select Case Index
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW

            cmdOperate(0).Enabled = False
            cmdOperate(1).Enabled = False
            cmdOperate(2).Enabled = False
            
            cmdOperate(3).Visible = True
            cmdOperate(4).Visible = True
            
            cmdLotUP.Visible = False
            cmdLotDown.Visible = False
            
            txtData(5).Locked = False
            cmdAddNew.Enabled = True
            cmdDelete.Enabled = True
            grdDyeAux.Editable = flexEDKbdMouse
            
            Call ShowPrevDyeAux '이전 염료자료 출력
            
        Case ID_UPDATE
            If grdBt.Rows = grdBt.FixedRows Then Exit Sub
            
            m_sFlag = ID_UPDATE
            cmdOperate(0).Enabled = False
            cmdOperate(1).Enabled = False
            cmdOperate(2).Enabled = False
            
            cmdOperate(3).Visible = True
            cmdOperate(4).Visible = True
                        
            cmdLotUP.Visible = False
            cmdLotDown.Visible = False
            
            txtData(5).Locked = False
            cmdAddNew.Enabled = True
            cmdDelete.Enabled = True
            grdDyeAux.Editable = flexEDKbdMouse
   
        Case ID_DELETE
            If grdBt.Rows = grdBt.FixedRows Then Exit Sub
            If grdDyeAux.Rows = grdDyeAux.FixedRows Then Exit Sub
            If Not QuestionBox(cboLot & " - 삭제하시겠습니까?") Then Exit Sub

            If DeleteData() Then Call ShowDyeAux
            
        Case ID_SAVE
            If SaveData() Then
                cmdOperate(0).Enabled = IIf(grdDyeAux.Rows > grdDyeAux.FixedRows, False, True)
                cmdOperate(1).Enabled = True
                cmdOperate(2).Enabled = True
                
                cmdAddNew.Enabled = False
                cmdDelete.Enabled = False
            
                cmdOperate(3).Visible = False
                cmdOperate(4).Visible = False
            
                Call ShowDyeAux
                m_sFlag = ID_SAVE
            End If
            
            cmdLotUP.Visible = True
            cmdLotDown.Visible = True
            
            grdDyeAux.Editable = flexEDNone
        Case ID_CANCEL
        
            If m_sFlag = ID_ADDNEW Then
                If cboLot.ListIndex > 0 Then
                    cboLot.ListIndex = cboLot.ListIndex - 1
                End If
            End If
            
            cmdOperate(0).Enabled = IIf(grdDyeAux.Rows > grdDyeAux.FixedRows, False, True)
            cmdOperate(1).Enabled = True
            cmdOperate(2).Enabled = True
            cmdOperate(3).Visible = False
            cmdOperate(4).Visible = False
        
            cmdAddNew.Enabled = False
            cmdDelete.Enabled = False
            
            cmdLotUP.Visible = True
            cmdLotDown.Visible = True
            
            grdDyeAux.Editable = flexEDNone
    End Select
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdAddNew_Click()
    Dim i%

    With grdDyeAux
        .Rows = .Rows + 1

        Call ChangeScrollDyeAux
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = CStr(i)
        Next i

        .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
        .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter
        .SetFocus
        .Select .Rows - 1, 1
        
    End With
End Sub

Private Sub cmdDelete_Click()
    Dim i%
    
    With grdDyeAux
        If .Rows = 1 Or .Row < 1 Then
            MsgBox LoadResString(200), vbInformation
        Else
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                .RemoveItem .Row
                
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, 0) = CStr(i)
                Next i
            End If

        End If
    End With

End Sub


Private Sub grdDyeAux_Click()
    With grdDyeAux
        If .MouseRow < .FixedRows Or .MouseRow > .Rows - 1 Or .MouseCol <> 2 Then Exit Sub
        If Not cmdOperate(3).Visible Then Exit Sub
        Dim Row%
        Row = .MouseRow

        If ReturnCode(LG_DYE, , True, txtTemp) Then
            .TextMatrix(Row, 1) = txtTemp
            .TextMatrix(Row, 5) = txtTemp
            .TextMatrix(Row, 4) = txtTemp.Tag
        End If
    End With
End Sub

Private Sub grdDyeAux_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdDyeAux
        Select Case Col
            Case 2
                Cancel = True
            Case 3
                If Len(.TextMatrix(Row, Col)) = 0 Then .TextMatrix(Row, Col) = "0.0000"
                .Cell(flexcpText, Row, Col) = Format(.TextMatrix(Row, Col), "###0.0000")
        End Select
    End With
End Sub

Private Sub grdDyeAux_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col <> 1 Or KeyAscii <> vbKeyReturn Then Exit Sub

    With grdDyeAux
        txtTemp = .EditText

        If ReturnCode(LG_DYE, , False, txtTemp) Then
            .TextMatrix(Row, 1) = txtTemp
            .TextMatrix(Row, 5) = txtTemp
'            .EditText = txtTemp
            .TextMatrix(Row, 4) = txtTemp.Tag
        End If
    End With
End Sub

Private Sub grdDyeAux_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdDyeAux
        If Col = 1 Then
            .Select Row, 3
'            .EditCell
        ElseIf Col = 3 Then
            .Cell(flexcpText, Row, Col) = SetCurrency(.TextMatrix(Row, Col), 4)

            If Row = .Rows - 1 Then
                If QuestionBox("염료를 계속 추가하시겠습니까 ?") Then
                    Call cmdAddNew_Click
                Else
                    If cmdOperate(3).Visible Then cmdOperate(3).SetFocus
                End If
            End If
        End If
    End With
End Sub


Private Sub cmdPrint_Click()
    Dim oBt As PlusLib2.CBt
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    
    Set rs = oBt.GetPrintBtList(MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(0), 1, 0), txtSearch(0).Tag, IIf(chkSearch(1), 1, 0), Replace(txtSearch(1), "-", ""))
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
    With grdBt
        .Cols = 13
        Call SetVSFlexGrid(grdBt)

        .Redraw = flexRDNone

        .TextArray(1) = "       거래처":                .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "접수번호":                     .ColWidth(2) = 1100:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "B/T NO":                       .ColWidth(3) = 1100:    .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "차수":                         .ColWidth(4) = 800:     .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "품명":                         .ColWidth(5) = 0:       .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "색상수":                       .ColWidth(6) = 0:       .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "실험자":                       .ColWidth(7) = 0:       .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "거래처ID":                     .ColWidth(8) = 0
        .TextArray(9) = "BTID":                         .ColWidth(9) = 0
        .TextArray(10) = "품명ID":                      .ColWidth(10) = 0
        .TextArray(11) = "접수일자":                    .ColWidth(11) = 0
        .TextArray(12) = "발송일자":                    .ColWidth(12) = 0
            
        
        .ColHidden(3) = True
        .Redraw = flexRDDirect
    End With
    
    With grdDyeAux
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux)

        .Redraw = flexRDNone

        .TextArray(1) = "염 료 명":         .ColWidth(1) = LIMIT_WIDTH4:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "염 료 명":         .ColWidth(2) = 300:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "투입비율":         .ColWidth(3) = 1200:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "염료코드":         .ColWidth(4) = 0
        .TextArray(5) = "염료명":           .ColWidth(5) = 0


        .ColFormat(3) = "###.0000"
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy

        .Redraw = flexRDDirect
    End With
    
    With grdColor
        .Cols = 6
        Call SetVSFlexGrid(grdColor)
        
        .Redraw = flexRDNone
        
        .TextArray(1) = "Color":        .ColWidth(1) = LIMIT_WIDTH3:       .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "순위":         .ColHidden(2) = True
        .TextArray(3) = "인증여부":     .ColHidden(3) = True
        .TextArray(4) = "인증LotNO":     .ColHidden(4) = True
        .TextArray(5) = "인증일자":    .ColHidden(5) = True
    
        .Redraw = flexRDDirect
    End With
    
    
    With grdConfirm
        .Cols = 4
        Call SetVSFlexGrid(grdConfirm)

        .Redraw = flexRDNone

        .TextArray(1) = "염 료 명":     .ColWidth(1) = LIMIT_WIDTH4:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "투입비율":     .ColWidth(3) = 1200:            .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "염료":         .ColWidth(3) = 0

        .ColFormat(2) = "###.0000"
        
        .FocusRect = flexFocusHeavy

        .Redraw = flexRDDirect
        
    End With
    
    
End Sub



Private Function MakeBTID(sBTID As String, nType As EORDERMAKE) As String
     If nType = OM_EXPAND Then
        MakeBTID = Left(sBTID, 2) & "-" & Mid(sBTID, 3, 2) & "-" & Mid(sBTID, 5, 4)
    Else
        MakeBTID = Replace(sBTID, "-", "")
    End If
    

End Function


Private Sub FillGridBt()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim nChkDate%, sDate$, eDate$
    Dim nChkSendDate%, SendDate$
    Dim nChkCustom%, sCustom$
    Dim nchkArticle%, sArticle$
    Dim nChkPerson%, sPersonID$
    Dim nChkBTID%, sBTID$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    
    nChkDate = IIf(chkSearch(3), 1, 0)
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkCustom = IIf(chkSearch(0), 1, 0)
    sCustom = txtSearch(0).Tag
    nchkArticle = IIf(chkSearch(1), 1, 0)
    sArticle = txtSearch(1).Tag
    nChkPerson = IIf(chkSearch(2), 1, 0)
    sPersonID = txtSearch(2).Tag
    nChkBTID = IIf(chkSearch(5), 1, 0)
    sBTID = txtSearch(3)
    
    Set rs = oBt.GetBtList(nChkDate, sDate, eDate, nChkSendDate, SendDate, nChkCustom, sCustom, _
                            nchkArticle, sArticle, nChkPerson, sPersonID, nChkBTID, sBTID)
    
    Set oBt = Nothing

    With grdBt
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & CheckNull(rs!KCustom) & vbTab & MakeBTID(rs!BTID, OM_EXPAND) & vbTab & rs!BTNO & vbTab & _
                    rs!BTIDSeq & vbTab & CheckNull(rs!Article) & vbTab & rs!ColorCnt & vbTab & CheckNull(rs!Name) & vbTab & _
                    CheckNull(rs!CustomID) & vbTab & CheckNull(rs!BTID) & vbTab & CheckNull(rs!ArticleID) & vbTab & _
                    CheckNull(rs!RecpDate) & vbTab & CheckNull(rs!SendDate)

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
            Call ShowData
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

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oBt = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Function SaveData() As Boolean
    Dim TBt()   As PlusLib2.tBtDyeAux
    Dim oBt   As PlusLib2.CBt
    Dim i%, nBTSub%

    SaveData = False
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    If grdDyeAux.Rows = grdDyeAux.FixedRows Then Exit Function

    nBTSub = (grdDyeAux.Rows - 2)
    ReDim TBt(nBTSub)

    For i = 0 To nBTSub
        If grdDyeAux.TextMatrix(grdDyeAux.FixedRows + i, 1) <> grdDyeAux.TextMatrix(grdDyeAux.FixedRows + i, 5) Then
            MsgBox "염료명을 정확히 입력해주십시오"
            Exit Function
        End If
        
        TBt(i).sBTID = MakeBTID(grdBt.TextMatrix(grdBt.Row, 2), OM_REDUCE)
        TBt(i).nBTSeq = grdBt.TextMatrix(grdBt.Row, 4)
        TBt(i).nColorSeq = grdColor.TextMatrix(grdColor.Row, 2)
        TBt(i).nLot = cboLot.ItemData(cboLot.ListIndex)
        TBt(i).nDyeAuxSeq = i + 1
        TBt(i).sDyeAuxID = grdDyeAux.TextMatrix(grdDyeAux.FixedRows + i, 4)
        If grdDyeAux.TextMatrix(grdDyeAux.FixedRows + i, 3) = "" Then
            Call MessageBox(grdDyeAux.TextMatrix(grdDyeAux.FixedRows + i, 1) & "의 투입비율이 입력되지 않았습니다" & vbCrLf & vbCrLf & _
                            "투입비율을 입력해주십시오")
            Exit Function
        Else
            TBt(i).nRecipeQty = grdDyeAux.TextMatrix(grdDyeAux.FixedRows + i, 3)
        End If
    Next i
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    oBt.UserName = g_sUserName

    If m_sFlag = ID_ADDNEW Then
        SaveData = oBt.AddNewBtDyeAux(TBt)
    Else
        SaveData = oBt.UpdateBTDyeAux(TBt)
    End If

    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    SaveData = False
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

Private Function DeleteData() As Boolean
    Dim sBTID$, nBTSeq%, nRecipeSeq%, nColorSeq%
    Dim oBt As PlusLib2.CBt

    On Error GoTo ErrHandler

    DeleteData = False

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    oBt.UserName = g_sUserName
    
    sBTID = MakeOrderID(grdBt.TextMatrix(grdBt.Row, 2), OM_REDUCE)
    nBTSeq = grdBt.TextMatrix(grdBt.Row, 4)
    nColorSeq = grdColor.TextMatrix(grdColor.Row, 2)
    nRecipeSeq = cboLot.ItemData(cboLot.ListIndex)
    
    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon
    DeleteData = oBt.DeleteBtDyeAux(sBTID, nBTSeq, nColorSeq, nRecipeSeq)
    
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


Private Sub ChangeScrollColor()
    With grdColor
        .ColWidth(1) = IIf(.Rows > 10 + .FixedRows, LIMIT_WIDTH3 - 240, LIMIT_WIDTH3)
    End With
End Sub


Private Sub ChangeScrollDyeAux()
    With grdDyeAux
        .ColWidth(1) = IIf(.Rows > LIMIT_ROW4 + .FixedRows, LIMIT_WIDTH4 - 240, LIMIT_WIDTH4)
    End With
End Sub
