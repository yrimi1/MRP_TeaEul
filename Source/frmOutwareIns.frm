VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutwareIns 
   Caption         =   "출고관리-검사(8020)"
   ClientHeight    =   10200
   ClientLeft      =   525
   ClientTop       =   1305
   ClientWidth     =   16365
   Icon            =   "frmOutwareIns.frx":0000
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   16365
   Begin Threed.SSPanel pnlRoll 
      Height          =   9255
      Left            =   960
      TabIndex        =   31
      Top             =   8790
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   16325
      _Version        =   196610
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdAdd 
         Caption         =   "▶"
         Height          =   645
         Left            =   5670
         TabIndex        =   37
         Top             =   4845
         Width           =   645
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "◀"
         Height          =   645
         Left            =   5670
         TabIndex        =   36
         Top             =   5715
         Width           =   645
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "전체삭제"
         Height          =   300
         Index           =   5
         Left            =   10860
         TabIndex        =   35
         Top             =   2520
         Width           =   900
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "전체선택"
         Height          =   300
         Index           =   4
         Left            =   9930
         TabIndex        =   34
         Top             =   2520
         Width           =   900
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "전체삭제"
         Height          =   300
         Index           =   3
         Left            =   4410
         TabIndex        =   33
         Top             =   2520
         Width           =   900
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "전체선택"
         Height          =   300
         Index           =   2
         Left            =   3480
         TabIndex        =   32
         Top             =   2520
         Width           =   900
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOutSum 
         Height          =   1455
         Left            =   6630
         TabIndex        =   38
         Top             =   6990
         Width           =   5115
         _cx             =   9022
         _cy             =   2566
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
      Begin VSFlex7LCtl.VSFlexGrid grdRollSum 
         Height          =   1455
         Left            =   90
         TabIndex        =   39
         Top             =   6990
         Width           =   5205
         _cx             =   9181
         _cy             =   2566
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
      Begin Threed.SSFrame fmeSearch 
         Height          =   2025
         Left            =   60
         TabIndex        =   40
         Top             =   450
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   3572
         _Version        =   196610
         Begin VB.TextBox txtOrderID1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   450
            Width           =   1185
         End
         Begin VB.ComboBox cboGrade 
            Height          =   300
            Left            =   4290
            Style           =   2  '드롭다운 목록
            TabIndex        =   44
            Top             =   450
            Width           =   1275
         End
         Begin VB.CommandButton cmdSearch1 
            Caption         =   "검색"
            Height          =   780
            Left            =   10890
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   43
            ToolTipText     =   "자료 저장"
            Top             =   90
            Width           =   780
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "전체삭제"
            Height          =   300
            Index           =   1
            Left            =   5910
            TabIndex        =   42
            Top             =   780
            Width           =   900
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "전체선택"
            Height          =   300
            Index           =   0
            Left            =   5910
            TabIndex        =   41
            Top             =   450
            Width           =   900
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   1
            Left            =   60
            TabIndex        =   46
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "접수번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid grdColor 
            Height          =   1875
            Left            =   6870
            TabIndex        =   47
            Top             =   90
            Width           =   4000
            _cx             =   7056
            _cy             =   3307
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
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   2
            Left            =   1320
            TabIndex        =   48
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch1 
               Caption         =   "검사 일자"
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   49
               Top             =   30
               Width           =   1095
            End
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   14
            Left            =   4320
            TabIndex        =   50
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch1 
               Caption         =   "등급 구분"
               Height          =   225
               Index           =   1
               Left            =   60
               TabIndex        =   51
               Top             =   30
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   2
            Left            =   2595
            TabIndex        =   52
            Top             =   90
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60489729
            CurrentDate     =   36871
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   3
            Left            =   2580
            TabIndex        =   53
            Top             =   450
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60489729
            CurrentDate     =   36871
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   17
            Left            =   5610
            TabIndex        =   54
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch1 
               Caption         =   "색 상 명"
               Height          =   225
               Index           =   2
               Left            =   60
               TabIndex        =   55
               Top             =   30
               Width           =   1095
            End
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "부터"
            Height          =   180
            Index           =   2
            Left            =   3900
            TabIndex        =   57
            Top             =   150
            Width           =   360
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "까지"
            Height          =   180
            Index           =   3
            Left            =   3900
            TabIndex        =   56
            Top             =   510
            Width           =   360
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdRoll 
         Height          =   4110
         Left            =   90
         TabIndex        =   58
         Top             =   2850
         Width           =   5220
         _cx             =   9208
         _cy             =   7250
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
      Begin Threed.SSCommand cmdRollQuit 
         Height          =   690
         Left            =   10155
         TabIndex        =   59
         Top             =   8505
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1217
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      닫기"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   690
         Left            =   8430
         TabIndex        =   60
         Top             =   8505
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1217
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "      확인"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   18
         Left            =   6630
         TabIndex        =   61
         Top             =   2520
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   196610
         Caption         =   "출고 내역"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOut 
         Height          =   4110
         Left            =   6630
         TabIndex        =   62
         Top             =   2850
         Width           =   5130
         _cx             =   9049
         _cy             =   7250
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
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   19
         Left            =   90
         TabIndex        =   63
         Top             =   2520
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   196610
         Caption         =   "검사 내역"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "자동패킹"
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   64
         Top             =   150
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   6  '내부 단색
         FillColor       =   &H00800000&
         Height          =   330
         Left            =   60
         Top             =   60
         Width           =   11715
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdSum 
      Height          =   300
      Left            =   0
      TabIndex        =   29
      Top             =   8070
      Width           =   15075
      _cx             =   26591
      _cy             =   529
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
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   3210
      Left            =   30
      TabIndex        =   0
      Top             =   960
      Width           =   7635
      _cx             =   13467
      _cy             =   5662
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
   Begin Threed.SSPanel pnlRollNo 
      Height          =   3630
      Left            =   7695
      TabIndex        =   17
      Top             =   960
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   6403
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlEdit 
         Height          =   2685
         Left            =   45
         TabIndex        =   19
         Top             =   870
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   4736
         _Version        =   196610
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtUnitClss 
            Alignment       =   2  '가운데 맞춤
            Enabled         =   0   'False
            Height          =   315
            Left            =   6675
            TabIndex        =   77
            Top             =   810
            Width           =   495
         End
         Begin VB.ComboBox cboWork 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1350
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   1950
            Width           =   2085
         End
         Begin VB.ComboBox cboOutClss 
            Height          =   300
            Left            =   5070
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   90
            Width           =   2085
         End
         Begin MSComCtl2.DTPicker dtpOutDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   9
            Top             =   1590
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60489728
            CurrentDate     =   37601
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   13
            Left            =   90
            TabIndex        =   25
            Top             =   1590
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "출고일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   3
            Left            =   3450
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   90
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196610
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin VB.TextBox txtCustom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   8
            Top             =   1215
            Width           =   2100
         End
         Begin VB.TextBox txtArticle 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   7
            Top             =   840
            Width           =   2100
         End
         Begin VB.TextBox txtOrder 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   465
            Width           =   2100
         End
         Begin VB.TextBox txtOrderID 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   90
            Width           =   2100
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   3
            Left            =   90
            TabIndex        =   20
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "관리번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   4
            Left            =   90
            TabIndex        =   21
            Top             =   465
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "오더번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   22
            Top             =   840
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "품      명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   6
            Left            =   90
            TabIndex        =   23
            Top             =   1950
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "가공구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   24
            Top             =   1215
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "거 래  처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   11
            Left            =   3840
            TabIndex        =   27
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "출고 구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtOutCustom 
            Height          =   300
            Left            =   1320
            TabIndex        =   11
            Top             =   2310
            Width           =   2100
            _ExtentX        =   3704
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
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   15
            Left            =   90
            TabIndex        =   28
            Top             =   2310
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "출 고  처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   7
            Left            =   3840
            TabIndex        =   65
            Top             =   825
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "수주량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   8
            Left            =   3840
            TabIndex        =   66
            Top             =   450
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "소요량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   10
            Left            =   3840
            TabIndex        =   67
            Top             =   1575
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "잔    량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtOrderQty 
            Height          =   300
            Left            =   5070
            TabIndex        =   68
            Top             =   825
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   529
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin MRPPlus2.WizText txtOutRealQty 
            Height          =   300
            Index           =   0
            Left            =   5070
            TabIndex        =   69
            Top             =   450
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin MRPPlus2.WizText txtLeftQty 
            Height          =   300
            Index           =   0
            Left            =   5070
            TabIndex        =   70
            Top             =   1575
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   16
            Left            =   3840
            TabIndex        =   71
            Top             =   1200
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "누계출고"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtOutSumQty 
            Height          =   300
            Index           =   0
            Left            =   5070
            TabIndex        =   72
            Top             =   1200
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin MRPPlus2.WizText txtOutRealQty 
            Height          =   300
            Index           =   1
            Left            =   6150
            TabIndex        =   13
            Top             =   450
            Width           =   1020
            _ExtentX        =   1799
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
            Alignment       =   1
         End
         Begin MRPPlus2.WizText txtOutSumQty 
            Height          =   300
            Index           =   1
            Left            =   6150
            TabIndex        =   73
            Top             =   1200
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin MRPPlus2.WizText txtLeftQty 
            Height          =   300
            Index           =   1
            Left            =   6150
            TabIndex        =   74
            Top             =   1590
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   0
            Left            =   3840
            TabIndex        =   75
            Top             =   1950
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            Caption         =   "비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtRemark 
            Height          =   300
            Left            =   5070
            TabIndex        =   14
            Top             =   1950
            Width           =   2100
            _ExtentX        =   3704
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
            Alignment       =   1
         End
         Begin MSComCtl2.DTPicker dtpResultDate 
            Height          =   300
            Left            =   5070
            TabIndex        =   78
            Top             =   2310
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60489728
            CurrentDate     =   37601
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   12
            Left            =   3840
            TabIndex        =   79
            Top             =   2310
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196610
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch 
               Caption         =   "출고일자"
               Height          =   210
               Index           =   4
               Left            =   60
               TabIndex        =   80
               Top             =   60
               Width           =   1080
            End
         End
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   4140
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   16
         ToolTipText     =   "자료 취소"
         Top             =   45
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   5730
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   2
         ToolTipText     =   "자료 수정"
         Top             =   45
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   6525
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   3
         ToolTipText     =   "자료 삭제"
         Top             =   45
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   4935
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 추가"
         Top             =   45
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   3345
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   15
         ToolTipText     =   "자료 저장"
         Top             =   45
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSCommand cmdInspect 
         Height          =   675
         Left            =   90
         TabIndex        =   30
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1191
         _Version        =   196610
         Caption         =   "자동패킹"
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13410
      TabIndex        =   18
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdPacking 
      Height          =   3450
      Left            =   0
      TabIndex        =   26
      Top             =   4590
      Width           =   15090
      _cx             =   26617
      _cy             =   6085
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "바탕"
         Size            =   9.75
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
      Height          =   405
      Left            =   30
      TabIndex        =   76
      Top             =   4170
      Width           =   7635
      _cx             =   13467
      _cy             =   714
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
   Begin Threed.SSFrame frmSearch 
      Height          =   885
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   1561
      _Version        =   196610
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   2040
         TabIndex        =   103
         Top             =   120
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196610
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "출고일자"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   104
            Top             =   60
            Value           =   1  '확인
            Width           =   1050
         End
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1365
         MousePointer    =   99  '사용자 정의
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   495
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1365
         MousePointer    =   99  '사용자 정의
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   105
         Width           =   600
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   6420
         TabIndex        =   85
         Top             =   120
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   6420
         TabIndex        =   84
         Top             =   510
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   9420
         TabIndex        =   83
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   720
         Left            =   10950
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   82
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3315
         TabIndex        =   88
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60489729
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3315
         TabIndex        =   89
         Top             =   510
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60489729
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   5220
         TabIndex        =   90
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   196610
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   91
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   7920
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   120
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196610
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   5220
         TabIndex        =   93
         Top             =   510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   196610
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
            TabIndex        =   94
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   7920
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   510
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196610
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
         Left            =   8280
         TabIndex        =   96
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   196610
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
            TabIndex        =   97
            Top             =   60
            Width           =   1035
         End
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   735
         Left            =   60
         TabIndex        =   98
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1296
         _Version        =   196610
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   4665
         TabIndex        =   102
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4665
         TabIndex        =   101
         Top             =   195
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11130
      TabIndex        =   105
      Top             =   8460
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "  거래명세서 양식지(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   8850
      TabIndex        =   106
      Top             =   8460
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      거래명세서 엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmOutwareIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'변경이력
' 요청 ID : S_201105_태을염직_01
' 요청자 : 김대진 대리
' 요청내용 : 거래명세서 엑셀 양식 개발
' 변경일자 : 2011.05.19
' 변경내용 : 엑셀양식 및 엑셀버튼 추가 - 유창바이오 소스 이용
'
' 요청 ID : S_201105_태을염직_02
' 요청자 : 김대진 대리
' 요청내용 : 거래명세서에 거래처  OrderNo나오게
' 변경일자 : 2011.05.25
' 변경내용 :

'--------------------------------------------------------------------------------------------
' 일자,    작업자,  요청자,     요청번호,         작업내용
'--------------------------------------------------------------------------------------------
' 2012.04. 이경미,  김대진대리, S_201204_태을염직_02 , 송장인쇄시 바로 인쇄 기능 추가
'2013.12.12   자체    오승욱   S_201312_태을염직_99   지번주소에서 도로명 주소로 입력가능하게,거래처 주소 도로명 주소 Select
'********************************************************************************************

Option Explicit

Private Const REPORTFILE   As String = "\Report\Roll.xls"
Private Const REPORTFILE1  As String = "\Report\TmpRoll.xls"

Private Const EXCEL_ROLL_ROW As Integer = 41

Private m_sOperate As String * 1
Private m_bloading As Boolean
Private m_sTranNo As String '거래명세서 발행월
Private m_nTranSeq As Integer '거래명세서 순번
Private m_sOrderID As String
Private m_nOutSeq  As Integer

Private Sub cboOutClss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call NextFocus
    End If
End Sub

Public Sub LoadOutWareIns(ByVal OrderID As String, ByVal OutSeq As Integer)
    Dim II As Integer
    Me.Show
    optOrder(1).Value = True
    chkSearch(0).Value = 0
    chkSearch(1).Value = 0
    chkSearch(2).Value = 0
    chkSearch(3).Value = 1
    txtSearch(3).Text = OrderID
    Call FillGridOrder
    With grdOrder
        For II = .FixedRows To .Rows - 1
            If .TextMatrix(II, 13) = OutSeq Then
                .Select II, 0
                If m_bloading Then Exit Sub
            
                Call ShowData
                
            End If
        Next II
    End With
End Sub

Private Sub chkSearch1_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch1(0).Value Then
            dtpDate(2).Enabled = True
            dtpDate(3).Enabled = True
        Else
            dtpDate(2).Enabled = False
            dtpDate(3).Enabled = False
        End If
    ElseIf Index = 1 Then
        If chkSearch1(Index).Value Then
            cboGrade.Enabled = True
        Else
            cboGrade.Enabled = False
        End If
    ElseIf Index = 2 Then
        If chkSearch1(Index).Value Then
            grdColor.Enabled = True
            cmdSelect(0).Enabled = True
            cmdSelect(1).Enabled = True
        Else
            grdColor.Enabled = False
            cmdSelect(0).Enabled = False
            cmdSelect(1).Enabled = False
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim i%, j%, k%, nRow%
    
    With grdRoll
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 1) = True And .TextMatrix(i, 10) <> "*" Then
                grdOut.AddItem "", grdOut.Rows
                
                grdOut.TextMatrix(grdOut.Rows - 1, 0) = grdOut.Rows - 1
                grdOut.TextMatrix(grdOut.Rows - 1, 1) = True
                grdOut.TextMatrix(grdOut.Rows - 1, 2) = .TextMatrix(i, 2) '색상명
                grdOut.TextMatrix(grdOut.Rows - 1, 3) = .TextMatrix(i, 3) 'Lot
                grdOut.TextMatrix(grdOut.Rows - 1, 4) = .TextMatrix(i, 4) '절번호
                grdOut.TextMatrix(grdOut.Rows - 1, 5) = .TextMatrix(i, 5) '수량
                grdOut.TextMatrix(grdOut.Rows - 1, 6) = .TextMatrix(i, 6) 'Loss
                grdOut.TextMatrix(grdOut.Rows - 1, 7) = .TextMatrix(i, 7) '색상순위
                grdOut.TextMatrix(grdOut.Rows - 1, 8) = .TextMatrix(i, 8) '합불
                grdOut.TextMatrix(grdOut.Rows - 1, 9) = .TextMatrix(i, 9) 'RollSeq
                
                .TextMatrix(i, 10) = "*"
                For k = grdRollSum.FixedRows To grdRollSum.Rows - grdRollSum.FixedRows
                    If grdRollSum.TextMatrix(k, 1) = .TextMatrix(i, 7) Then
                        grdRollSum.TextMatrix(k, 6) = CLng(grdRollSum.TextMatrix(k, 6)) - CLng(.TextMatrix(i, 5))
                    End If
                Next k
            End If
        Next i
        .Redraw = flexRDDirect
    End With
    
    grdOutSum.Rows = grdOutSum.FixedRows
    With grdOut
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - .FixedRows
            If grdOutSum.Rows = grdOutSum.FixedRows Then
                For k = grdRollSum.FixedRows To grdRollSum.Rows - grdRollSum.FixedRows
                    If grdRollSum.TextMatrix(k, 1) = .TextMatrix(i, 7) Then
                        grdOutSum.AddItem ""
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 0) = grdOutSum.Rows - 1
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 1) = grdRollSum.TextMatrix(k, 1)
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 2) = grdRollSum.TextMatrix(k, 2)
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 3) = grdRollSum.TextMatrix(k, 3)
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 4) = grdRollSum.TextMatrix(k, 4)
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 5) = grdRollSum.TextMatrix(k, 5)
                        grdOutSum.TextMatrix(grdOutSum.Rows - 1, 6) = .TextMatrix(i, 5)
                        Exit For
                    End If
                Next k
            Else
                nRow = 0
                For j = grdOutSum.FixedRows To grdOutSum.Rows - grdOutSum.FixedRows
                    If grdOutSum.TextMatrix(j, 1) = .TextMatrix(i, 7) Then
                        nRow = j
                        Exit For
                    End If
                Next j
                If nRow > 0 Then
                    grdOutSum.TextMatrix(grdOutSum.Rows - 1, 6) = CLng(grdOutSum.TextMatrix(grdOutSum.Rows - 1, 6)) + CLng(.TextMatrix(i, 5))
                Else
                    For k = grdRollSum.FixedRows To grdRollSum.Rows - grdRollSum.FixedRows
                        If grdRollSum.TextMatrix(k, 1) = .TextMatrix(i, 7) Then
                            grdOutSum.AddItem ""
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 0) = grdOutSum.Rows - 1
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 1) = grdRollSum.TextMatrix(k, 1)
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 2) = grdRollSum.TextMatrix(k, 2)
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 3) = grdRollSum.TextMatrix(k, 3)
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 4) = grdRollSum.TextMatrix(k, 4)
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 5) = grdRollSum.TextMatrix(k, 5)
                            grdOutSum.TextMatrix(grdOutSum.Rows - 1, 6) = .TextMatrix(i, 5)
                            Exit For
                        End If
                    Next k
                End If
            End If
        Next i
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdExcel_Click()
    Dim oOutware As PlusLib2.COutWare
    
    On Error GoTo ErrHandler
    
    If txtOrderID.Tag <> "" Then
        Set oOutware = New PlusLib2.COutWare
        oOutware.Connection = g_adoCon
        oOutware.UserName = g_sUserName
        
        Me.PopupMenu PlusMDI.mnuPopup               ' 인쇄 미리보기, S_201204_태을염직_02 추가
        
        Call oOutware.UpdateTranNo(txtOrderID.Tag, CInt(txtOrder.Tag), m_sTranNo, m_nTranSeq)
        
        Call MakeExcelPacking                       'Excel 거래명세서 인쇄, 2011.05.19, 요청번호: S_201105_태을염직_01 에 따른 추가
        Set oOutware = Nothing
        
    End If

    Exit Sub

ErrHandler:
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdRemove_Click()
    Dim i%, j%, k%
    
    With grdOut
        .Redraw = flexRDNone
        For i = .Rows - .FixedRows To .FixedRows Step -1
            If .TextMatrix(i, 1) = False Then
                For j = grdRoll.FixedRows To grdRoll.Rows - grdRoll.FixedRows
                    If .TextMatrix(i, 9) = grdRoll.TextMatrix(j, 9) Then
                        grdRoll.TextMatrix(j, 10) = ""
                        Exit For
                    End If
                Next j
            
                For k = grdOutSum.FixedRows To grdOutSum.Rows - grdOutSum.FixedRows
                    If grdOutSum.TextMatrix(k, 1) = .TextMatrix(i, 7) Then
                        grdOutSum.TextMatrix(k, 6) = CLng(grdOutSum.TextMatrix(k, 6)) - CLng(.TextMatrix(i, 5))
                    End If
                Next k
            
                For k = grdRollSum.FixedRows To grdRollSum.Rows - grdRollSum.FixedRows
                    If grdRollSum.TextMatrix(k, 1) = .TextMatrix(i, 7) Then
                        grdRollSum.TextMatrix(k, 6) = CLng(grdRollSum.TextMatrix(k, 6)) + CLng(.TextMatrix(i, 5))
                    End If
                Next k
            
                .RemoveItem i
            End If
        Next i
        
        For i = .FixedRows To .Rows - .FixedRows
            .TextMatrix(i, 0) = i
        Next i
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub cmdInspect_Click()
    Dim i%
    
    If Len(txtOrderID.Tag) = 0 Then
        MsgBox "접수번호를 먼저선택하고 난후에 자동패킹 작업을 할수 있습니다", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    pnlRoll.Move 0, 0
    
    dtpDate(2) = Now
    dtpDate(3) = Now
    cboGrade.ListIndex = 0
        
    For i = 0 To 2
        chkSearch1(i).Value = vbUnchecked
    Next i
    grdColor.Rows = grdColor.FixedRows
    grdRoll.Rows = grdRoll.FixedRows
    grdOut.Rows = grdOut.FixedRows
    grdRollSum.Rows = grdRollSum.FixedRows
    grdOutSum.Rows = grdOutSum.FixedRows
    grdColor.Enabled = False
    
    pnlRoll.Visible = True
    txtOrderID1 = txtOrderID
    txtOrderID1.Tag = txtOrderID.Tag
    
    Call FillGridColor
    
    If m_sOperate = ID_UPDATE Then
        Call FillGridOut
        Call FillGridOutSum
        grdOut.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i%, nRow%, nCol%, nOrderSeq%, sLotNo$
    
    pnlRoll.Visible = False
    nRow = 0
    nCol = 0
    grdPacking.Rows = grdPacking.FixedRows
    
    With grdOut
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 1) = True Then
                If .TextMatrix(i, 7) <> nOrderSeq Then
                    nRow = nRow + 1
                    nCol = 0
                    grdPacking.AddItem ""
                    grdPacking.TextMatrix(nRow, 0) = 0 'Box
                    grdPacking.TextMatrix(nRow, 1) = .TextMatrix(i, 2) '색상명
                    grdPacking.TextMatrix(nRow, 2) = .TextMatrix(i, 3) 'Lot
                    grdPacking.TextMatrix(nRow, 15) = .TextMatrix(i, 7) '색상순위
                    sLotNo = .TextMatrix(i, 3)
                End If
                
                If .TextMatrix(i, 3) <> sLotNo Then
                    nRow = nRow + 1
                    nCol = 0
                    grdPacking.AddItem ""
                    grdPacking.TextMatrix(nRow, 0) = 0 'Box
                    grdPacking.TextMatrix(nRow, 1) = .TextMatrix(i, 2) '색상명
                    grdPacking.TextMatrix(nRow, 2) = .TextMatrix(i, 3) 'Lot
                    grdPacking.TextMatrix(nRow, 15) = .TextMatrix(i, 7) '색상순위
                End If
                
                If nCol = 10 Then
                    nRow = nRow + 1
                    nCol = 0
                    grdPacking.AddItem ""
                    grdPacking.TextMatrix(nRow, 0) = 0 'Box
                    grdPacking.TextMatrix(nRow, 1) = .TextMatrix(i, 2) '색상명
                    grdPacking.TextMatrix(nRow, 2) = .TextMatrix(i, 3) 'Lot
                    grdPacking.TextMatrix(nRow, 15) = .TextMatrix(i, 7) '색상순위
                    nCol = 0
                End If
                
                grdPacking.TextMatrix(nRow, 3 + nCol) = .TextMatrix(i, 5) '수량
                grdPacking.TextMatrix(nRow, 16 + nCol) = .TextMatrix(i, 9)    '절번호
    '            grdPacking.TextMatrix(nRow, 26 + nCol) = .TextMatrix(i, 6)    'Loss
                
                nOrderSeq = .TextMatrix(i, 7)
                sLotNo = .TextMatrix(i, 3)
                nCol = nCol + 1
            End If
        Next i
    End With
    Call CalcRollSum
    
    pnlRoll.Visible = False
End Sub

Private Sub cmdRollQuit_Click()
    pnlRoll.Visible = False
End Sub

Private Sub cmdSearch1_Click()
    Call FillGridRoll
    Call FillGridRollSum
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim i%

    With grdColor
        If Index = 0 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, 1) = flexChecked
            Next i
        ElseIf Index = 1 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, 1) = flexUnchecked
            Next i
        End If
    End With
    
    With grdRoll
        If Index = 2 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, 1) = flexChecked
            Next i
        ElseIf Index = 3 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, 1) = flexUnchecked
            Next i
        End If
    End With
    
    With grdOut
        If Index = 4 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, 1) = flexChecked
            Next i
        ElseIf Index = 5 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, 1) = flexUnchecked
            Next i
        End If
    End With
End Sub

Private Sub dtpOutDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call NextFocus
    End If
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 15300, 9660

    Call SetOperate(Me)
    Call ChangeMode(Me, True)
    Call MakeCodeCombo(cboWork, CD_WORK)
    Call MakeCodeCombo(cboGrade, CD_GRADE)
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    dtpDate(2) = Now
    dtpDate(3) = Now
    dtpOutDate = Now
    dtpResultDate = Now

    Call InitGrid

    For i = 1 To 3
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
        
    With cboOutClss
        .AddItem "1. 정상출고":        .ItemData(0) = 1
        .AddItem "2. 반입":            .ItemData(1) = 2
        .AddItem "3. 제직불량":        .ItemData(2) = 3
        .AddItem "4. 가공불량":        .ItemData(3) = 4
        .AddItem "5. Sample, 시가공":  .ItemData(4) = 5
        .AddItem "6. 재출고":          .ItemData(5) = 6
        
        .ListIndex = 0
    End With
    
''    With cboResultClss
''        .AddItem "0. 전체":                      .ItemData(0) = 1
''        .AddItem "1. 출고":                  .ItemData(1) = 2
''        .AddItem "2. 미출고":                .ItemData(2) = 3
''        .ListIndex = 0
''    End With
    
    cmdSelect(0).Enabled = False
    cmdSelect(1).Enabled = False
    cmdInspect.Enabled = False
    dtpDate(2).Enabled = False
    dtpDate(3).Enabled = False
    cboGrade.Enabled = False
    pnlRoll.Visible = False
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(Index) Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True

            dtpDate(0).SetFocus
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False

            cmdSearch.SetFocus
        End If
    ElseIf Index >= 1 And Index <= 3 Then
        If chkSearch(Index) Then
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
        Else
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
            txtSearch(Index).Enabled = False
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdColor_Click()
    With grdColor
        If .Row < .FixedRows Or .Col <> 1 Then Exit Sub
        
        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, 1) = flexChecked
        End If
    End With
End Sub

Private Sub grdOut_Click()
    With grdOut
        If .Row < .FixedRows Or .Col <> 1 Then Exit Sub
        
        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, 1) = flexChecked
        End If
    End With
End Sub

Private Sub grdRoll_Click()
    With grdRoll
        If .Row < .FixedRows Or .Col <> 1 Then Exit Sub
        
        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, 1) = flexChecked
        End If
    End With
End Sub



Private Sub pnlName_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch1(0).Value Then
            dtpDate(2).Enabled = True
            dtpDate(3).Enabled = True
        Else
            dtpDate(2).Enabled = False
            dtpDate(3).Enabled = False
        End If
    ElseIf Index = 1 Then
        If chkSearch1(Index).Value Then
            cboGrade.Enabled = True
        Else
            cboGrade.Enabled = False
        End If
    ElseIf Index = 2 Then
        If chkSearch1(Index).Value Then
            grdColor.Enabled = True
            cmdSelect(0).Enabled = True
            cmdSelect(1).Enabled = True
        Else
            grdColor.Enabled = False
            cmdSelect(0).Enabled = False
            cmdSelect(1).Enabled = False
        End If
    End If
End Sub

Private Sub txtOutRealQty_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 And KeyCode = vbKeyReturn Then
        grdPacking.Select 1, 1
    End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset

    On Error GoTo ErrHandler
    
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    ElseIf Index = 3 Then

        Call ReturnCode(LG_ORDER, , False, txtOrderID)
        Set oOutware = New PlusLib2.COutWare
        oOutware.Connection = g_adoCon
        
        Set rs = oOutware.GetOrderOne(txtOrderID.Tag)
        Set oOutware = Nothing
        
        If Not rs.EOF Then
            txtOrderID = txtOrderID.Tag
            txtOrder = rs!OrderNo
            txtCustom = rs!kCustom
            txtArticle = rs!Article
            cboWork.ListIndex = FindComboBox(cboWork, CLng(rs!WorkID))
            txtOutRealQty(0) = 0
            txtOutRealQty(1) = 0
            txtLeftQty(0) = Format(CheckNum(rs!OrderQty) - CheckNum(rs!OutSumQty), "#,##0")
            txtLeftQty(1) = Format(CheckNum(rs!OrderQty) - CheckNum(rs!OutSumQty), "#,##0")
            txtLeftQty(0).Tag = rs!ChunkRate
            txtOutCustom.Tag = rs!UnitClss
            txtOutCustom = CheckNull(rs!DvlyPlace)
            txtOutSumQty(0) = Format(CheckNum(rs!OutSumQty), "#,##0")
            txtOutSumQty(1) = Format(CheckNum(rs!OutSumQty), "#,##0")
            txtRemark = rs!OutTelNO
            rs.Close
            Set rs = Nothing
            Call MakeColorGridCombo
        Else
            txtOrder = ""
            txtArticle = ""
            txtArticle.Tag = ""
            txtCustom = ""
            txtCustom.Tag = ""
            txtOrderQty = 0
            txtOutRealQty(0) = 0
            txtOutRealQty(1) = 0
            cboWork.ListIndex = -1
            txtLeftQty(0) = 0
            txtLeftQty(1) = 0
            txtLeftQty(0).Tag = 0
            txtOutCustom = ""
            txtOutCustom.Tag = ""
            txtOutSumQty(0) = 0
            txtOutSumQty(1) = 0
            txtRemark = ""
        End If
    End If

    Exit Sub
ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmOutwareIns.cmdFind_click", Err.Description)
End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub

Private Sub grdOrder_RowColChange()
    If m_bloading Then Exit Sub

    Call ShowData
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Select Case Index
    Case ID_ADDNEW
        m_sOperate = ID_ADDNEW
        Call ClearData

        Call ChangeMode(Me, False)

        cmdInspect.Enabled = True
        frmSearch.Enabled = False
        grdOrder.Enabled = False
        pnlEdit.Enabled = True
        txtOrderID.Locked = False
        txtOrderID.SetFocus
        txtOrderID = Left(MakeDate(DF_SHORT, Now), 4)
        txtOrderID.SelStart = 5
    Case ID_UPDATE
        If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub
'        If grdOrder.TextMatrix(grdOrder.Row, 25) = "0" Then
'            MsgBox "수동작업으로 출고작업한건입니다." & vbCrLf & "검사자동패킹화면에서 수정할 수 없습니다", vbInformation
'            Exit Sub
'        End If

        m_sOperate = ID_UPDATE
        m_sOrderID = MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 1), OM_REDUCE)
        m_nOutSeq = grdOrder.TextMatrix(grdOrder.Row, 13)

        Call ChangeMode(Me, False)

        Call MakeColorGridCombo
        
        With grdPacking
            .Editable = flexEDKbdMouse
            .AddItem ""
            .Row = .Rows - 1
            .Col = .FixedCols
        End With
        
        cmdInspect.Enabled = True
        frmSearch.Enabled = False
        grdOrder.Enabled = False
        pnlEdit.Enabled = True
        txtOrderID.Locked = True
        cmdFind(3).Enabled = False
        txtOutCustom.SetFocus
    Case ID_DELETE
        If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub

        If Not QuestionBox(LoadResString(201)) Then Exit Sub

        If DeleteData() Then Call FillGridOrder
    Case ID_SAVE
        If SaveData() Then
            Call ChangeMode(Me, True)
            pnlEdit.Enabled = False
            grdPacking.Editable = flexEDNone

            frmSearch.Enabled = True
            grdOrder.Enabled = True
            cmdInspect.Enabled = False
            
            Call FillGridOrder
            Call FindOrder
        End If
    Case ID_CANCEL
        Call ChangeMode(Me, True)
        pnlEdit.Enabled = False
        grdPacking.Editable = flexEDNone
        
        frmSearch.Enabled = True
        grdOrder.Enabled = True
        cmdInspect.Enabled = False

        grdOrder.SetFocus

        If grdOrder.Rows > grdOrder.FixedRows Then
            Call FillGridOrder
            Call FindOrder
        Else
            Call ClearData
        End If
    End Select
End Sub

Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
                
    On Error GoTo ErrHandler
    
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_ORDER, , False, txtOrderID)
        
        Set oOutware = New PlusLib2.COutWare
        oOutware.Connection = g_adoCon
        
        Set rs = oOutware.GetOrderOne(txtOrderID.Tag)
        Set oOutware = Nothing
        If Not rs.EOF Then
            txtOrderID = txtOrderID.Tag
            txtOrder = rs!OrderNo
            txtCustom = rs!kCustom
            txtCustom.Tag = rs!CustomID
            txtArticle = rs!Article
            txtArticle.Tag = rs!ArticleID
            cboWork.ListIndex = FindComboBox(cboWork, CLng(rs!WorkID))
            txtOrderQty = Format(rs!OrderQty, "#,##0")
            txtOutRealQty(0) = 0
            txtOutRealQty(1) = 0
            txtLeftQty(0) = Format(CheckNum(rs!OrderQty) - CheckNum(rs!OutSumQty), "#,##0")
            txtLeftQty(1) = Format(CheckNum(rs!OrderQty) - CheckNum(rs!OutSumQty), "#,##0")
            txtLeftQty(0).Tag = rs!ChunkRate
            txtOutCustom.Tag = rs!UnitClss
            txtOutCustom = CheckNull(rs!DvlyPlace)
            txtOutSumQty(0) = Format(CheckNum(rs!OutSumQty), "#,##0")
            txtOutSumQty(1) = Format(CheckNum(rs!OutSumQty), "#,##0")
            txtRemark = rs!OutTelNO
            
            rs.Close
            Set rs = Nothing
            Call MakeColorGridCombo
        Else
            txtOrder = ""
            txtArticle = ""
            txtArticle.Tag = ""
            txtCustom = ""
            txtCustom.Tag = ""
            txtOrderQty = 0
            txtOutRealQty(0) = 0
            txtOutRealQty(1) = 0
            cboWork.ListIndex = -1
            txtLeftQty(0) = 0
            txtLeftQty(1) = 0
            txtLeftQty(0).Tag = 0
            txtOutCustom = ""
            txtOutCustom.Tag = ""
            txtOutSumQty(0) = 0
            txtOutSumQty(1) = 0
            txtRemark = ""
        End If
    End If
    
    Exit Sub
ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareIns.txtOrderID_KeyPress", Err.Description)
End Sub

Private Sub cboOutClss_Click()
'    If grdPacking.Rows = grdPacking.FixedRows Then Exit Sub
'    Call CalcRollSum
End Sub

Private Sub grdPacking_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdPacking
        If Row < .FixedRows Or Col < .FixedCols Then Exit Sub

        If Col >= 13 Then Cancel = True
    End With
End Sub

Private Sub grdPacking_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdPacking
        If KeyCode = vbKeyDown Then
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1
                .Col = 3
                .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
                .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 2, 2)
    
                ' ColorID 복사
               .TextMatrix(.Rows - 1, 15) = .TextMatrix(.Rows - 2, 15)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Row <> 1 Then
                .RemoveItem .Row
                Call CalcRollSum
            End If
        ElseIf KeyCode = vbKeyInsert Then
                .AddItem "", .Row + 1
                .TextMatrix(.Row + 1, 0) = .TextMatrix(.Row, 0)
                .TextMatrix(.Row + 1, 1) = .TextMatrix(.Row, 1)
                .TextMatrix(.Row + 1, 2) = .TextMatrix(.Row, 2)

                ' ColorID 복사
               .TextMatrix(.Row + 1, 15) = .TextMatrix(.Row, 15)
               .Select .Row + 1, 3
        End If
    End With
End Sub

Private Sub grdPacking_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i%, nRoll%, nQty#, iSign%
    Dim nRow%, nCol%
    
    On Error Resume Next

    With grdPacking
        If Row < .FixedRows Or Col < .FixedCols Then Exit Sub


        If Col >= 3 Then
            If Len(.EditText) > 0 And CheckNum(.EditText) <> 0 Then
                iSign = InStr(.EditText, "*")
                If iSign > 0 Then
                    nQty = Left(.EditText, iSign - 1)
                    nRoll = Right(.EditText, Len(.EditText) - iSign)
                Else
                    nQty = CSng(.EditText)
                    nRoll = 1
                End If

                .EditText = CStr(nQty)
                nCol = Col
                For i = 0 To nRoll - 1
                    If nCol = 13 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
                        .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
                        .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 2, 2)
            
                        ' ColorID 복사
                       .TextMatrix(.Rows - 1, 15) = .TextMatrix(.Rows - 2, 15)
                       
                       nRow = nRow + 1
                       nCol = 3
                       
                       .TextMatrix(Row + nRow, nCol) = CStr(nQty)
                    Else
                        .TextMatrix(Row + nRow, nCol) = CStr(nQty)

                    End If
                    nCol = nCol + 1
                Next i
            Else
                nRoll = 1
            End If

            .Col = Col + nRoll

            If .Col = 13 Then
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .Col = 1
                ElseIf .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    .Col = 1
    
                     '레코드가 추가되면 바로위 레코드 내역 복사
                    .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
                    .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
                    .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 2, 2)
    
                     'ColorID 복사
                   .TextMatrix(.Rows - 1, 15) = .TextMatrix(.Rows - 2, 15)
                End If
            End If
        ElseIf Col = 1 Then
            .TextMatrix(Row, 15) = .Cell(flexcpText, Row, 1)
            If Row = .Rows - 2 Then
                .TextMatrix(Row + 1, Col) = .Cell(flexcpText, Row, Col)
                .TextMatrix(Row + 1, 15) = .TextMatrix(Row, 15)
            End If

            .Col = Col + 1
        ElseIf Col = 0 Or Col = 2 Then
            If Row = .Rows - 2 Then .TextMatrix(Row + 1, Col) = .Cell(flexcpText, Row, Col)

            .Col = Col + 1
        End If

        Call CalcRollSum
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        grdOrder.ColWidth(1) = 0
        grdOrder.ColWidth(2) = 1350
        chkSearch(3).Caption = "Order No"
    Else
        grdOrder.ColWidth(1) = 1350
        grdOrder.ColWidth(2) = 0
        chkSearch(3).Caption = "관리번호"
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim oOutware As PlusLib2.COutWare
    Dim sPrinter As String
    
    On Error GoTo ErrHandler
    
    If grdOrder.Rows = grdOrder.FixedRows Then Exit Sub

    sPrinter = Printer.DeviceName
    If frmPrinter.SelectPrinter(sPrinter) Then
        Set oOutware = New PlusLib2.COutWare
        oOutware.Connection = g_adoCon
        oOutware.UserName = g_sUserName
    
        Call oOutware.UpdateTranNo(txtOrderID.Tag, CInt(txtOrder.Tag), m_sTranNo, m_nTranSeq)
               
        Call PrintLinePrint
        Call ReturnPrinter(sPrinter)
    End If
    Set oOutware = Nothing

    Exit Sub

ErrHandler:
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%

    With grdOrder
        .Cols = 28
        Call SetVSFlexGrid(grdOrder)

        .Redraw = flexRDNone

        .TextArray(0) = " "
        .TextArray(1) = "관리번호":     .ColWidth(1) = 1350:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order NO":     .ColWidth(2) = 0:       .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "거래처명":     .ColWidth(3) = 1500:    .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "출고일자":     .ColWidth(4) = 1000:    .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "품명":         .ColWidth(5) = 0
        .TextArray(6) = "출고처":       .ColWidth(6) = 0
        .TextArray(7) = "출고구분":     .ColWidth(7) = 0
        .TextArray(8) = "수주량":     .ColWidth(8) = 0
        .TextArray(9) = "소요량":     .ColWidth(9) = 0
        .TextArray(10) = "잔량":      .ColWidth(10) = 0
        .TextArray(11) = "전화번호":    .ColWidth(11) = 0
        .TextArray(12) = "비고":    .ColWidth(12) = 0
        .TextArray(13) = "OutSeq":      .ColWidth(13) = 0
        .TextArray(14) = "출고일자":    .ColWidth(14) = 0
        .TextArray(15) = "단위":      .ColWidth(15) = 0
        .TextArray(16) = "누계출고":      .ColWidth(16) = 0
        .TextArray(17) = "원단폭":      .ColWidth(17) = 0
        .TextArray(18) = "가공구분코드":        .ColWidth(18) = 0
        .TextArray(19) = "가공구분":    .ColWidth(19) = 0
        .TextArray(20) = "거래처코드":  .ColWidth(20) = 0
        .TextArray(21) = "품명코드":    .ColWidth(21) = 0
        .TextArray(22) = "축율":        .ColWidth(22) = 0
        .TextArray(23) = "비고":        .ColWidth(23) = 0
        .TextArray(24) = "출고량":      .ColWidth(24) = 1000
        .TextArray(25) = "출고방식":    .ColWidth(25) = 0
        .TextArray(26) = "출고일자":     .ColWidth(26) = 1100:    .ColAlignment(26) = flexAlignCenterCenter
        .TextArray(27) = "송장번호":     .ColWidth(27) = 1000:    .ColAlignment(27) = flexAlignCenterCenter
        
        .ColFormat(24) = "#,###"
        .Redraw = flexRDDirect
    End With

    With grdTotal
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 4
        .FixedRows = 0
        .FixedCols = 0
        .RowHeightMin = 400
        
        .TextArray(0) = "합 계":    .ColWidth(0) = 3200:        .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "":         .ColWidth(1) = 1100:        .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "":         .ColWidth(2) = 1100:        .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "":         .ColWidth(3) = 1000:        .ColAlignment(3) = flexAlignRightCenter
        
        .ColFormat(1) = "#,###"
        .ColFormat(2) = "#,###"
        
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 3) = COLOR_GRIDROW
        
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .Redraw = flexRDDirect
    End With


    With grdPacking
        .Redraw = flexRDNone

        .Rows = 1
        .RowHeight(0) = 450

        .ScrollBars = flexScrollBarBoth

        .Cols = 26
        .FixedCols = 0

        Call SetVSFlexGrid(grdPacking)

        .TextArray(0) = "Box No":       .ColWidth(0) = 800:             .ColAlignment(0) = flexAlignLeftCenter
        .TextArray(1) = "색상명":       .ColWidth(1) = 2400:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "Lot NO":       .ColWidth(2) = 750:             .ColAlignment(2) = flexAlignCenterCenter
        
        For i = 0 To 9
            .TextArray(i + 3) = CStr(i + 1):  .ColWidth(i + 3) = 1000:     .ColAlignment(i + 3) = flexAlignLeftCenter
        Next i
        .TextArray(13) = "절수":        .ColWidth(13) = 700:             .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "수량":        .ColWidth(14) = 900:             .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "OrderSeq":    .ColWidth(15) = 0

        For i = 16 To 25
             .ColWidth(i) = 0
        Next i

        .ColFormat(13) = "#,###"
        .ColFormat(14) = "#,###"

        .SelectionMode = flexSelectionFree
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .ScrollTrack = True
        .FillStyle = flexFillRepeat
        .ExplorerBar = flexExNone
        .MousePointer = flexCustom

        .RowHeightMin = 350
        .WordWrap = True

        .ExtendLastCol = True
        .ColHidden(0) = True
        
        .Redraw = flexRDDirect
    End With
    
    With grdSum
        .Redraw = flexRDNone
        
        Call SetVSFlexGrid(grdSum)
        
        .RowHeight(0) = 300
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 3
        .Rows = 1
        
        .TextArray(0) = "합계":        .ColWidth(0) = 13200:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "0":        .ColWidth(1) = 700:   .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "0":        .ColWidth(2) = 900:   .ColAlignment(2) = flexAlignRightCenter
        
        .ColFormat(1) = "#,##0"
        .ColFormat(2) = "#,##0"
        
        .HighLight = flexHighlightNever
        .Redraw = flexRDDirect
    End With

    With grdColor
        .Redraw = flexRDNone
        .Cols = 4
        
        Call SetVSFlexGrid(grdColor)
        
        .TextArray(0) = "":         .ColWidth(0) = 500:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "선택":         .ColWidth(1) = 450:   .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "색상순위":     .ColWidth(2) = 0:     .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "색상명":       .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignLeftCenter
        
        .ColDataType(1) = flexDTBoolean
        .Redraw = flexRDDirect
    End With
    
    With grdRoll
        .Redraw = flexRDNone
        .Cols = 11
        
        Call SetVSFlexGrid(grdRoll)

        .TextArray(0) = " ":             .ColWidth(0) = 450
        .TextArray(1) = "선택":         .ColWidth(1) = 450:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":       .ColWidth(2) = 1400:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "LOT":          .ColWidth(3) = 500:     .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "절번호":      .ColWidth(4) = 650:     .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "수량":         .ColWidth(5) = 500:     .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "LOSS":         .ColWidth(6) = 600:       .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "색상순위":      .ColWidth(7) = 0
        .TextArray(8) = "합불":         .ColWidth(8) = 0
        .TextArray(9) = "RollSeq":      .ColWidth(9) = 0
        .TextArray(10) = "구분":        .ColWidth(10) = 450

        .ColDataType(1) = flexDTBoolean

        .MergeCells = flexMergeFree
        .MergeCol(2) = True
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With

    With grdOut
        .Redraw = flexRDNone
        .Cols = 10
        
        Call SetVSFlexGrid(grdOut)

        .TextArray(0) = " ":             .ColWidth(0) = 450
        .TextArray(1) = "선택":         .ColWidth(1) = 450:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":       .ColWidth(2) = 1400:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "LOT":          .ColWidth(3) = 500:     .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "절번호":      .ColWidth(4) = 650:     .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "수량":         .ColWidth(5) = 500:     .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "LOSS":         .ColWidth(6) = 500:       .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "색상순위":      .ColWidth(7) = 0
        .TextArray(8) = "합불":         .ColWidth(8) = 0
        .TextArray(9) = "RollSeq":      .ColWidth(9) = 0

        .ColDataType(1) = flexDTBoolean

        .MergeCells = flexMergeFree
        .MergeCol(2) = True
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With

    With grdRollSum
        .Redraw = flexRDNone
        .Cols = 7
        
        Call SetVSFlexGrid(grdRollSum)
        
        .TextArray(0) = "":             .ColWidth(0) = 500:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "색상순위":     .ColWidth(1) = 0:     .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "색상명":       .ColWidth(2) = 1000:  .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "수주량":       .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "출고량":       .ColWidth(4) = 0:     .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "잔량":         .ColWidth(5) = 1000:  .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "출고가능량":       .ColWidth(6) = 1000:  .ColAlignment(6) = flexAlignRightCenter
        
        .ColFormat(3) = "#,###"
        .ColFormat(4) = "#,###"
        .ColFormat(5) = "#,###"
        .ColFormat(6) = "#,###"
        
        .Redraw = flexRDDirect
    End With

    With grdOutSum
        .Redraw = flexRDNone
        .Cols = 7
        
        Call SetVSFlexGrid(grdOutSum)
        
        .TextArray(0) = "":             .ColWidth(0) = 500:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "색상순위":     .ColWidth(1) = 0:     .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "색상명":       .ColWidth(2) = 1000:  .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "수주량":       .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "출고량":       .ColWidth(4) = 0:     .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "잔량":         .ColWidth(5) = 1000:  .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "출고예정량":       .ColWidth(6) = 1000:  .ColAlignment(6) = flexAlignRightCenter
        
        .ColFormat(3) = "#,###"
        .ColFormat(4) = "#,###"
        .ColFormat(5) = "#,###"
        .ColFormat(6) = "#,###"
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub FillGridOrder()
    Dim oOutware As PlusLib2.COutWare
    Dim rs       As Recordset
    Dim i%, iNowRow%, nTRoll#, nTQty#, nRecCnt As Integer

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon

    Set rs = oOutware.GetOutware(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                        IIf(chkSearch(1), 1, 0), txtSearch(1).Tag, _
                        IIf(chkSearch(2), 1, 0), txtSearch(2).Tag, _
                        IIf(chkSearch(3), IIf(optOrder(0), 2, 1), 0), txtSearch(3))
    Set oOutware = Nothing
    
    nRecCnt = rs.RecordCount

    With grdOrder
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            DoEvents

            .AddItem CStr(i) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                rs!kCustom & vbTab & MakeDate(DF_LONG, rs!OutDate) & vbTab & rs!Article & vbTab & rs!OutCustom & vbTab & _
                rs!OutClss & vbTab & rs!OrderQty & vbTab & rs!OutRealQty & vbTab & rs!OrderQty - rs!OutSumQty & vbTab & _
                "" & vbTab & "" & vbTab & rs!OutSeq & vbTab & rs!OutDate & vbTab & rs!UnitClss & vbTab & rs!OutSumQty & vbTab & _
                rs!WorkWidth & vbTab & rs!WorkID & vbTab & rs!WorkName & vbTab & rs!CustomID & vbTab & _
                rs!ArticleID & vbTab & rs!ChunkRate & vbTab & rs!Remark & vbTab & rs!OutQty & vbTab & rs!OutType & vbTab & _
                IIf(Trim(rs!ResultDate) = "", "", MakeDate(DF_LONG, rs!ResultDate)) & vbTab & rs!TranNo & "-" & rs!TranSeq

            nTRoll = nTRoll + rs!OutRoll
            nTQty = nTQty + rs!OutQtyY

            If rs!OutType = "1" Then
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = &HC0FFFF
            End If

            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        .Redraw = flexRDDirect
        .SetFocus
    End With

    m_bloading = False
    Screen.MousePointer = vbDefault

    With grdTotal
        .TextArray(1) = Format(nRecCnt, "##,##0 건")
        .TextArray(2) = Format(nTRoll, "##,##0 절")
        .TextArray(3) = Format(nTQty, "#,###,##0 YDS")
    End With

    Call ShowData
    Exit Sub

ErrHandler:
    m_bloading = False
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub ShowData()
    If m_bloading Then Exit Sub

    On Error GoTo ErrHandler

    With grdOrder
        If .Rows = .FixedRows Then
            Call ClearData
            Exit Sub
        End If

        txtOrderID = MakeOrderID(.TextMatrix(.Row, 1), OM_REDUCE)  ' 접수번호
        txtOrderID.Tag = MakeOrderID(.TextMatrix(.Row, 1), OM_REDUCE)  ' 접수번호
        txtOrder = .TextMatrix(.Row, 2)      ' Order No
        txtOrder.Tag = .TextMatrix(.Row, 13) ' OutSeq
        txtCustom = .TextMatrix(.Row, 3)      ' 거래처
        txtCustom.Tag = .TextMatrix(.Row, 20)  '거래처 코드
        txtArticle = .TextMatrix(.Row, 5)       '품명
        txtArticle.Tag = .TextMatrix(.Row, 21)  '품명 코드
        txtOutCustom = .TextMatrix(.Row, 6)      '출고처
        txtOutCustom.Tag = .TextMatrix(.Row, 15)  '단위
        txtOrderQty = Format(.TextMatrix(.Row, 8), "#,##0")      '수주량
        txtUnitClss = IIf(.TextMatrix(.Row, 15) = "0", "YDS", "MTS")
        txtOutRealQty(0) = Format(.TextMatrix(.Row, 9), "#,##0")     '소요량
        txtOutRealQty(0).Tag = Format(.TextMatrix(.Row, 24), "#,##0")     '출고량
        txtOutRealQty(1) = Format(.TextMatrix(.Row, 9), "#,##0")     '소요량
        txtOutSumQty(0) = Format(.TextMatrix(.Row, 16), "#,##0")     '누계출고
        txtOutSumQty(1) = Format(.TextMatrix(.Row, 16), "#,##0")     '누계출고
        txtLeftQty(0) = Format(.TextMatrix(.Row, 10), "#,##0")      '잔량
        txtLeftQty(1) = Format(.TextMatrix(.Row, 10), "#,##0")      '잔량
        txtLeftQty(0).Tag = .TextMatrix(.Row, 22)  '축율
        txtRemark = .TextMatrix(.Row, 23) '비고사항
        
        cboWork.ListIndex = FindComboBox(cboWork, CLng(.TextMatrix(.Row, 18))) '가공구분
        cboOutClss.ListIndex = FindComboBox(cboOutClss, CLng(.TextMatrix(.Row, 7)))  ' 출고구분
        dtpOutDate = MakeDate(DF_LONG, .TextMatrix(.Row, 4))   '작성일자
        
        chkSearch(4).Value = IIf(.TextMatrix(.Row, 26) <> "", vbChecked, vbUnchecked)
        dtpResultDate = IIf(.TextMatrix(.Row, 26) <> "", MakeDate(DF_LONG, .TextMatrix(.Row, 26)), Now) '출고일자
    End With

    Call FillGridPacking

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmOutwareIns.ShowData", Err.Description)

    Resume Next
End Sub

Private Sub FillGridPacking()
    Dim oOutware As PlusLib2.COutWare
    Dim rs       As Recordset
    Dim i%, j%, iCol%, nPoint%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon

    Set rs = oOutware.GetOutwareSub(txtOrderID.Tag, CInt(txtOrder.Tag))
    Set oOutware = Nothing

    With grdPacking
        .Redraw = flexRDDirect
        .Rows = .FixedRows

        For i = 1 To rs.RecordCount
            If rs!OrderSeq <> .TextMatrix(.Rows - 1, 15) Or rs!LotNo <> .TextMatrix(.Rows - 1, 2) _
                Or CStr(rs!BoxNo) <> .TextMatrix(.Rows - 1, 0) Or Len(.TextMatrix(.Rows - 1, 13)) > 0 Or nPoint > 9 Then
                .AddItem CStr(rs!BoxNo) & vbTab & rs!Color & " " & rs!DesignNO & vbTab & rs!LotNo
                .TextMatrix(.Rows - 1, 0) = rs!BoxNo
                .TextMatrix(.Rows - 1, 1) = rs!Color & " " & rs!DesignNO
                .TextMatrix(.Rows - 1, 2) = rs!LotNo
                .TextMatrix(.Rows - 1, 15) = rs!OrderSeq
                nPoint = 0
            End If

            For j = 0 To 9
                If Len(.TextMatrix(.Rows - 1, j + 3)) = 0 Then
                    .TextMatrix(.Rows - 1, j + 3) = rs!OutQty
                    .TextMatrix(.Rows - 1, j + 16) = IIf(rs!RollSeq = 0, "", rs!RollSeq)
                    nPoint = nPoint + 1
                    Exit For
                End If
            Next j

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        Call CalcRollSum

        .Redraw = flexRDDirect
    End With

    m_bloading = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    m_bloading = False
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, "frmOutwareIns.FillGridPacking", Err.Description)
End Sub

Private Sub ClearData()
    txtOrderID = ""
    txtOrderID.Tag = ""
    txtOrder = ""
    txtOrder.Tag = ""
    txtCustom = ""
    txtArticle = ""
    txtArticle.Tag = ""
    txtOutCustom = ""
    txtOutCustom.Tag = ""
    txtOrderQty = 0
    txtOutRealQty(0) = 0
    txtOutRealQty(0).Tag = 0
    txtOutRealQty(1) = 0
    txtLeftQty(0) = 0
    txtLeftQty(1) = 0
    txtLeftQty(0).Tag = ""
    txtOutSumQty(0) = 0
    txtOutSumQty(1) = 0
    txtRemark = ""
    grdPacking.Rows = grdPacking.FixedRows
    cboOutClss.ListIndex = 0
    dtpOutDate = Now
    chkSearch(4).Value = vbChecked
    grdSum.TextMatrix(0, 1) = 0
    grdSum.TextMatrix(0, 2) = 0
End Sub

Private Function SaveData() As Boolean
    Dim ow       As PlusLib2.TOUTWARE
    Dim owSub()  As PlusLib2.TOUTWARESUB
    Dim oOutware As PlusLib2.COutWare
    Dim i%, j%, iSub%, nSeq%, nOrderSeq%
    
    SaveData = False

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    With ow
        ow.OrderID = txtOrderID.Tag
        If m_sOperate = ID_UPDATE Then
            ow.OutSeq = txtOrder.Tag
        End If
        ow.OutClss = cboOutClss.ItemData(cboOutClss.ListIndex)
        ow.WorkID = Format(cboWork.ItemData(cboWork.ListIndex), "0000")
        ow.ExchRate = 0
        ow.UnitPrice = 0
        ow.OutCustom = Trim(txtOutCustom)
        ow.LossRate = 0
        ow.LossQty = 0
        ow.OutDate = MakeDate(DF_SHORT, dtpOutDate)
        ow.ResultDate = ow.OutDate
        ow.OutTime = Format(time, "HHMM")
        ow.BoOutClss = ""
        ow.LoadTime = Format(time, "HHMM")
        ow.OutRoll = grdSum.TextMatrix(0, 1)
        ow.OutQty = grdSum.TextMatrix(0, 2)
        ow.OutRealQty = CheckNum(txtOutRealQty(1))
        ow.OutType = "1"
        ow.Remark = txtRemark
    End With

    With grdPacking
        ReDim owSub(Abs(ow.OutRoll) - 1)

        iSub = 0
        nSeq = -1
        For i = .FixedRows To .Rows - 1
            For j = 3 To 12
                If CheckNum(.TextMatrix(i, j)) <> 0 Then
                    owSub(iSub).OrderID = ow.OrderID
                    owSub(iSub).OutSubSeq = iSub + 1
                    owSub(iSub).OrderSeq = .TextMatrix(i, 15)
                    owSub(iSub).BoxNo = CheckNum(.TextMatrix(i, 0))
                    owSub(iSub).LotNo = .TextMatrix(i, 2)
                    owSub(iSub).OutQty = CSng(.TextMatrix(i, j))

                    If IsNumeric(.TextMatrix(i, j + 13)) Then
                        owSub(iSub).RollSeq = CInt(.TextMatrix(i, j + 13))
                    End If

                    iSub = iSub + 1
                    nOrderSeq = .TextMatrix(i, 15)
                End If
            Next j
        Next i
    End With

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    oOutware.UserName = g_sUserName

    If m_sOperate = ID_ADDNEW Then
        SaveData = oOutware.AddNewOutware(ow, owSub)
        m_nOutSeq = ow.OutSeq
    ElseIf m_sOperate = ID_UPDATE Then
        SaveData = oOutware.UpdateOutware(ow, owSub)
    End If

    Set oOutware = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Set oOutware = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, "frmOutwareIns.SaveData", Err.Description)
End Function

Private Function DeleteData() As Boolean
    Dim oOutware As PlusLib2.COutWare

    On Error GoTo ErrHandler

    DeleteData = False

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    oOutware.UserName = g_sUserName

    DeleteData = oOutware.DeleteOutware(txtOrderID.Tag, CInt(txtOrder.Tag))

    Set oOutware = Nothing

    Exit Function

ErrHandler:
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, "frmOutwareIns.DeleteData", Err.Description)
End Function

Private Sub CalcRollSum()
    Dim i%, j%, nRoll%, nQty#, sKey$
    Dim iSumRow&, nRollSign%, nPoint%
    Dim nTRoll%, nTQty#
    
    With grdPacking
        .Redraw = flexRDNone

        nRollSign = IIf(cboOutClss.ItemData(cboOutClss.ListIndex) = 2, -1, 1)

        For i = .FixedRows To .Rows
            If i = .Rows Then
                .TextMatrix(iSumRow, 13) = CStr(nRoll * nRollSign)
                .TextMatrix(iSumRow, 14) = CStr(nQty * nRollSign)
                
                grdSum.TextMatrix(0, 1) = CStr(nTRoll * nRollSign)
                grdSum.TextMatrix(0, 2) = CStr(nTQty * nRollSign)
                
                Exit For
            End If

            If sKey <> Format(.TextMatrix(i, 15), "000") & .TextMatrix(i, 0) & .TextMatrix(i, 2) Then
                If i > .FixedRows Then
                    .TextMatrix(iSumRow, 13) = CStr(nRoll * nRollSign)
                    .TextMatrix(iSumRow, 14) = nQty
                End If

                If pnlEdit.Enabled And i = .Rows Then Exit For

                sKey = Format(.TextMatrix(i, 15), "000") & .TextMatrix(i, 0) & .TextMatrix(i, 2)
                iSumRow = i

                nRoll = 0
                nQty = 0
            End If

            For j = 3 To 12
                If CheckNum(.TextMatrix(i, j)) <> 0 Then
                    nRoll = nRoll + 1
                    nQty = nQty + CSng(.TextMatrix(i, j))
                    
                    nTRoll = nTRoll + 1
                    nTQty = nTQty + CSng(.TextMatrix(i, j))
                End If
            Next j
            .TextMatrix(i, 13) = ""
            .TextMatrix(i, 14) = ""
        Next i

        .Redraw = flexRDDirect
    End With
    If Not m_bloading Then
        If txtOutCustom.Tag = "0" Then
            txtOutRealQty(1) = Format(MakeNeedQty(grdSum.TextMatrix(0, 2), txtLeftQty(0).Tag), "#,##0")
        Else
            txtOutRealQty(1) = Format(MakeNeedQty(CLng((grdSum.TextMatrix(0, 2) / 0.9144)), txtLeftQty(0).Tag), "#,##0")
        End If
    
'        txtOutRealQty(1) = Format(MakeNeedQty(grdSum.TextMatrix(0, 2), txtLeftQty(0).Tag), "#,##0")
    End If
    txtOutSumQty(1) = Format(txtOutSumQty(0) - txtOutRealQty(0).Tag + grdSum.TextMatrix(0, 2), "#,##0")
    txtLeftQty(1) = Format(txtOrderQty - txtOutSumQty(1), "#,##0")
    
End Sub

Private Sub MakeColorGridCombo()
    Dim oOrder As PlusLib2.COrder
    Dim rs     As Recordset
    Dim i%, sCombo$

    On Error GoTo ErrHandler


    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon

    Set rs = oOrder.GetOrderSub(txtOrderID.Tag)
    Set oOrder = Nothing

    Do Until rs.EOF
        sCombo = sCombo & "#" & rs!OrderSeq & ";" & rs!Color & " " & rs!DesignNO & vbTab & Format(rs!ColorQty, "#,###") & "|"

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    grdPacking.ColComboList(1) = sCombo

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing

    Call ErrorBox(Err.Number, "frmOutwareIns.MakeColorGridCombo", Err.Description)
End Sub

'Private Sub MakeExcelPacking()
'    Dim oExcel      As Excel.Application
'    Dim oExcelBook  As Excel.Workbook
'    Dim oExcelSheet As Excel.Worksheet
'    Dim oFs         As FileSystemObject
'    Dim i%, j%, nRow%, nLimitRow%, nBaseRow%, nCurRow%, nPage%, sReport$
'    Dim nOrderSeq%, sLotNo$
'    Dim sUnit$, nColorRoll%, nColorQty#
'
'    On Error GoTo ErrHandler
'
'    Screen.MousePointer = vbHourglass
'
'    Set oExcel = New Excel.Application
'    Set oExcelBook = oExcel.Workbooks.Open(App.Path & REPORTFILE)
'
'    oExcel.WindowState = xlMaximized
'    oExcel.Application.Visible = True
'
'    sUnit = IIf(grdOrder.TextMatrix(grdOrder.Row, 18) = "0", "Y", "M")
'    With oExcel
'        ' Make Sum
'        .Worksheets("Form").Activate
'
'        .Cells(5, 4) = MakeDate(DF_FULL, Date)
'        .Cells(12, 1) = Trim(txtArticle)
'        .Cells(12, 5) = Trim(grdOrder.TextMatrix(grdOrder.Row, 19))
'        .Cells(12, 10) = Trim(grdOrder.TextMatrix(grdOrder.Row, 16))
'        .Cells(12, 12) = Trim(txtOrder)
'        .Cells(12, 17) = Format(grdOrder.TextMatrix(grdOrder.Row, 17), "#,###") & sUnit
'        .Cells(12, 21) = Format(grdSum.TextMatrix(0, 1), "#,###")
'        .Cells(12, 24) = Format(grdSum.TextMatrix(0, 2), "#,###") & sUnit
'        If Len(m_sTranNo) > 0 Then
'            .Cells(5, 13) = m_sTranNo & "-" & m_nTranSeq
'        Else
'            .Cells(5, 13) = ""
'        End If
'
'        .Worksheets("Print").Activate
'
'        nPage = 1
'        nBaseRow = GetExcelRollBaseRow(nPage)
'        Call InsertExcelForm(oExcel, nPage)
'        nCurRow = nBaseRow + 15
'        For i = grdPacking.FixedRows To grdPacking.Rows - 1
'            If nCurRow + nRow > nBaseRow + 37 Then
'                nPage = nPage + 1
'                nBaseRow = GetExcelRollBaseRow(nPage)
'                Call InsertExcelForm(oExcel, nPage)
'                nCurRow = nBaseRow + 15
'                nRow = 0
'            End If
'
'            If nOrderSeq <> grdPacking.TextMatrix(i, 15) Then
'                If i > grdPacking.FixedRows Then
'                    nRow = nRow + 1
'                    .Cells(nCurRow + nRow - 1, 3) = "COLOR 계 : "
'                    .Cells(nCurRow + nRow - 1, 6) = Format(nColorRoll, "#,###")
'                    .Cells(nCurRow + nRow - 1, 8) = Format(nColorQty, "#,###") & sUnit
'                End If
'                If nCurRow + nRow > nBaseRow + 37 Then
'                    nPage = nPage + 1
'                    nBaseRow = GetExcelRollBaseRow(nPage)
'                    Call InsertExcelForm(oExcel, nPage)
'                    nCurRow = nBaseRow + 15
'                    nRow = 0
'                End If
'                .Cells(nCurRow + nRow, 1) = grdPacking.TextMatrix(i, 15)
'                .Cells(nCurRow + nRow, 3) = Trim(grdPacking.TextMatrix(i, 1))
'                .Cells(nCurRow + nRow, 5) = grdPacking.TextMatrix(i, 2)
'                .Cells(nCurRow + nRow, 6) = Format(grdPacking.TextMatrix(i, 13), "#,###")
'                .Cells(nCurRow + nRow, 8) = Format(grdPacking.TextMatrix(i, 14), "#,###") & sUnit
'
'                nColorRoll = CheckNum(grdPacking.TextMatrix(i, 13))
'                nColorQty = CheckNum(grdPacking.TextMatrix(i, 14))
'            ElseIf nOrderSeq = grdPacking.TextMatrix(i, 15) And sLotNo <> grdPacking.TextMatrix(i, 2) Then
'                .Cells(nCurRow + nRow, 5) = grdPacking.TextMatrix(i, 2)
'                .Cells(nCurRow + nRow, 6) = Format(grdPacking.TextMatrix(i, 13), "#,###")
'                .Cells(nCurRow + nRow, 8) = Format(grdPacking.TextMatrix(i, 14), "#,###") & sUnit
'
'                nColorRoll = nColorRoll + CheckNum(grdPacking.TextMatrix(i, 13))
'                nColorQty = nColorQty + CheckNum(grdPacking.TextMatrix(i, 14))
'            End If
'
'           .Cells(nCurRow + nRow, 11) = grdPacking.TextMatrix(i, 3)
'           .Cells(nCurRow + nRow, 13) = grdPacking.TextMatrix(i, 4)
'           .Cells(nCurRow + nRow, 14) = grdPacking.TextMatrix(i, 5)
'           .Cells(nCurRow + nRow, 16) = grdPacking.TextMatrix(i, 6)
'           .Cells(nCurRow + nRow, 19) = grdPacking.TextMatrix(i, 7)
'           .Cells(nCurRow + nRow, 20) = grdPacking.TextMatrix(i, 8)
'           .Cells(nCurRow + nRow, 22) = grdPacking.TextMatrix(i, 9)
'           .Cells(nCurRow + nRow, 25) = grdPacking.TextMatrix(i, 10)
'           .Cells(nCurRow + nRow, 27) = grdPacking.TextMatrix(i, 11)
'           .Cells(nCurRow + nRow, 29) = grdPacking.TextMatrix(i, 12)
'
'            nOrderSeq = grdPacking.TextMatrix(i, 15)
'            sLotNo = grdPacking.TextMatrix(i, 2)
'            nRow = nRow + 1
'        Next i
'
'        If nCurRow + nRow > nBaseRow + 37 Then
'            nPage = nPage + 1
'            nBaseRow = GetExcelRollBaseRow(nPage)
'            Call InsertExcelForm(oExcel, nPage)
'            nCurRow = nBaseRow + 15
'        End If
'        .Cells(nCurRow + nRow, 3) = "COLOR 계 : "
'        .Cells(nCurRow + nRow, 6) = Format(nColorRoll, "#,###")
'        .Cells(nCurRow + nRow, 8) = Format(nColorQty, "#,###") & sUnit
'
'    End With
'
'    sReport = App.Path & REPORTFILE1
'
'    Set oFs = New FileSystemObject
'    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
'    Set oFs = Nothing
'
'    Call oExcelBook.SaveAs(sReport)
'
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
''    oExcel.ActiveWindow.SelectedSheets.PrintPreview
'
'    Screen.MousePointer = vbDefault
'
'    Set oExcelSheet = Nothing
'    Set oExcelBook = Nothing
'    Set oExcel = Nothing
'    Set oFs = Nothing
'
'    Exit Sub
'
'ErrHandler:
'    Screen.MousePointer = vbDefault
'
'    Set oExcelSheet = Nothing
'    Set oExcelBook = Nothing
'    Set oExcel = Nothing
'    Set oFs = Nothing
'
'    Call Err.Raise(Err.Number, "frmOutwareIns.MakeExcelPacking", Err.Description)
'End Sub
'
'Private Function InsertExcelForm(oExcel As Excel.Application, nPage As Integer)
'    Dim i%, nBaseRow%
'
'    nBaseRow = GetExcelRollBaseRow(nPage)
'    With oExcel
'        .Sheets("Form").Select
'
'        .Rows("1:" & CStr(EXCEL_ROLL_ROW)).Select
'        .Selection.Copy
'
'        .Sheets("Print").Select
'        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
'        .Selection.Insert Shift:=xlDown
'
'        .Cells(nBaseRow + 3, 27) = "PAGE : " & nPage
'    End With
'End Function
'
'Private Function GetExcelRollBaseRow(nPage)
'    GetExcelRollBaseRow = (nPage - 1) * EXCEL_ROLL_ROW
'End Function

Private Sub PrintLinePrint()
    Dim oCustom As PlusLib2.CCustom
    Dim rs As ADODB.Recordset
    Dim vCustom(5) As String
    Dim nOrderSeq%, sLotNo$, sUnit$, nColorRoll%, nColorQty#
    Dim i%, nRow%, nPage%, nLineFlag%, nLotCnt%, nLineCnt%
    Dim vXPOS(22) As Integer
    
    On Error GoTo ErrHandler
    
    vXPOS(0) = 93
    vXPOS(1) = 99
    vXPOS(2) = 105
    vXPOS(3) = 111
    vXPOS(4) = 118
    vXPOS(5) = 124
    vXPOS(6) = 131
    vXPOS(7) = 137
    vXPOS(8) = 144
    vXPOS(9) = 150
    vXPOS(10) = 157
    vXPOS(11) = 163
    vXPOS(12) = 169
    vXPOS(13) = 176
    vXPOS(14) = 182
    vXPOS(15) = 188
    vXPOS(16) = 194
    vXPOS(17) = 201
    vXPOS(18) = 207
    vXPOS(19) = 213
    vXPOS(20) = 219
    vXPOS(21) = 226
    vXPOS(22) = 232
    
    nLineCnt = 21
    
    
    Set oCustom = New PlusLib2.CCustom
    oCustom.Connection = g_adoCon
    Set rs = oCustom.GetCustomOne(grdOrder.TextMatrix(grdOrder.Row, 20))
    Set oCustom = Nothing
        
    If rs.EOF Then
        For i = 0 To 5
            vCustom(i) = ""
        Next i
    Else
        vCustom(0) = CheckNull(rs!kCustom)
        vCustom(1) = txtOutCustom
        vCustom(2) = MakeOrderID(txtOrderID, OM_EXPAND)
    End If
    Set rs = Nothing
    
    sUnit = IIf(grdOrder.TextMatrix(grdOrder.Row, 15) = "0", "Y", "M")
    
    Printer.Orientation = vbPRORPortrait
    Printer.ScaleMode = vbMillimeters
    
    nPage = 1
    Call PrintHead(nPage, vCustom)
    
    For i = grdPacking.FixedRows To grdPacking.Rows - 1
        If nRow > nLineCnt Then
            nPage = nPage + 1
            Printer.NewPage
            Call PrintHead(nPage, vCustom)
            nRow = 0
        End If
    
        If nOrderSeq <> grdPacking.TextMatrix(i, 15) Then
            If i > grdPacking.FixedRows Then
                If nLotCnt > 1 Then
                    nRow = nRow + 1
                    Call PrintDot(17, vXPOS(nRow - 1), "COLOR 계 :")
                    Call PrintDot(50, vXPOS(nRow - 1), MakeStrBySpace(Format(nColorRoll, "#,###"), 3, 0))
                    Call PrintDot(58, vXPOS(nRow - 1), MakeStrBySpace(Format(nColorQty, "#,###") & sUnit, 9, 0))
                End If
            End If
            If nRow > nLineCnt Then
                nPage = nPage + 1
                Printer.NewPage
                Call PrintHead(nPage, vCustom)
                nRow = 0
            End If
            
            If nLineFlag = 1 Then
                If nRow > 0 Then
                    Printer.Line (8, vXPOS(nRow - 1) + 4)-(174, vXPOS(nRow - 1) + 4)
                End If
            End If

'            Call PrintDot(8, vXPOS(nRow), grdPacking.TextMatrix(i, 15))
            Call PrintDot(8, vXPOS(nRow), Left(grdPacking.TextMatrix(i, 1), 16))
            Call PrintDot(42, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 2), 3, 0))
            Call PrintDot(50, vXPOS(nRow), MakeStrBySpace(Format(grdPacking.TextMatrix(i, 13), "#,###"), 3, 0))
            Call PrintDot(58, vXPOS(nRow), MakeStrBySpace(Format(grdPacking.TextMatrix(i, 14), "#,###") & sUnit, 9, 0))
            
            nColorRoll = CheckNum(grdPacking.TextMatrix(i, 13))
            nColorQty = CheckNum(grdPacking.TextMatrix(i, 14))
            nLineFlag = 1
            nLotCnt = 1
        ElseIf nOrderSeq = grdPacking.TextMatrix(i, 15) And sLotNo <> grdPacking.TextMatrix(i, 2) Then
'            Printer.Line (42, 95 + (nRow - 1) * 5)-(174, 95 + (nRow - 1) * 5)
            Call PrintDot(42, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 2), 3, 0))
            Call PrintDot(50, vXPOS(nRow), MakeStrBySpace(Format(grdPacking.TextMatrix(i, 13), "#,###"), 3, 0))
            Call PrintDot(58, vXPOS(nRow), MakeStrBySpace(Format(grdPacking.TextMatrix(i, 14), "#,###") & sUnit, 9, 0))
            
            nColorRoll = nColorRoll + CheckNum(grdPacking.TextMatrix(i, 13))
            nColorQty = nColorQty + CheckNum(grdPacking.TextMatrix(i, 14))
            nLotCnt = nLotCnt + 1
        End If
        
''        Call PrintDot(77, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 3), 4, 0))
''        Call PrintDot(87, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 4), 4, 0))
''        Call PrintDot(97, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 5), 4, 0))
''        Call PrintDot(107, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 6), 4, 0))
''        Call PrintDot(117, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 7), 4, 0))
''        Call PrintDot(127, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 8), 4, 0))
''        Call PrintDot(137, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 9), 4, 0))
''        Call PrintDot(147, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 10), 4, 0))
''        Call PrintDot(157, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 11), 4, 0))
''        Call PrintDot(167, 93 + nRow * 6, MakeStrBySpace(grdPacking.TextMatrix(i, 12), 4, 0))
       
        Call PrintDot(77, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 3), 4, 0))
        Call PrintDot(87, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 4), 4, 0))
        Call PrintDot(97, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 5), 4, 0))
        Call PrintDot(107, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 6), 4, 0))
        Call PrintDot(117, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 7), 4, 0))
        Call PrintDot(127, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 8), 4, 0))
        Call PrintDot(137, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 9), 4, 0))
        Call PrintDot(147, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 10), 4, 0))
        Call PrintDot(157, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 11), 4, 0))
        Call PrintDot(167, vXPOS(nRow), MakeStrBySpace(grdPacking.TextMatrix(i, 12), 4, 0))
       
        nOrderSeq = grdPacking.TextMatrix(i, 15)
        sLotNo = grdPacking.TextMatrix(i, 2)
        nRow = nRow + 1
    Next i
    
    If nLotCnt > 1 Then
        If nRow > nLineCnt Then
            nPage = nPage + 1
            Printer.NewPage
            Call PrintHead(nPage, vCustom)
            nRow = 0
        End If
    
        Call PrintDot(17, vXPOS(nRow), "COLOR 계 :")
        Call PrintDot(50, vXPOS(nRow), MakeStrBySpace(Format(nColorRoll, "#,###"), 3, 0))
        Call PrintDot(58, vXPOS(nRow), MakeStrBySpace(Format(nColorQty, "#,###") & sUnit, 9, 0))
        Printer.Line (8, vXPOS(nRow) + 4)-(174, vXPOS(nRow) + 4)
    Else
        Printer.Line (8, vXPOS(nRow - 1) + 4)-(174, vXPOS(nRow - 1) + 4)
    End If

    Printer.EndDoc
    Exit Sub
    
ErrHandler:
    Printer.KillDoc
    Set oCustom = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmOutwareIns.PrintLinePrint", Err.Description)
End Sub
Private Function PrintHead(nPage As Integer, vCustom() As String)
    
    Call PrintDot(160, 17, "PAGE : " & nPage)
    Call PrintDot(23, 32, Left(m_sTranNo, 4) & "-" & Right(m_sTranNo, 2) & "-" & m_nTranSeq) '일련번호
    Call PrintDot(68, 32, Left(grdOrder.TextMatrix(grdOrder.Row, 14), 4))
    Call PrintDot(85, 32, Mid(grdOrder.TextMatrix(grdOrder.Row, 14), 5, 2))
    Call PrintDot(97, 32, Right(grdOrder.TextMatrix(grdOrder.Row, 14), 2))
    
    Call PrintDot(114, 39, vCustom(0))  '거래처
    Call PrintDot(114, 45, vCustom(1))  '출고처
    Call PrintDot(114, 51, Trim(txtRemark))  '비고사항
    Call PrintDot(114, 58, Trim(txtOrder)) 'Order No

    Call PrintDot(125, 23, "담당")
    Call PrintDot(138, 23, "과장")
    Call PrintDot(152, 23, "이사")
    Call PrintDot(165, 23, "사장")
    
    Call PrintDot(8, 73, vCustom(2))  '관리번호
'    Call PrintDot(8, 73, Trim(txtOrder)) 'Order No
    Call PrintDot(43, 73, Left(Trim(txtArticle), 16)) '품명
    Call PrintDot(78, 73, Trim(grdOrder.TextMatrix(grdOrder.Row, 17)))  '규격
    Call PrintDot(92, 73, Trim(grdOrder.TextMatrix(grdOrder.Row, 19)))   '가공구분
    Call PrintDot(115, 73, Format(grdOrder.TextMatrix(grdOrder.Row, 8), "#,###")) '오더량
    Call PrintDot(138, 73, IIf(grdOrder.TextMatrix(grdOrder.Row, 15) = "0", "Y", "M"))  '단위
    
    Call PrintDot(150, 73, Format(grdSum.TextMatrix(0, 1), "#,###"))    '절수
    Call PrintDot(163, 73, Format(grdSum.TextMatrix(0, 2), "#,###"))    '출고량
    
End Function

Private Function PrintDot(nXPos As Integer, nYPos As Integer, sStr As String, Optional nFont As Integer = 10)
    With Printer
        .CurrentX = nXPos
        .CurrentY = nYPos
        .Font.Size = nFont
    End With
    Printer.Print sStr
End Function

Private Sub FillGridColor()
    Dim oOrder As PlusLib2.COrder
    Dim rs     As Recordset
    Dim i%

    On Error GoTo ErrHandler


    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon

    Set rs = oOrder.GetOrderSub(txtOrderID1.Tag)
    Set oOrder = Nothing

    With grdColor
        .Redraw = flexRDNone
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & False & vbTab & rs!OrderSeq & vbTab & rs!Color
            
            rs.MoveNext
        Next i
        .Redraw = flexRDDirect
    End With
    rs.Close
    Set rs = Nothing

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing

    Call ErrorBox(Err.Number, "frmOutwareIns.FillGridColor", Err.Description)
End Sub

Private Sub FillGridRoll()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, sColor$
    
    On Error GoTo ErrHandler
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Call GetColor(sColor)
    
    Set rs = oOutware.GetInspect(txtOrderID1.Tag, IIf(chkSearch1(0).Value, 1, 0), MakeDate(DF_SHORT, dtpDate(2)), MakeDate(DF_SHORT, dtpDate(3)), _
                            IIf(chkSearch1(1).Value, 1, 0), cboGrade.ListIndex + 1, IIf(chkSearch1(2).Value, 1, 0), sColor)
    Set oOutware = Nothing
    
    With grdRoll
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & True & vbTab & rs!Color & vbTab & rs!LotNo & vbTab & rs!RollNo & vbTab & _
                    rs!CtrlQty & vbTab & rs!LossQty & vbTab & rs!OrderSeq & vbTab & rs!GradeID & vbTab & _
                    rs!RollSeq & vbTab & rs!OutClss
                    
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
    
ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareIns.FillGridRoll", Err.Description)
End Sub

Private Sub FillGridRollSum()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, sColor$
    
    On Error GoTo ErrHandler
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Call GetColor(sColor)
    
    Set rs = oOutware.GetInspectByColorSum(txtOrderID1.Tag, IIf(chkSearch1(0).Value, 1, 0), MakeDate(DF_SHORT, dtpDate(2)), MakeDate(DF_SHORT, dtpDate(3)), _
                            IIf(chkSearch1(1).Value, 1, 0), cboGrade.ListIndex + 1, IIf(chkSearch1(2).Value, 1, 0), sColor)
    Set oOutware = Nothing
    
    With grdRollSum
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & rs!OrderSeq & vbTab & rs!Color & vbTab & rs!ColorQty & vbTab & rs!OutQty & vbTab & rs!ColorQty - rs!OutQty & vbTab & rs!InspectQty
                    
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
    
ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareIns.FillGridRollSum", Err.Description)
End Sub

Private Sub FillGridOutSum()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, sColor$
    
    On Error GoTo ErrHandler
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Call GetColor(sColor)
    
    Set rs = oOutware.GetOutwareSubTotal(txtOrderID.Tag, txtOrder.Tag)
    Set oOutware = Nothing
    
    With grdOutSum
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & rs!OrderSeq & vbTab & rs!Color & vbTab & rs!ColorQty & vbTab & rs!OutQty & vbTab & rs!ColorQty - rs!OutQty & vbTab & rs!OutQty
                    
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
    
ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareIns.FillGridOutSum", Err.Description)
End Sub

Private Sub GetColor(sColor As String)
    Dim i%
    
    sColor = "A.OrderSeq IN ("
    
    With grdColor
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                sColor = sColor & .TextMatrix(i, 2) & ", "
            End If
        Next i
    End With
    
    If Len(sColor) > 2 Then
        sColor = Left(sColor, Len(sColor) - 2) & ")"
    End If
End Sub

Private Sub FillGridOut()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, sColor$
    
    On Error GoTo ErrHandler
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Call GetColor(sColor)
    
    Set rs = oOutware.GetOutwareSub(txtOrderID.Tag, txtOrder.Tag)
    Set oOutware = Nothing
    
    With grdOut
        .Redraw = flexRDNone
        .Rows = .FixedRows
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & True & vbTab & rs!Color & vbTab & rs!LotNo & vbTab & rs!RollNo & vbTab & _
                    rs!OutQty & vbTab & rs!LossQty & vbTab & rs!OrderSeq & vbTab & rs!GradeID & vbTab & _
                    rs!RollSeq
                    
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
    
ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareIns.FillGridOut", Err.Description)
End Sub


Private Sub FindOrder()
    Dim i%
    
    With grdOrder
        For i = .FixedRows To .Rows - 1
            If m_sOrderID = MakeOrderID(.TextMatrix(i, 1), OM_REDUCE) And m_nOutSeq = .TextMatrix(i, 13) Then
                .Row = i
                .TopRow = i
                Exit Sub
            End If
        Next i
    End With
End Sub

'S_201105_태을염직_01 에 의한 추가
Private Sub MakeExcelPacking()
    Dim oCustom                         As PlusLib2.CCustom
    Dim rs                              As ADODB.Recordset
    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oFs                             As FileSystemObject
    Dim i%, j%, nRow%, nLimitRow%, nBaseRow%, nCurRow%, nPage%, sReport$
    Dim nOrderSeq%, sLotNo$
    Dim sUnit$, nColorRoll%, nColorQty#
    Dim vCustom(5)                      As String
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    
    
    '****************************************************************************************
    '공급받는자 정보 Get
    '---------------------------------------------------------------------------------------
    Set oCustom = New PlusLib2.CCustom
    oCustom.Connection = g_adoCon
    
    Set rs = oCustom.GetCustomOne(grdOrder.TextMatrix(grdOrder.Row, 20))
    Set oCustom = Nothing
        
    If rs.EOF Then
        For i = 0 To 5
            vCustom(i) = ""
        Next i
    Else
        vCustom(0) = IIf(Len(CheckNull(rs!CustomNo)) > 0, Left(rs!CustomNo, 3) & " - " & Mid(rs!CustomNo, 4, 2) & " - " & Right(rs!CustomNo, 5), "")
        vCustom(1) = CheckNull(rs!kCustom)
        vCustom(2) = CheckNull(rs!Chief)

''        'S_201312_태을염직_99 에 의한 수정-OLD소스
''        vCustom(3) = CheckNull(rs!Address1) & " " & CheckNull(rs!Address2)
        
        'S_201312_태을염직_99 에 의한 수정-NEW소스
        If CheckNull(rs!Address1) <> "" Then            '도로명 주소값 있을경우
            vCustom(3) = CheckNull(rs!Address1) & " " & CheckNull(rs!Address2)
        Else
            vCustom(3) = CheckNull(rs!AddressJiBun1) & " " & CheckNull(rs!AddressJiBun2)
        End If
        
        vCustom(4) = CheckNull(rs!Condition)
        vCustom(5) = CheckNull(rs!Category)
    End If
    Set rs = Nothing
    '****************************************************************************************
    
    Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & REPORTFILE)

    oExcel.WindowState = xlMaximized
    'oExcel.Application.Visible = True                                       '개발자 debug Mode 에서는 풀어도됨. S_201204_태을염직_02 추가

    sUnit = IIf(grdOrder.TextMatrix(grdOrder.Row, 15) = "0", "Y", "M")      '수주단위
    With oExcel
        
        .Worksheets("Form").Activate
''        .Cells(4, 1) = MakeDate(DF_FULL, dtpOutDate.Value)
        '송장번호-출고일자
        .Cells(4, 1) = "일련번호:" & Left(m_sTranNo, 4) & "-" & Right(m_sTranNo, 2) & "-" & m_nTranSeq & _
                      Space(15) & MakeDate(DF_FULL, dtpOutDate.Value)

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
        
        .Cells(38, 1) = "●가공 불량으로 원단반품 시, 커팅 된 원단은 절대 반품받지 않음."
        .Cells(39, 1) = "●출고 후 30일 이후 이의제기 불가"
        .Cells(40, 1) = "●가공완료 후 3개월 이상 장기간 보관중인 원단(분실,화재)사고 발생시,"
        .Cells(41, 1) = "   당사에서는 책임(변상)지지 않음."
        '*****************************************************************
        
        
        '****************************************************************************************
        '공급받는자 정보 출력
        '---------------------------------------------------------------------------------------
        .Cells(5, 18) = vCustom(0)                                          '사업자번호
        .Cells(6, 18) = vCustom(1)                                          '회사명
        .Cells(6, 26) = vCustom(2)                                          '대표
        .Cells(7, 18) = vCustom(3)                                          '주소
        .Cells(8, 18) = vCustom(4)                                          '업태
        .Cells(8, 26) = vCustom(5)                                          '종목
        '****************************************************************************************
        
        Dim nChunRateQty As Long
        
        '2011.05.19 김대진 대리 요청- ORderNo대신 관리번호
''        .Cells(9, 4) = Trim(txtOrder)                                       'order No
        .Cells(9, 4) = MakeOrderID(Trim(txtOrderID), OM_EXPAND)                                       '관리번호-OrderID
        .Cells(9, 13) = Format(grdOrder.TextMatrix(grdOrder.Row, 8), "#,###") & " " & sUnit           'Order 량
        
'        If grdOrder.TextMatrix(grdOrder.Row, 15) = "0" Then
'            nChunRateQty = grdSum.TextMatrix(0, 2) + (grdSum.TextMatrix(0, 2) * txtChunkRate.Text / 100)
'
'        Else
'            nChunRateQty = (grdSum.TextMatrix(0, 2) * 1.0936) + ((grdSum.TextMatrix(0, 2) * 1.0936) * txtChunkRate.Text / 100)
'
'        End If
         
      '  .Cells(9, 17) = txtChunkRate.Text + " %"                            '축율
        .Cells(9, 21) = Trim(txtOutCustom.Text)                             '출고처
        
        .Cells(12, 1) = Trim(txtArticle.Text)                               '품명
        .Cells(12, 7) = Trim(grdOrder.TextMatrix(grdOrder.Row, 17))         '규격
        .Cells(12, 11) = Trim(grdOrder.TextMatrix(grdOrder.Row, 19))         '가공구분
        
''        .Cells(12, 11) = Format(grdSum.TextMatrix(0, 1), "#,###")           '생지마수-절***
''        .Cells(12, 13) = Format(txtOutRealQty(1), "#,###")                  '생지마수- 길이(소요량)  ***
        
        .Cells(12, 14) = Format(grdSum.TextMatrix(0, 1), "#,###")           '가공마수-절 ***
        If txtOutCustom.Tag = "0" Then
            .Cells(12, 17) = Format(grdSum.TextMatrix(0, 2), "#,###") & "Y" '가공마수 길이
            
        Else
             .Cells(12, 17) = Format(grdSum.TextMatrix(0, 2), "#,###") & "M" & vbLf & _
                              "(" & Format(CLng((grdSum.TextMatrix(0, 2) / 0.9144)), "#,###") & "Y)"  ' 가공마수길이-출고량
        End If
        
        'S_201105_태을염직_02 에 의한 추가
        .Cells(10, 22) = Trim(txtOrder)                                'OrderNo
        .Cells(12, 22) = Trim(txtRemark.Text)                               '비고 (OLD:12,20)
        
        .Worksheets("Print").Activate
        
        nPage = 1
        nBaseRow = GetExcelRollBaseRow(nPage)
        Call InsertExcelForm(oExcel, nPage)
        nCurRow = nBaseRow + 15
        For i = grdPacking.FixedRows To grdPacking.Rows - 1
            If nCurRow + nRow > nBaseRow + 37 Then
                nPage = nPage + 1
                nBaseRow = GetExcelRollBaseRow(nPage)
                Call InsertExcelForm(oExcel, nPage)
                nCurRow = nBaseRow + 15
                nRow = 0
            End If
        
            If nOrderSeq <> grdPacking.TextMatrix(i, 15) Then                                       'OrderSeq
                If i > grdPacking.FixedRows Then
                    nRow = nRow + 1
                    .Cells(nCurRow + nRow - 1, 3) = "COL 계 :"
                    .Cells(nCurRow + nRow - 1, 6) = Format(nColorRoll, "#,###")
                    .Cells(nCurRow + nRow - 1, 8) = Format(nColorQty, "#,###")
                End If
                If nCurRow + nRow > nBaseRow + 37 Then
                    nPage = nPage + 1
                    nBaseRow = GetExcelRollBaseRow(nPage)
                    Call InsertExcelForm(oExcel, nPage)
                    nCurRow = nBaseRow + 15
                    nRow = 0
                End If
                .Cells(nCurRow + nRow, 1) = grdPacking.TextMatrix(i, 15)                            'OrderSeq
                .Cells(nCurRow + nRow, 3) = Trim(grdPacking.TextMatrix(i, 1))                       'Color
                .Cells(nCurRow + nRow, 5) = grdPacking.TextMatrix(i, 2)                             'Lot
                .Cells(nCurRow + nRow, 6) = Format(grdPacking.TextMatrix(i, 13), "#,###")           'PCS(절수)
                .Cells(nCurRow + nRow, 8) = Format(grdPacking.TextMatrix(i, 14), "#,###")           '수량
                
                nColorRoll = CheckNum(grdPacking.TextMatrix(i, 13))                                  'PCS(절수)
                nColorQty = CheckNum(grdPacking.TextMatrix(i, 14))                                   '수량
            ElseIf nOrderSeq = grdPacking.TextMatrix(i, 15) And sLotNo <> grdPacking.TextMatrix(i, 2) Then
                .Cells(nCurRow + nRow, 5) = grdPacking.TextMatrix(i, 2)                             'Lot
                .Cells(nCurRow + nRow, 6) = Format(grdPacking.TextMatrix(i, 13), "#,###")           'PCS(절수)
                .Cells(nCurRow + nRow, 8) = Format(grdPacking.TextMatrix(i, 14), "#,###")           '수량
                
                nColorRoll = nColorRoll + CheckNum(grdPacking.TextMatrix(i, 13))                    'PCS(절수)
                nColorQty = nColorQty + CheckNum(grdPacking.TextMatrix(i, 14))                      '수량
            End If
            
           .Cells(nCurRow + nRow, 11) = grdPacking.TextMatrix(i, 3)
           .Cells(nCurRow + nRow, 13) = grdPacking.TextMatrix(i, 4)
           .Cells(nCurRow + nRow, 14) = grdPacking.TextMatrix(i, 5)
           .Cells(nCurRow + nRow, 16) = grdPacking.TextMatrix(i, 6)
           .Cells(nCurRow + nRow, 19) = grdPacking.TextMatrix(i, 7)
           .Cells(nCurRow + nRow, 20) = grdPacking.TextMatrix(i, 8)
           .Cells(nCurRow + nRow, 22) = grdPacking.TextMatrix(i, 9)
           .Cells(nCurRow + nRow, 25) = grdPacking.TextMatrix(i, 10)
           .Cells(nCurRow + nRow, 27) = grdPacking.TextMatrix(i, 11)
           .Cells(nCurRow + nRow, 29) = grdPacking.TextMatrix(i, 12)
           
            nOrderSeq = grdPacking.TextMatrix(i, 15)
            sLotNo = grdPacking.TextMatrix(i, 2)
            nRow = nRow + 1
        Next i
        
        If nCurRow + nRow > nBaseRow + 37 Then
            nPage = nPage + 1
            nBaseRow = GetExcelRollBaseRow(nPage)
            Call InsertExcelForm(oExcel, nPage)
            nCurRow = nBaseRow + 15
            nRow = 0
        End If
        .Cells(nCurRow + nRow, 3) = "COL 계 : "
        .Cells(nCurRow + nRow, 6) = Format(nColorRoll, "#,###")
        .Cells(nCurRow + nRow, 8) = Format(nColorQty, "#,###")
        
        
        
        
        
    End With

    sReport = App.Path & REPORTFILE1

    Set oFs = New FileSystemObject
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
    Set oFs = Nothing

    Call oExcelBook.SaveAs(sReport)

    
    '인쇄 미리보기 일때만 미리보기, S_201204_태을염직_02 의한 추가
    If PlusMDI.PrintPreview = True Then
        oExcel.WindowState = xlMaximized
        oExcel.Application.Visible = True
        oExcel.ActiveWindow.SelectedSheets.PrintPreview
    Else
        oExcel.ActiveWindow.SelectedSheets.PrintOut
        Call ProcessClose("XLMAIN")
    End If

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

    MsgBox Err.Number & "," & Err.Description & "," & Erl, vbCritical, "frmOutwareIns.MakeExcelPacking"
    
    

End Sub

'S_201105_태을염직_01 에 의한 추가
Private Function GetExcelRollBaseRow(nPage)
    GetExcelRollBaseRow = (nPage - 1) * EXCEL_ROLL_ROW
End Function

'S_201105_태을염직_01 에 의한 추가
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

