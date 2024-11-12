VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmModiRecipe 
   ClientHeight    =   9300
   ClientLeft      =   -75
   ClientTop       =   480
   ClientWidth     =   14970
   Icon            =   "frmModiRecipe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   14970
   Begin Threed.SSPanel pnlMsg 
      Height          =   585
      Left            =   0
      TabIndex        =   73
      Top             =   8700
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   1032
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdUP 
         Height          =   525
         Left            =   30
         TabIndex        =   74
         Top             =   30
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   926
         _Version        =   196609
         Alignment       =   8
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   525
         Left            =   990
         TabIndex        =   75
         Top             =   30
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   926
         _Version        =   196609
         Alignment       =   8
      End
      Begin VB.Label Label2 
         Caption         =   "※  계획공정을 추가하려면 공정코드 리스트에서 [더블클릭] 하십시요"
         Height          =   165
         Index           =   1
         Left            =   2100
         TabIndex        =   77
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "※  계획공정을 제거하려면 [더블클릭] 하십시요"
         Height          =   165
         Index           =   0
         Left            =   2100
         TabIndex        =   76
         Top             =   90
         Width           =   4725
      End
   End
   Begin Threed.SSPanel pnlProcess 
      Height          =   3525
      Left            =   0
      TabIndex        =   68
      Top             =   5160
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   6218
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdCardPattern 
         Height          =   3165
         Left            =   30
         TabIndex        =   69
         Top             =   330
         Width           =   1935
         _cx             =   3413
         _cy             =   5583
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
      Begin VSFlex7LCtl.VSFlexGrid grdProcess 
         Height          =   3165
         Left            =   1980
         TabIndex        =   70
         Top             =   330
         Width           =   1935
         _cx             =   3413
         _cy             =   5583
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
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   12
         Left            =   30
         TabIndex        =   71
         Top             =   30
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "실적 및 계획공정"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   14
         Left            =   1980
         TabIndex        =   72
         Top             =   30
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "공정코드 리스트"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "발행 여부"
      Height          =   585
      Left            =   9420
      TabIndex        =   65
      Top             =   8700
      Width           =   2475
      Begin VB.OptionButton optPrn 
         Caption         =   "미 발행"
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   67
         Top             =   210
         Width           =   945
      End
      Begin VB.OptionButton optPrn 
         Caption         =   "저장후 발행"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   66
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   570
      Left            =   11925
      TabIndex        =   2
      Top             =   8715
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1005
      _Version        =   196609
      Caption         =   "      저장(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   570
      Left            =   11925
      TabIndex        =   0
      Top             =   8685
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1005
      _Version        =   196609
      Enabled         =   0   'False
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   13575
      TabIndex        =   1
      Top             =   8715
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1005
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   8235
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   3525
      Left            =   3930
      TabIndex        =   3
      Top             =   5160
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   6218
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtProcOpinion 
         Height          =   285
         Left            =   5730
         TabIndex        =   56
         Top             =   30
         Width           =   5535
      End
      Begin VB.TextBox txtProcPerson 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3270
         TabIndex        =   55
         Top             =   15
         Width           =   1125
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
         Height          =   2805
         Index           =   0
         Left            =   3240
         TabIndex        =   4
         Top             =   705
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
         Height          =   3150
         Left            =   30
         TabIndex        =   5
         Top             =   345
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   5556
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtBox 
            BackColor       =   &H00FFC0C0&
            Height          =   300
            Index           =   0
            Left            =   1065
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   15
            Top             =   45
            Width           =   2100
         End
         Begin VB.TextBox txtBox 
            BackColor       =   &H00FFC0C0&
            Height          =   300
            Index           =   1
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   645
            Width           =   2100
         End
         Begin VB.TextBox txtBox 
            BackColor       =   &H00FFFFC0&
            Height          =   300
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1890
            Width           =   2100
         End
         Begin VB.TextBox txtBox 
            Height          =   300
            Index           =   3
            Left            =   1080
            TabIndex        =   11
            Top             =   1575
            Width           =   1770
         End
         Begin VB.TextBox txtTemp 
            Height          =   300
            Left            =   2430
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1290
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtBox 
            BackColor       =   &H00FFC0C0&
            Height          =   300
            Index           =   4
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   345
            Width           =   2100
         End
         Begin VB.TextBox txtRemark 
            Height          =   660
            IMEMode         =   10  '한글 
            Left            =   1080
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   8
            Top             =   2505
            Width           =   2115
         End
         Begin VB.TextBox txtBox 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   5
            Left            =   1065
            TabIndex        =   7
            Top             =   960
            Width           =   2100
         End
         Begin VB.TextBox txtBox 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   6
            Left            =   1080
            TabIndex        =   6
            Top             =   2190
            Width           =   2100
         End
         Begin MSComCtl2.DTPicker dtpRecipe 
            Height          =   300
            Left            =   1080
            TabIndex        =   12
            Top             =   1275
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   529
            _Version        =   393216
            Format          =   68616193
            CurrentDate     =   37112
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   30
            TabIndex        =   16
            Top             =   45
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   30
            TabIndex        =   17
            Top             =   645
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   30
            TabIndex        =   18
            Top             =   1890
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   30
            TabIndex        =   19
            Top             =   1275
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   30
            TabIndex        =   20
            Top             =   1575
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "처  방  자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   1
            Left            =   2850
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1605
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
            Left            =   30
            TabIndex        =   22
            Top             =   345
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "품      명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   630
            Index           =   16
            Left            =   30
            TabIndex        =   23
            Top             =   2505
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   1111
            _Version        =   196609
            Caption         =   "특기사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   30
            TabIndex        =   24
            Top             =   960
            Width           =   1020
            _ExtentX        =   1799
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
            Left            =   30
            TabIndex        =   25
            Top             =   2190
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "축      율"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   360
         Index           =   0
         Left            =   6105
         TabIndex        =   26
         Top             =   345
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "염료삭제(&W)"
      End
      Begin Threed.SSCommand cmdAddNew 
         Height          =   360
         Index           =   0
         Left            =   4935
         TabIndex        =   27
         Top             =   345
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "염료추가(&Q)"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
         Height          =   2805
         Index           =   1
         Left            =   7260
         TabIndex        =   28
         Top             =   705
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
         Height          =   360
         Index           =   1
         Left            =   10125
         TabIndex        =   29
         Top             =   345
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "조제삭제(&R)"
      End
      Begin Threed.SSCommand cmdAddNew 
         Height          =   360
         Index           =   1
         Left            =   8955
         TabIndex        =   30
         Top             =   345
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "조제추가(&E)"
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   0
         Left            =   3240
         TabIndex        =   31
         Top             =   360
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   196609
         Caption         =   "염료 사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   4
         Left            =   7290
         TabIndex        =   32
         Top             =   360
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         _Version        =   196609
         Caption         =   "조제 사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Index           =   0
         Left            =   4710
         TabIndex        =   57
         Top             =   30
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   196609
         Caption         =   "처리방안"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   58
         Top             =   30
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   196609
         Caption         =   "처리자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Index           =   2
         Left            =   30
         TabIndex        =   59
         Top             =   30
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   196609
         Caption         =   "처리일시"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpProcDate 
         Height          =   300
         Left            =   1110
         TabIndex        =   60
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   68616193
         CurrentDate     =   36871
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   4410
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   30
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdHold 
      Height          =   4455
      Left            =   0
      TabIndex        =   33
      Top             =   690
      Width           =   15255
      _cx             =   26908
      _cy             =   7858
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
   Begin Threed.SSPanel SSPanel4 
      Height          =   675
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1191
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtSplitID 
         Height          =   300
         Left            =   12810
         MaxLength       =   4
         TabIndex        =   62
         Top             =   30
         Width           =   660
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   2955
         MousePointer    =   99  '사용자 정의
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   330
         Width           =   675
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   2250
         MousePointer    =   99  '사용자 정의
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   330
         Width           =   705
      End
      Begin VB.TextBox txtCardID 
         Height          =   300
         Left            =   11775
         MaxLength       =   8
         TabIndex        =   37
         Top             =   30
         Width           =   1020
      End
      Begin VB.TextBox txtOrderID 
         Height          =   300
         Left            =   7725
         TabIndex        =   36
         Top             =   330
         Width           =   1770
      End
      Begin VB.ComboBox cboProcID 
         Height          =   300
         Left            =   11760
         Style           =   2  '드롭다운 목록
         TabIndex        =   35
         Top             =   330
         Width           =   1725
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   8
         Left            =   6240
         TabIndex        =   38
         Top             =   330
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkOrderID 
            Caption         =   "관리번호"
            Height          =   240
            Left            =   210
            TabIndex        =   39
            Top             =   45
            Width           =   1050
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   9
         Left            =   10470
         TabIndex        =   40
         Top             =   30
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkCardID 
            Caption         =   "카드번호"
            Height          =   240
            Left            =   150
            TabIndex        =   41
            Top             =   45
            Width           =   1050
         End
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3645
         TabIndex        =   42
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   68616193
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3645
         TabIndex        =   43
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   68616193
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   10
         Left            =   2250
         TabIndex        =   44
         Top             =   30
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkDate 
            Caption         =   "보류일자"
            Height          =   240
            Left            =   180
            TabIndex        =   45
            Top             =   45
            Value           =   1  '확인
            Width           =   1140
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   11
         Left            =   10470
         TabIndex        =   46
         Top             =   330
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkProcID 
            Caption         =   "보류공정"
            Height          =   240
            Left            =   150
            TabIndex        =   47
            Top             =   30
            Width           =   1050
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   285
         Left            =   6240
         TabIndex        =   48
         Top             =   30
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   60
            Width           =   1170
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   1770
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   60
            Value           =   -1  'True
            Width           =   1110
         End
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   660
         Left            =   13830
         TabIndex        =   64
         Top             =   0
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1164
         _Version        =   196609
         Caption         =   "        검색(&F)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "보류건 조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   63
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   2
         Left            =   4950
         TabIndex        =   52
         Top             =   375
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   3
         Left            =   4950
         TabIndex        =   51
         Top             =   90
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmModiRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\Recipe.rpt"
Private Const LIMIT_ROW4 = 11
Private Const LIMIT_WIDTH4 = 2085

Private gSelCnt As Integer

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

Private Sub cmdDown_Click()
Dim iRow%, iCol%, iCurRow%

    With grdCardPattern
        If .Rows > .FixedRows And .Row >= .FixedRows And _
            .Cell(flexcpBackColor, .Row, 1) = 0 And .Row < .Rows - 1 Then
            iCurRow = .Row
            .Rows = .Rows + 1
            For iCol = 1 To .Cols - 1
                .TextMatrix(.Rows - 1, iCol) = .TextMatrix(iCurRow + 1, iCol)
            Next iCol
            For iCol = 1 To .Cols - 1
                .TextMatrix(iCurRow + 1, iCol) = .TextMatrix(iCurRow, iCol)
            Next iCol
            For iCol = 1 To .Cols - 1
                .TextMatrix(iCurRow, iCol) = .TextMatrix(.Rows - 1, iCol)
            Next iCol
            .Rows = .Rows - 1
            .Row = iCurRow + 1
        End If
    End With

End Sub

Private Sub cmdUP_Click()
Dim iRow%, iCol%, iCurRow%

    With grdCardPattern
        If .Rows > .FixedRows And .Row >= .FixedRows And .Cell(flexcpBackColor, .Row, 1) = 0 Then
            If .Cell(flexcpBackColor, .Row - 1, 1) = 0 Then
                iCurRow = .Row
                .Rows = .Rows + 1
                For iCol = 1 To .Cols - 1
                    .TextMatrix(.Rows - 1, iCol) = .TextMatrix(iCurRow - 1, iCol)
                Next iCol
                For iCol = 1 To .Cols - 1
                    .TextMatrix(iCurRow - 1, iCol) = .TextMatrix(iCurRow, iCol)
                Next iCol
                For iCol = 1 To .Cols - 1
                    .TextMatrix(iCurRow, iCol) = .TextMatrix(.Rows - 1, iCol)
                Next iCol
                .Rows = .Rows - 1
                .Row = iCurRow - 1
            End If
        End If
    End With

End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 15360, 9840
    
    If PlusMDI.pnlMenu.Visible = False Then
        PlusMDI.pnlMenu.Visible = True
    End If

    Call SetOperate(Me)

    pnlEdit.Enabled = True
    dtpDate(0) = Now
    dtpDate(1) = Now
    dtpProcDate = Now
    dtpRecipe = Now
    cmdSave.MousePointer = ssCustom
    cmdSave.Picture = LoadResPicture("SAVE", vbResIcon)
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdSave.MousePointer = ssCustom
    cmdUP.Picture = LoadResPicture("UP", vbResIcon)
    cmdDown.Picture = LoadResPicture("DOWN", vbResIcon)

    Call SetComboProcss(cboProcID, AllStr)
    Call InitGrid
    Call ClearData
    Call FillGridProcess
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
Dim i%
    Call SetVSFlexGrid(grdHold)
    With grdHold
        .Redraw = flexRDNone
        
        .Editable = flexEDKbdMouse
        .Cols = 27:         .Rows = 4
        .FixedCols = 1:     .FixedRows = 4

        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 0
        Next i

        .TextMatrix(3, 0) = "":                 .ColWidth(0) = 300
        .TextMatrix(3, 1) = "":                 .ColWidth(1) = 400:             .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "보류일":           .ColWidth(2) = 700:             .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "카드번호":         .ColWidth(3) = 1400:            .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "관리번호":         .ColWidth(4) = 1400:            .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "OrderNo":          .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(3, 6) = "거래처":           .ColWidth(6) = 1200:            .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "품명":             .ColWidth(7) = 2500:            .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(3, 8) = "색상명":           .ColWidth(8) = 2000:            .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "절수":             .ColWidth(9) = 600:             .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "수량":            .ColWidth(10) = 700:            .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(3, 11) = "보류공정":        .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignCenterCenter
        .TextMatrix(3, 12) = "보류원인":        .ColWidth(12) = 1000:           .ColAlignment(12) = flexAlignLeftCenter
        .TextMatrix(3, 13) = "작성자":          .ColWidth(13) = 0:              .ColAlignment(13) = flexAlignCenterCenter
        .TextMatrix(3, 14) = "축율":            .ColWidth(14) = 0:              .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(3, 15) = "WorkUnitID":      .ColWidth(15) = 0:              .ColAlignment(15) = flexAlignCenterCenter
        .TextMatrix(3, 16) = "WorkUnitSeq":     .ColWidth(16) = 0:              .ColAlignment(16) = flexAlignCenterCenter
        
        .TextMatrix(3, 20) = "CardID"
        .TextMatrix(3, 21) = "SplitID"
        .TextMatrix(3, 22) = "OrderID"
        .TextMatrix(3, 23) = "OrderSeq"
        .TextMatrix(3, 24) = "WriteDate"
        .TextMatrix(3, 25) = "WriteProcID"
        .TextMatrix(3, 26) = "WriteProcSeq"
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .RowHeight(0) = 400
        .RowHeight(1) = 400
        .RowHeight(2) = 400
        .RowHeight(3) = 400
        
        .ColDataType(1) = flexDTBoolean

        .MergeCells = flexMergeFree
        
        .AllowUserResizing = flexResizeBoth
        .Redraw = flexRDDirect
    End With

    With grdDyeAux(0)
        .Cols = 6
        Call SetVSFlexGrid(grdDyeAux(0))

        .Redraw = flexRDNone

        .TextArray(1) = "염료":         .ColWidth(1) = 2085:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "염료":         .ColWidth(2) = 300:         .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "염료투입비율": .ColWidth(3) = 1200:        .ColAlignment(3) = flexAlignRightCenter
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

        .TextArray(1) = "조제":         .ColWidth(1) = 2085:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "조제":         .ColWidth(2) = 300:     .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "조제투입비율": .ColWidth(3) = 1200:    .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "조제":         .ColWidth(4) = 0
        .TextArray(5) = "조제명":       .ColWidth(5) = 0

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With
    
    With grdCardPattern
        .Redraw = flexRDNone
        .Cols = 7
        
        Call SetVSFlexGrid(grdCardPattern)
        .ExplorerBar = flexExNone

        .Rows = 1
        
        .TextArray(0) = "순서":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColHidden(1) = True
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "완료여부":     .ColHidden(3) = True
        .TextArray(4) = "요구폭":       .ColHidden(4) = True
        .TextArray(5) = "지시사항":     .ColHidden(5) = True
        .TextArray(6) = "비고":         .ColHidden(6) = True
        
        .Redraw = flexRDDirect
    End With
    
    With grdProcess
        .Redraw = flexRDNone
        .Cols = 3
        
        Call SetVSFlexGrid(grdProcess)
        .ExplorerBar = flexExNone
        
        .Rows = 1
        
        .TextArray(0) = "":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColHidden(1) = True
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub FillGridProcess()
    Dim oCard As PlusLib2.CCard
    Dim Rs As Recordset
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon

    Set Rs = oCard.GetProcess()
    Set oCard = Nothing

    With grdProcess
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        Do Until Rs.EOF
            .AddItem CStr(.Rows) & vbTab & Rs!processid & vbTab & Rs!Process
            Rs.MoveNext
        Loop
        
        Rs.Close
        Set Rs = Nothing
        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHandler:
    Set Rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmModiRecipe.FillGridProcess", Err.Description)
End Sub

Private Sub ClearData()
    Dim oRecipe As PlusLib2.CRecipe

    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    dtpRecipe = Date

    txtBox(2) = oRecipe.GetMaxRecipeNo
    
    txtBox(0) = ""
    txtBox(1) = ""
    txtBox(3) = ""
    txtBox(3).Tag = ""
    txtBox(4) = ""
    txtBox(5) = 0
    txtBox(6) = 0
    txtRemark = ""
    txtProcOpinion = ""
    txtProcPerson.Text = g_sPersonName
    txtProcPerson.Tag = g_sUserName
    txtBox(3).Text = g_sPersonName
    txtBox(3).Tag = g_sUserName
    
    grdDyeAux(0).Rows = grdDyeAux(0).FixedRows
    grdDyeAux(1).Rows = grdDyeAux(1).FixedRows
    
    Set oRecipe = Nothing

End Sub

Private Sub cmdSearch_Click()
    Call FillgrdHold
End Sub

Private Sub FillgrdHold()
    Dim oCard As PlusLib2.CCard
    Dim dRS As ADODB.Recordset
    Dim TParaHold As TCardHold
    Dim iCnt%, iCount%
    Dim sWorkUnitID$
    Dim bToggle As Boolean
    
    With TParaHold
        If chkOrderID.Value = vbChecked Then
            If optOrder(0).Value = True Then  'Order NO
                .nCheckOrderID = 0
                .OrderID = ""
                .nCheckOrderNo = 1
                .OrderNo = txtOrderID.Text
            Else
                .nCheckOrderID = 1
                .OrderID = txtOrderID.Text
                .nCheckOrderNo = 0
                .OrderNo = ""
            End If
        Else
            .nCheckOrderID = 0
            .OrderID = ""
            .nCheckOrderNo = 0
            .OrderNo = ""
        End If
        
        If chkDate.Value = vbChecked Then
            .nCheckDate = 1
            .sDate = MakeDate(DF_SHORT, dtpDate(0))
            .eDate = MakeDate(DF_SHORT, dtpDate(1))
        Else
            .nCheckDate = 0
            .sDate = ""
            .eDate = ""
        End If
        
        If chkCardID.Value = vbChecked Then
            .nCheckCardID = 1
            .CardID = txtCardID.Text
            .SplitID = Trim(txtSplitID.Text)
        Else
            .nCheckCardID = 0
            .CardID = ""
            .SplitID = ""
        End If
        
        If chkProcID.Value = vbChecked Then
            .nCheckProcID = 1
            .WriteProcID = GetProcessID(Trim(cboProcID.Text))
        Else
            .nCheckProcID = 0
            .WriteProcID = ""
        End If
    End With
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon

    Set dRS = oCard.GetHoldingCard(TParaHold)
    
    With grdHold
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        If .Rows > .FixedRows Then
            .Row = 1
        End If

        For iCnt = 1 To dRS.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 350
            
            If iCount = 1 Then
                sWorkUnitID = dRS!WorkUnitId
            End If
            If sWorkUnitID <> dRS!WorkUnitId Then
                bToggle = Not (bToggle)
            End If
            
            
            .TextMatrix(.Rows - 1, 0) = CStr(iCnt)
            .TextMatrix(.Rows - 1, 2) = MakeDate(DF_MD, dRS!WriteDate)
            .TextMatrix(.Rows - 1, 3) = IIf(Trim(dRS!SplitID) = "", MakeCardID(dRS!CardID, OM_EXPAND), MakeCardID(dRS!CardID, OM_EXPAND) & "(" & Trim(dRS!SplitID) & ")")
            .TextMatrix(.Rows - 1, 4) = MakeOrderID(dRS!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 5) = Trim(dRS!OrderNo)
            .TextMatrix(.Rows - 1, 6) = Trim(dRS!kCustom)
            .TextMatrix(.Rows - 1, 7) = Trim(dRS!Article)
            .TextMatrix(.Rows - 1, 8) = Trim(dRS!Color)
            .TextMatrix(.Rows - 1, 9) = dRS!Roll
            .TextMatrix(.Rows - 1, 10) = Format(dRS!Qty, "##,##0")
            .TextMatrix(.Rows - 1, 11) = Trim(dRS!Process)
            .TextMatrix(.Rows - 1, 12) = Trim(dRS!HoldReason)
            .TextMatrix(.Rows - 1, 13) = Trim(dRS!PersonID)
            .TextMatrix(.Rows - 1, 14) = dRS!ChunkRate
            .TextMatrix(.Rows - 1, 15) = dRS!WorkUnitId
            .TextMatrix(.Rows - 1, 16) = dRS!WorkUnitSeq
            
            .TextMatrix(.Rows - 1, 20) = dRS!CardID
            .TextMatrix(.Rows - 1, 21) = Trim(dRS!SplitID)
            .TextMatrix(.Rows - 1, 22) = dRS!OrderID
            .TextMatrix(.Rows - 1, 23) = dRS!OrderSeq
            .TextMatrix(.Rows - 1, 24) = dRS!WriteDate
            .TextMatrix(.Rows - 1, 25) = dRS!WriteProcID
            .TextMatrix(.Rows - 1, 26) = dRS!WriteSeq
            
            
             If bToggle = True Then
                 .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
             Else
                 .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 0
             End If
            
             sWorkUnitID = dRS!WorkUnitId
            
            
            
            dRS.MoveNext
        Next iCnt
        dRS.Close
        Set dRS = Nothing
        .Redraw = flexRDDirect
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
'        Else
'            MsgBox LoadResString(203), vbInformation
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdAddNew_Click(Index As Integer)
    With grdDyeAux(Index)
        .Rows = .Rows + 1

        Call ChangeScrollDyeAux(Index)

        .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
        .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter
        .SetFocus
        .Select .Rows - 1, 1
    End With

End Sub

Private Sub ChangeScrollDyeAux(Index As Integer)
    With grdDyeAux(Index)
        .ColWidth(1) = IIf(.Rows > LIMIT_ROW4 + .FixedRows, LIMIT_WIDTH4 - 240, LIMIT_WIDTH4)
    End With
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    With grdDyeAux(Index)
        If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub

        .RemoveItem .Row

        cmdSave.SetFocus
    End With
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then       ' 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    Else                    ' 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    ' 입력 - 처리자
    If Index = 0 Then
        Call ReturnCode(LG_PERSON, , False, txtProcPerson)
    ' 입력 - 처방자
    ElseIf Index = 1 Then
        Call ReturnCode(LG_PERSON, , False, txtBox(3))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed

End Sub

Private Sub grdCardPattern_DblClick()
Dim iRow As Integer
Dim iCol As Integer

    With grdCardPattern
        If .Rows > .FixedRows And .Row >= .FixedRows And .Cell(flexcpBackColor, .Row, 1) = 0 Then
            For iRow = .Row To .Rows - 2
                .TextMatrix(iRow, 0) = CStr(CInt(.TextMatrix(iRow + 1, 0)) - 1)
                For iCol = 1 To .Cols - 1
                    .TextMatrix(iRow, iCol) = .TextMatrix(iRow + 1, iCol)
                Next iCol
            Next iRow
            .Rows = .Rows - 1
        End If
    End With

End Sub


Private Sub grdHold_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdHold
        If Col = 1 Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub grdHold_RowColChange()
    With grdHold
        If .Rows > .FixedRows Then
            txtBox(0).Tag = .TextMatrix(.Row, 22)
            txtBox(0).Text = .TextMatrix(.Row, 4)
            txtBox(4) = .TextMatrix(.Row, 7)
            txtBox(1).Tag = .TextMatrix(.Row, 23)
            txtBox(1).Text = .TextMatrix(.Row, 8)
            txtBox(6) = .TextMatrix(.Row, 14)
            txtProcOpinion = "수정 처방전 작성"
            
            Call FillGridPattern
        End If
    End With
End Sub

Private Sub FillGridPattern()
    Dim oCard As PlusLib2.CCard
    Dim Rs As ADODB.Recordset
    Dim i%, iSeq%
    
    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set Rs = oCard.GetCardPattern(grdHold.TextMatrix(grdHold.Row, 20), Trim(grdHold.TextMatrix(grdHold.Row, 21)))
    Set oCard = Nothing
    
    If Rs.EOF Then
        grdCardPattern.Rows = grdCardPattern.FixedRows
        Set Rs = Nothing
        Exit Sub
    End If
    
    With grdCardPattern
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        iSeq = 0
        For i = 1 To Rs.RecordCount
            If Rs!CompleteClss = "*" Then
                .Rows = .Rows + 1
                iSeq = Rs!PlanSeq
                .TextMatrix(.Rows - 1, 0) = Rs!PlanSeq
                .TextMatrix(.Rows - 1, 1) = Rs!processid
                .TextMatrix(.Rows - 1, 2) = Rs!Process
                .TextMatrix(.Rows - 1, 3) = Rs!CompleteClss
                .TextMatrix(.Rows - 1, 4) = Rs!NeedWidth
                .TextMatrix(.Rows - 1, 5) = Rs!InstRemark
                .TextMatrix(.Rows - 1, 6) = Rs!Remark
            Else
                Exit For
            End If
            Rs.MoveNext
        Next i
        
        .Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1, 6) = &HFFFFC0
        
        .Rows = .Rows + 1
        iSeq = iSeq + 1
        .TextMatrix(.Rows - 1, 0) = CStr(iSeq)
        .TextMatrix(.Rows - 1, 1) = "4301"
        .TextMatrix(.Rows - 1, 2) = "염색"
        .TextMatrix(.Rows - 1, 3) = ""
        .TextMatrix(.Rows - 1, 4) = ""
        .TextMatrix(.Rows - 1, 5) = ""
        .TextMatrix(.Rows - 1, 6) = ""
        
        .Rows = .Rows + 1
        iSeq = iSeq + 1
        .TextMatrix(.Rows - 1, 0) = CStr(iSeq)
        .TextMatrix(.Rows - 1, 1) = "6401"
        .TextMatrix(.Rows - 1, 2) = "Dry"
        .TextMatrix(.Rows - 1, 3) = ""
        .TextMatrix(.Rows - 1, 4) = ""
        .TextMatrix(.Rows - 1, 5) = ""
        .TextMatrix(.Rows - 1, 6) = ""
        
        .Rows = .Rows + 1
        iSeq = iSeq + 1
        .TextMatrix(.Rows - 1, 0) = CStr(iSeq)
        .TextMatrix(.Rows - 1, 1) = "7601"
        .TextMatrix(.Rows - 1, 2) = "가공"
        .TextMatrix(.Rows - 1, 3) = ""
        .TextMatrix(.Rows - 1, 4) = ""
        .TextMatrix(.Rows - 1, 5) = ""
        .TextMatrix(.Rows - 1, 6) = ""
        
        .Rows = .Rows + 1
        iSeq = iSeq + 1
        .TextMatrix(.Rows - 1, 0) = CStr(iSeq)
        .TextMatrix(.Rows - 1, 1) = "8201"
        .TextMatrix(.Rows - 1, 2) = "검사"
        .TextMatrix(.Rows - 1, 3) = ""
        .TextMatrix(.Rows - 1, 4) = ""
        .TextMatrix(.Rows - 1, 5) = ""
        .TextMatrix(.Rows - 1, 6) = ""
        
        .Redraw = flexRDDirect
    End With
    
    Rs.Close
    Set Rs = Nothing
        
    Exit Sub
    
ErrHandler:
    Set oCard = Nothing
    Set Rs = Nothing
    
    Call ErrorBox(Err.Number, "frmCardPattern.FillGridPattern", Err.Description)
End Sub

Private Sub grdDyeAux_Click(Index As Integer)
    With grdDyeAux(Index)
        If .MouseRow < .FixedRows Or .MouseRow > .Rows - 1 Or .MouseCol <> 2 Then Exit Sub

        Dim Row%
        Row = .MouseRow
        txtTemp = .TextMatrix(Row, 1)

        If ReturnCode(IIf(Index = 0, LG_DYE, LG_AUX), , False, txtTemp) Then
            .TextMatrix(Row, 1) = txtTemp
            .TextMatrix(Row, 5) = txtTemp
            .TextMatrix(Row, 4) = txtTemp.Tag
        End If
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
        ElseIf Col = 3 Then
            .Cell(flexcpText, Row, Col) = SetCurrency(.TextMatrix(Row, Col), 6)

            If Row = .Rows - 1 Then
                If QuestionBox(IIf(Index = 0, "염료", "조제") & "를 계속 추가하시겠습니까 ?") Then
                    Call cmdAddNew_Click(Index)
                Else
                    cmdSave.SetFocus
                End If
            End If
        End If
    End With
End Sub

Private Sub grdDyeAux_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdDyeAux(Index)
        Select Case Col
            Case 2
                Cancel = True
            Case 3
                If Len(.TextMatrix(Row, Col)) = 0 Then .TextMatrix(Row, Col) = "0.000000"
                .Cell(flexcpText, Row, Col) = Format(.TextMatrix(Row, Col), "###0.000000")
        End Select
    End With
End Sub

Private Sub grdProcess_DblClick()
Dim iRow%, iCol%, iCurRow%

    With grdCardPattern
        If grdProcess.Rows > grdProcess.FixedRows And grdProcess.Row >= grdProcess.FixedRows _
                        And grdCardPattern.Cell(flexcpBackColor, grdCardPattern.Row, 1) = 0 Then
            iCurRow = .Row
            If .Rows = .FixedRows Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "1"
                .TextMatrix(.Rows - 1, 1) = grdProcess.TextMatrix(grdProcess.Row, 1)
                .TextMatrix(.Rows - 1, 2) = grdProcess.TextMatrix(grdProcess.Row, 2)
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .Row = 1
            Else
                .Rows = .Rows + 1
                For iRow = .Rows - 2 To .Row Step -1
                    For iCol = 0 To .Cols - 1
                        If iCol = 0 Then
                            .TextMatrix(iRow + 1, iCol) = CStr(CInt(.TextMatrix(iRow, iCol)) + 1)
                        Else
                            .TextMatrix(iRow + 1, iCol) = .TextMatrix(iRow, iCol)
                        End If
                    Next iCol
                Next iRow
                If iCurRow = 1 Then
                    .TextMatrix(iCurRow, 0) = "1"
                Else
                    .TextMatrix(iCurRow, 0) = CStr(CInt(.TextMatrix(iCurRow - 1, 0)) + 1)
                End If
                
                .TextMatrix(iCurRow, 1) = grdProcess.TextMatrix(grdProcess.Row, 1)
                .TextMatrix(iCurRow, 2) = grdProcess.TextMatrix(grdProcess.Row, 2)
                .TextMatrix(iCurRow, 3) = ""
                .TextMatrix(iCurRow, 4) = ""
                .TextMatrix(iCurRow, 5) = ""
                .TextMatrix(iCurRow, 6) = ""
            End If
        End If
    End With

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

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkOrderID.Caption = "Order No"
        grdHold.ColWidth(3) = 0
        grdHold.ColWidth(4) = 1300
    Else
        chkOrderID.Caption = "관리번호"
        grdHold.ColWidth(3) = 1300
        grdHold.ColWidth(4) = 0
    End If
End Sub

Private Sub cmdSave_Click()
    If SaveData() Then
        If optPrn(0).Value = True Then
            Call cmdPrint_Click
        End If
        Call MessageBox("저장 되었습니다.")
        
        Call InitGrid
        Call ClearData
        Call FillgrdHold
    End If
End Sub

Private Function CheckData() As Boolean
Dim sOrderID$
Dim iRow%, iCntData%, i%

    CheckData = False
    gSelCnt = 0
    If Trim(txtProcOpinion) = "" Then
        MsgBox "처리 방안을 입력하지 않았습니다", vbInformation, "처리방안 작성"
        txtProcOpinion.SetFocus
        Exit Function
    End If
    
    With grdHold
        For iRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, iRow, 1) = flexChecked Then
                If gSelCnt = 0 Then
                    sOrderID = Trim(.TextMatrix(iRow, 22))
                End If
                If sOrderID <> Trim(.TextMatrix(iRow, 22)) Then
                    Call MessageBox("같은 관리번호만 선택해야 합니다")
                    Exit Function
                End If
                gSelCnt = gSelCnt + 1
                sOrderID = Trim(.TextMatrix(iRow, 22))
            End If
        Next iRow
    End With
    
    
    If gSelCnt = 0 Then
        Call MessageBox("카드를 적어도 하나는 체크선택해야 합니다")
        Exit Function
    End If
    
    If Len(txtBox(0).Tag) <= 0 Then
        Call MessageBox("카드를 선택해야 합니다")
        Exit Function
    End If
    If Len(txtBox(1).Tag) <= 0 Then
        Call MessageBox("카드를 선택해야 합니다")
        Exit Function
    End If
    
    With grdCardPattern
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpBackColor, i, 1) = 0 Then
                iCntData = iCntData + 1
            End If
        Next i
    End With
    If iCntData = 0 Then
        Call MessageBox("계획공정이 지정되지 않았습니다")
        Exit Function
    End If
        
    
    If Len(txtBox(2)) <> 10 Then
        Call MessageBox("처방전번호를 정확히 입력하십시오.")
        txtBox(2).SetFocus
        Exit Function
    End If
    If Len(txtBox(3).Tag) = 0 Then
        Call MessageBox("'처방자'를 입력하십시오.")
        txtBox(3).SetFocus
        Exit Function
    End If

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
    Dim TOpinion() As THoldOpinion
    Dim oCard As PlusLib2.CCard
    Dim tItemSub() As PlusLib2.TPlanPattern
    Dim TRec      As PlusLib2.TRecipe
    Dim tRecSub() As PlusLib2.TRecipeSub
    Dim oRecipe   As PlusLib2.CRecipe
    Dim i%, nDyeCnt%, nRecSub%
    Dim sOrder$, nOrderSeq%
    Dim iCntData%
    Dim bFind As Boolean
    Dim iBaseRow%, idx%
    Dim sAfterProc$
    Dim nQty As Long
    Dim sCardList$


    SaveData = False
    
    If Not CheckData Then Exit Function

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    oCard.UserName = g_sUserName
    
    ReDim TOpinion(gSelCnt - 1)
    
    With grdHold
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                TOpinion(idx).nAffected = 0
                TOpinion(idx).WriteDate = .TextMatrix(i, 24)
                TOpinion(idx).WriteProcID = .TextMatrix(i, 25)
                TOpinion(idx).WriteSeq = CInt("0" & .TextMatrix(i, 26))
                TOpinion(idx).ProcOpinion = Trim(txtProcOpinion.Text)
                TOpinion(idx).ProcDate = MakeDate(DF_SHORT, dtpProcDate)
                TOpinion(idx).ProcPerson = g_sUserName
                TOpinion(idx).CardID = .TextMatrix(i, 20)
                TOpinion(idx).SplitID = .TextMatrix(i, 21)
                
                nQty = nQty + CLng("0" & .TextMatrix(i, 10))
                sCardList = sCardList & .TextMatrix(i, 3) & ", "
                idx = idx + 1
            End If
        Next i
    End With
    
    
    With grdCardPattern
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 3) <> "*" Then
                sAfterProc = sAfterProc & .TextMatrix(i, 2) & "→"
            End If
        Next i
        sAfterProc = Left(sAfterProc, Len(sAfterProc) - 1)
    End With
    
    With grdCardPattern
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpBackColor, i, 1) = 0 Then
                iCntData = iCntData + 1
                If bFind = False Then
                    iBaseRow = i
                End If
                bFind = True
            End If
        Next i
        
        ReDim tItemSub(iCntData - 1)
        
        idx = 0
        For i = iBaseRow To .Rows - 1
            tItemSub(idx).sProcessID = .TextMatrix(i, 1)
            tItemSub(idx).sCompleteClss = ""
            tItemSub(idx).nNeedWidth = 0
            tItemSub(idx).sInstRemark = ""
            tItemSub(idx).sRemark = ""
            
            idx = idx + 1
        Next i
    End With
    
    With TRec
        .OrderID = txtBox(0).Tag
        .OrderSeq = txtBox(1).Tag
        .RecipeSeq = 1     ' 재처방
        .ModifySeq = 1     ' 변경순위
        .RecipeNO = txtBox(2)
        .RecipeDate = MakeDate(DF_SHORT, dtpRecipe)
        .PersonID = txtBox(3).Tag
        .UnitWght = IIf(IsNumeric(txtBox(5)), txtBox(5), 0)
        .ChunkRate = IIf(IsNumeric(txtBox(6)), txtBox(6), 0)
        .ModiClss = "*"     ' 수정 처방 구분
        .Qty = nQty
        .Remark = sCardList & Trim(txtRemark)
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
            tRecSub(i).ModifySeq = 1
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
                tRecSub(i + nDyeCnt).ModifySeq = 1
                tRecSub(i + nDyeCnt).DyeAuxSeq = i + nDyeCnt + 1
                tRecSub(i + nDyeCnt).DyeAuxID = .TextMatrix(.FixedRows + i, 4)
                tRecSub(i + nDyeCnt).DyeAuxRate = CDbl(.TextMatrix(.FixedRows + i, 3))
            Next i
        End If
    End With

    SaveData = oCard.UpdateCardProcANDRecipe(TOpinion, TRec, tRecSub, tItemSub, sAfterProc)

    Set oCard = Nothing
    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    SaveData = False
    Set oCard = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

Private Sub cmdPrint_Click()
    On Error GoTo ErrHandler
    Dim Rs As ADODB.Recordset
    Dim sParam() As String
    Dim oRecipe As PlusLib2.CRecipe
    Dim sRecipeNO$
    Dim i%, nCnt%


'    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    sRecipeNO = txtBox(2)
    
  
    Set oRecipe = New PlusLib2.CRecipe
    oRecipe.Connection = g_adoCon

    Set Rs = oRecipe.GetRecipeOne(sRecipeNO)
    
    Set oRecipe = Nothing
    
    nCnt = 0
    
    ReDim Preserve sParam(40)
    
    For i = 0 To 40
        sParam(i) = " "
    Next i
    
    ' 염료 처방내역
    With grdDyeAux(0)
    
        For i = 1 To .Rows - 1
            sParam(i - 1) = .TextMatrix(i, 1)
            sParam(i + 9) = .TextMatrix(i, 3)
            nCnt = nCnt + 1
        Next i
    
    End With
    
    
    With grdDyeAux(1)
        For i = 1 To .Rows - 1
            sParam(i + 19) = .TextMatrix(i, 1)
            sParam(i + 29) = .TextMatrix(i, 3)
            nCnt = nCnt + 1
            
        Next i
    
    End With
       
   sParam(40) = "작업 처방전(수정)"
    
    Call PrintReport(REPORTFILE, Rs, sParam, False)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set Rs = Nothing
    Set oRecipe = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub






