VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInstRapid_NEW 
   ClientHeight    =   9465
   ClientLeft      =   2985
   ClientTop       =   3225
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   15285
   Begin Threed.SSFrame pnlPattern 
      Height          =   8295
      Left            =   3600
      TabIndex        =   63
      Top             =   720
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14631
      _Version        =   196609
      BackColor       =   12632256
      PictureMaskColor=   14737632
      Begin Threed.SSPanel pnlName 
         Height          =   375
         Index           =   18
         Left            =   90
         TabIndex        =   76
         Top             =   150
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   661
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수정작업을 위한 작업 패턴을 지정하여주십시오."
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   675
         Left            =   6840
         TabIndex        =   75
         Top             =   7470
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1191
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   675
         Left            =   5370
         TabIndex        =   74
         Top             =   7470
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1191
         _Version        =   196609
         Caption         =   "저장"
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   15
         Left            =   90
         TabIndex        =   64
         Top             =   570
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "현재 공정 패턴"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdProcess 
         Height          =   6435
         Left            =   5850
         TabIndex        =   65
         Top             =   930
         Width           =   2325
         _cx             =   4101
         _cy             =   11351
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
      Begin VSFlex7LCtl.VSFlexGrid grdNewPattern 
         Height          =   6435
         Left            =   1650
         TabIndex        =   66
         Top             =   930
         Width           =   2955
         _cx             =   5212
         _cy             =   11351
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12648447
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
      Begin VSFlex7LCtl.VSFlexGrid grdCardPattern 
         Height          =   6435
         Left            =   90
         TabIndex        =   67
         Top             =   930
         Width           =   1485
         _cx             =   2619
         _cy             =   11351
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
         Height          =   315
         Index           =   16
         Left            =   1650
         TabIndex        =   68
         Top             =   570
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "새 공정 패턴"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   795
         Left            =   4800
         TabIndex        =   69
         Top             =   4890
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "삭제"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   795
         Left            =   4800
         TabIndex        =   70
         Top             =   3960
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "추가"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdUP 
         Height          =   795
         Left            =   4800
         TabIndex        =   71
         Top             =   1800
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "위"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   795
         Left            =   4800
         TabIndex        =   72
         Top             =   2700
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         _Version        =   196609
         Caption         =   "아래"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   17
         Left            =   5850
         TabIndex        =   73
         Top             =   570
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "공 정 명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         FillColor       =   &H00FF0000&
         Height          =   8175
         Left            =   60
         Top             =   60
         Width           =   8175
      End
   End
   Begin Threed.SSPanel pnlButton 
      Height          =   615
      Left            =   60
      TabIndex        =   57
      Top             =   8790
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1085
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdCardDivide 
         Height          =   495
         Left            =   1290
         TabIndex        =   58
         Top             =   60
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "카드분리"
      End
      Begin Threed.SSCommand cmdCardChange 
         Height          =   495
         Left            =   75
         TabIndex        =   59
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "카드변경"
      End
   End
   Begin VB.TextBox txtResultDT 
      Height          =   315
      Left            =   720
      TabIndex        =   51
      Top             =   8880
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtRapidSeq 
      Height          =   315
      Left            =   2280
      TabIndex        =   50
      Top             =   8880
      Visible         =   0   'False
      Width           =   1365
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   3165
      Left            =   60
      TabIndex        =   13
      Top             =   5610
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   5583
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtProcID 
         Enabled         =   0   'False
         Height          =   360
         Left            =   11340
         TabIndex        =   55
         Top             =   690
         Width           =   2445
      End
      Begin VB.ComboBox cboHold 
         Enabled         =   0   'False
         Height          =   300
         Left            =   11340
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   2790
         Width           =   2145
      End
      Begin VB.TextBox txtRoll 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   11340
         TabIndex        =   32
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox txtRemarkResult 
         Height          =   315
         Left            =   11340
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2460
         Width           =   3315
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   13800
         TabIndex        =   30
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox txtCardID 
         Height          =   360
         Left            =   11340
         TabIndex        =   29
         Top             =   1050
         Width           =   2445
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
         Index           =   6
         ItemData        =   "frmInstRapid_NEW.frx":0000
         Left            =   8700
         List            =   "frmInstRapid_NEW.frx":0002
         TabIndex        =   28
         Tag             =   "염색패턴"
         Top             =   690
         Width           =   1365
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
         Index           =   7
         Left            =   3840
         TabIndex        =   27
         Tag             =   "염색구분"
         Top             =   690
         Width           =   1545
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
         Index           =   8
         Left            =   5400
         TabIndex        =   26
         Tag             =   "염색구분"
         Top             =   690
         Width           =   1365
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
         Index           =   9
         Left            =   6780
         TabIndex        =   25
         Tag             =   "작업자"
         Top             =   690
         Width           =   1095
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
         Index           =   10
         Left            =   7890
         TabIndex        =   24
         Tag             =   "작업자"
         Top             =   690
         Width           =   795
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
         Index           =   0
         Left            =   30
         TabIndex        =   23
         Tag             =   "작업자"
         Top             =   690
         Width           =   795
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
         Left            =   840
         TabIndex        =   22
         Tag             =   "염색구분"
         Top             =   690
         Width           =   2985
      End
      Begin Threed.SSPanel pnlTitle 
         Height          =   315
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   16761024
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "염색 작업 일지 작성"
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   2
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작업구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   5
         Left            =   6780
         TabIndex        =   16
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작업자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   8
         Left            =   5400
         TabIndex        =   17
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "염색구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   9
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "염색패턴"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   11
         Left            =   7890
         TabIndex        =   19
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작업 조"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   20
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "호기"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   7
         Left            =   8700
         TabIndex        =   21
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "선택된카드"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   1
         Left            =   10140
         TabIndex        =   34
         Top             =   1440
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "절수"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   4
         Left            =   12600
         TabIndex        =   35
         Top             =   1440
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "수량"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   10
         Left            =   10140
         TabIndex        =   36
         Top             =   2460
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "비고사항"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   315
         Index           =   6
         Left            =   10140
         TabIndex        =   37
         Top             =   2790
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "작 업 조"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkHold 
            Caption         =   "보류처리"
            Height          =   180
            Left            =   60
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   60
            Width           =   1020
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   3
         Left            =   10140
         TabIndex        =   39
         Top             =   1800
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "시작일시"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   6
         Left            =   10140
         TabIndex        =   40
         Top             =   2130
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkEnd 
            Caption         =   "종료일시"
            Height          =   180
            Left            =   60
            TabIndex        =   41
            Top             =   60
            Width           =   1035
         End
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Index           =   0
         Left            =   11340
         TabIndex        =   42
         Top             =   2130
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyy년 MM월 dd일 (ddd)"
         Format          =   120848387
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Index           =   0
         Left            =   11340
         TabIndex        =   43
         Top             =   1800
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyy년 MM월 dd일 (ddd)"
         Format          =   120848387
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Index           =   1
         Left            =   13530
         TabIndex        =   44
         Top             =   2130
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH시 MM분"
         Format          =   120848386
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Index           =   1
         Left            =   13530
         TabIndex        =   45
         Top             =   1800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH시 MM분"
         Format          =   120848386
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   12
         Left            =   10140
         TabIndex        =   46
         Top             =   1080
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "카드번호"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   13
         Left            =   10140
         TabIndex        =   52
         Top             =   390
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "스케줄번호"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   14
         Left            =   10140
         TabIndex        =   56
         Top             =   720
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "공정명"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdCard 
         Height          =   495
         Left            =   13980
         TabIndex        =   62
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "카드입력"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " ~ "
         Height          =   180
         Left            =   11670
         TabIndex        =   53
         Top             =   450
         Width           =   255
      End
      Begin VB.Label lblSchID 
         Alignment       =   1  '오른쪽 맞춤
         BorderStyle     =   1  '단일 고정
         Height          =   285
         Left            =   11340
         TabIndex        =   48
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lblSchSeq 
         AutoSize        =   -1  'True
         BorderStyle     =   1  '단일 고정
         Caption         =   "         "
         Height          =   270
         Left            =   11940
         TabIndex        =   47
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "다시읽어오기"
      Height          =   435
      Left            =   13320
      TabIndex        =   9
      Top             =   0
      Width           =   1905
   End
   Begin TabDlg.SSTab grdTab 
      Height          =   5565
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   9816
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   626
      TabMaxWidth     =   4410
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "염색공정 대기"
      TabPicture(0)   =   "frmInstRapid_NEW.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pnlWaitTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdList(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "염색대기카드"
      TabPicture(1)   =   "frmInstRapid_NEW.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList(1)"
      Tab(1).Control(1)=   "pnlWaitTab(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "염색진행중카드"
      TabPicture(2)   =   "frmInstRapid_NEW.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdList(2)"
      Tab(2).Control(1)=   "pnlWaitTab(2)"
      Tab(2).Control(2)=   "Label4"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "수정작업 처리"
      TabPicture(3)   =   "frmInstRapid_NEW.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pnlWaitTab(3)"
      Tab(3).Control(1)=   "grdList(3)"
      Tab(3).ControlCount=   2
      Begin VSFlex7LCtl.VSFlexGrid grdList 
         Height          =   5070
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   396
         Width           =   15060
         _cx             =   26564
         _cy             =   8943
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
      Begin VSFlex7LCtl.VSFlexGrid grdList 
         Height          =   5070
         Index           =   1
         Left            =   -74940
         TabIndex        =   2
         Top             =   396
         Width           =   15060
         _cx             =   26564
         _cy             =   8943
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
      Begin VSFlex7LCtl.VSFlexGrid grdList 
         Height          =   5070
         Index           =   2
         Left            =   -74940
         TabIndex        =   3
         Top             =   396
         Width           =   15060
         _cx             =   26564
         _cy             =   8943
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
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   36
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   12259610
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "염색공정 대기"
         BevelOuter      =   0
         FloodColor      =   12259610
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   345
         Index           =   1
         Left            =   -72420
         TabIndex        =   5
         Top             =   36
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   12539970
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "염색대기카드"
         BevelOuter      =   0
         FloodColor      =   12539970
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   345
         Index           =   2
         Left            =   -69870
         TabIndex        =   6
         Top             =   36
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   14389120
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "염색진행중카드"
         BevelOuter      =   0
         FloodColor      =   14389120
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   345
         Index           =   3
         Left            =   -67320
         TabIndex        =   60
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   14389120
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수정작업처리"
         BevelOuter      =   0
         FloodColor      =   14389120
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdList 
         Height          =   5070
         Index           =   3
         Left            =   -74940
         TabIndex        =   61
         Top             =   390
         Width           =   15060
         _cx             =   26564
         _cy             =   8943
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "■  수정/추가에만 선택하십시오."
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   -64770
         TabIndex        =   12
         Top             =   66
         Width           =   2640
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   615
      Left            =   8670
      TabIndex        =   7
      Top             =   8790
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1085
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCommand cmdCancelStart 
         Height          =   495
         Left            =   2445
         TabIndex        =   8
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "작업취소"
      End
      Begin Threed.SSCommand cmdWorkStart 
         Height          =   495
         Left            =   90
         TabIndex        =   10
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "작업시작"
      End
      Begin Threed.SSCommand cmdWorkEnd 
         Height          =   495
         Left            =   1275
         TabIndex        =   11
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "작업완료"
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   5190
         TabIndex        =   49
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   900
         _Version        =   196609
         Caption         =   "      닫기(&X)"
         PictureAlignment=   1
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   495
         Left            =   3630
         TabIndex        =   54
         Top             =   60
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "추가작업"
      End
   End
End
Attribute VB_Name = "frmInstRapid_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bEnableWork As Boolean
Private Const Custom = "태을염직"   ' Rapid: 11(1~11), CPB: 1(12)
Private Const nMachNo = 12      ' 12호기 추가
Private nMachineID As Integer
Private m_sReWorkClss As String

Private Sub chkEnd_Click()
    If chkEnd.Value = vbChecked Then
        dtpEndDate(0).Enabled = True
        dtpEndDate(1).Enabled = True
        dtpEndDate(0) = Now
        dtpEndDate(1) = Now
        cmdWorkEnd.Enabled = True
        cmdWorkStart.Enabled = False
    Else
        dtpEndDate(0).Enabled = False
        dtpEndDate(1).Enabled = False
        cmdWorkEnd.Enabled = False
        cmdWorkStart.Enabled = True
    End If
End Sub

Private Sub chkHold_Click()
    If chkHold.Value = 1 Then
        cboHold.Enabled = True
    Else
        cboHold.Enabled = False
    End If
End Sub



Private Sub cmdButton_Click(Index As Integer)
'''    Dim oRapid As PlusLib2.CRapid
'''
'''    Select Case Index
''''        Case 0: '카드분리
''''            Dim sRs As Recordset
''''
''''            If Trim(pnlCardID) <> "" Then
''''                Set oRapid = New PlusLib2.CRapid
''''                oRapid.Connection = g_adoCon
''''                oRapid.UserName = g_sUserName
''''
''''                Set sRs = oRapid.GetCheckDyeSch(Trim(pnlCardID), Trim(pnlSplitID))
''''                Set oRapid = Nothing
''''
''''                If sRs.RecordCount > 0 Then
''''                    If Trim(sRs!Complitclss) = "" Then
''''                        MsgBox "염색작업지시가 내려진 카드는 카드분리를 할수 없습니다", vbInformation, "카드분리 불가"
''''                        Exit Sub
''''                    End If
''''                End If
''''                Set sRs = Nothing
''''                frmCardDivide.chkSearch(4).Value = vbChecked
''''                frmCardDivide.txtSearch(4).Text = Select_TabRow_No("카드번호")
''''                Call frmCardDivide.cmdSearch_Click
''''            End If
''''        Case 1: '색상변경
''''            frmCardChange.chkSearch(4).Value = vbChecked
''''            frmCardChange.txtSearch(4).Text = Select_TabRow_No("카드번호")
''''            Call frmCardChange.cmdSearch_Click
'''        Case 2: '처방조회
'''            frmRecipeView.optOrder(1).Value = True
'''            frmRecipeView.chkSearch(3).Value = vbUnchecked
'''            frmRecipeView.chkSearch(2).Value = vbChecked
'''            frmRecipeView.tabMain.Tab = 0
''''            If shpBox.Visible = True Then   ' 스케쥴에 근거한 관리번호
''''                frmRecipeView.txtSearch(2).Text = Select_TabRow_No("관리번호")
''''            Else
''''                If grdTab(0).TextMatrix(grdTab(0).Row, 1) = "실적" Then
''''                    frmRecipeView.txtSearch(2).Text = Select_TabRow_No("관리번호")
''''                Else            ' 카드에 근거한 관리번호
''''                    frmRecipeView.txtSearch(2).Text = lblOrderID
''''                End If
''''            End If
'''            Call frmRecipeView.FillGridRecipe
'''
'''        Case 3: '평량지시
'''            Dim sSchIDSeq As String
'''            Dim rs As Recordset
'''
''''            If shpBox.Visible = False Then
''''                MsgBox "염색지시건을 선택해야 합니다", vbInformation, "선택 요구"
''''                Exit Sub
''''            End If
''''            sSchIDSeq = Select_TabRow_No("스케쥴")
'''
'''            Set oRapid = New PlusLib2.CRapid
'''            oRapid.Connection = g_adoCon
'''            oRapid.UserName = g_sUserName
'''
'''            Set rs = oRapid.GetCheckDyeWorking(CLng(Left(sSchIDSeq, 9)), CInt(Right(sSchIDSeq, 2)))
'''            Set oRapid = Nothing
'''
'''            If rs.RecordCount > 0 Then
'''                If (Trim(rs!UseClss) = "작업" Or Len(Trim(rs!UseClss)) = 8) And Left(rs!procid, 2) = "43" Then
'''                    Set rs = Nothing
'''                    MsgBox "선택되어진 건은 현재 작업중입니다" & vbCrLf & vbCrLf & "평량지시를 내릴수 없습니다", vbCritical, "편집 불가"
'''                    Exit Sub
'''                End If
'''            End If
'''            Set rs = Nothing
'''            Call frmRecipeCalc.SetInstruction(CLng(Left(sSchIDSeq, 9)), CInt(Right(sSchIDSeq, 2)))
'''        Case 4: '수주상세
'''            frmOrderHistory.optOrder(0).Value = True
'''
''''            If shpBox.Visible = True Then   ' 스케쥴에 근거한 관리번호
''''                frmOrderHistory.txtSearch.Text = Select_TabRow_No("관리번호")
''''            Else                            ' 카드에 근거한 관리번호
''''                frmOrderHistory.txtSearch.Text = lblOrderID
''''            End If
'''
'''            frmOrderHistory.txtSearch_KeyPress (vbKeyReturn)
''''        Case 5: '카드상세
''''            frmCardHistory.txtCard.Text = Select_TabRow_No("카드번호")
''''            frmCardHistory.txtCard_KeyPress (vbKeyReturn)
'''        Case 6: '염색일지 조회
'''            frmDyeResultView.dtpDate(0) = Now:   frmDyeResultView.dtpDate(1) = Now
'''            Call frmDyeResultView.cmdSearch_Click
'''        Case 7: '염색패턴
'''            frmDyePattern.Show 1
'''        Case 8: '패턴변경
'''            If Trim(pnlCardID) <> "" And pnlCardID <> "카드번호" Then
'''                frmCardPattern.chkSearch(4).Value = vbChecked
'''                frmCardPattern.txtSearch(4).Text = pnlCardID
'''                frmCardPattern.txtSearch(5).Text = pnlSplitID
'''                frmCardPattern.cmdSearch_Click
'''            Else
'''                MsgBox "카드 선택하고 버튼을 눌러주십시요", vbInformation, "카드 선택 요망"
'''                Exit Sub
'''            End If
'''        Case 9: '색상변경
'''            If pnlCardID = "카드번호" Or Trim(pnlCardID) = "" Then
'''                MsgBox "카드를 선택해야 합니다", vbInformation, "카드선택 요망"
'''                Exit Sub
'''            End If
'''            If cboColor.ListIndex < 0 Then
'''                MsgBox "색상을 선택해야합니다", vbInformation, "색상선택 요망"
'''                Exit Sub
'''            End If
'''
'''            Set oRapid = New PlusLib2.CRapid
'''            oRapid.Connection = g_adoCon
'''            oRapid.UserName = g_sUserName
'''
'''            If oRapid.UpdateCardColor(pnlCardID, pnlSplitID, cboColor.ItemData(cboColor.ListIndex), g_sUserName) Then
'''                MsgBox "카드의 칼라를 변경했습니다", vbOKOnly, "칼라 변경"
'''            End If
'''            Set oRapid = Nothing
'''          '  Call FillGridData
'''          '  Call FillSchData
'''    End Select
End Sub




Private Sub cmdAdd_Click()
    Dim i%
    
    With grdNewPattern
        If .Row < .Rows - 1 Then
            If .TextMatrix(.Row + 1, 3) = "*" Then
                MsgBox "작업이 완료된 공정안에는 공정을 추가할 수 없습니다.", vbInformation + vbOKOnly
                Exit Sub
            End If
        End If
        
        .Redraw = flexRDNone
        
        .AddItem "" & vbTab & grdProcess.TextMatrix(grdProcess.Row, 1) & vbTab & grdProcess.TextMatrix(grdProcess.Row, 2) & vbTab & "" & vbTab & 0, .Row + 1
        .Cell(flexcpForeColor, .Row + 1, 1, .Row + 1, .Cols - 1) = vbBlue
        
        .Cell(flexcpChecked, .Row + 1, 7) = flexChecked
        .Select .Row + 1, 1
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdCancel_Click()
    pnlPattern.Visible = False
    grdList(3).RemoveItem grdList(3).Rows - 1
End Sub

Private Sub cmdCancelStart_Click()
    Dim oRapid As PlusLib2.CRapid_NEW
    Dim iCol As Integer
    Dim nSchID As Long
    Dim nSeq As Integer
    
    If grdTab.Tab = 2 Or grdTab.Tab = 3 Then
    
        If MsgBox("작업중인 건을 취소하시겠습니까?", vbQuestion + vbYesNo, "취소 여부") = vbYes Then
        
            Set oRapid = New PlusLib2.CRapid_NEW
            oRapid.Connection = g_adoCon
            oRapid.UserName = g_sUserName
        
            Screen.MousePointer = vbHourglass
            
            With grdList(grdTab.Tab)
            
            If oRapid.SetRapiddCancel(.TextMatrix(.Row, 17), .TextMatrix(.Row, 18), .TextMatrix(.Row, 19), .TextMatrix(.Row, 20)) Then
                MsgBox "작업이 취소되었습니다", vbOKOnly, "취소 성공"
            Else
                MsgBox "작업 취소중 오류 발생...", vbOKOnly, "취소 실패"
            
            End If
            
            End With
            
            Set oRapid = Nothing
            Call cmdRefresh_Click
            Screen.MousePointer = vbDefault
        End If
        Call cmdRefresh_Click
    End If
    
End Sub


Private Function SaveData(ByVal JobMode As String, ByVal JoBStr As String) As Boolean
    Dim oRapid As PlusLib2.CRapid_NEW
    Dim TWkRapid As PlusLib2.TWkRapid
    Dim TWkRapidSUB() As PlusLib2.TWkRapidSUB
    Dim JJ As Integer, II As Integer, nCount As Integer
    Dim vSchId As Variant
    Dim iChkHold As Integer
    Dim sHoldReason As String
    Dim dMsg_Str As String
    
    SaveData = False
  '  If Not CheckWorkEnd() Then Exit Sub
  
    If lstArray(6).ListCount = 0 Then Exit Function
    
  
    If m_sReWorkClss = "수정" Or m_sReWorkClss = "재염" Then
        If lstArray(8).ListIndex < 2 Then
            MsgBox "재염, 수정작업시에는 염색구분을 본염, 추가를 선택할 수 없습니다.", vbOKOnly
            Exit Function
        End If
    Else
        If lstArray(8).ListIndex > 1 Then
            MsgBox "본작업시에는 염색구분을 수정을 선택할 수 없습니다.", vbOKOnly
            Exit Function
        End If
    End If
  
    If chkEnd.Value = vbChecked Then
        dMsg_Str = lstArray(7) & vbCrLf & vbCrLf & JoBStr & " 하시겠습니까?"
    Else
        dMsg_Str = lstArray(7) & vbCrLf & vbCrLf & JoBStr & " 하시겠습니까?"
    
    End If

    If MsgBox(dMsg_Str, vbYesNo + vbQuestion, "염색일지 저장 여부") = vbNo Then
        Exit Function
    End If

    Set oRapid = New PlusLib2.CRapid_NEW
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

    If chkHold.Value = 0 Then
        iChkHold = 0
        sHoldReason = ""
    Else
        iChkHold = 1
        sHoldReason = Trim(cboHold.Text)
    End If
    
    If JobMode = "U" And chkEnd.Value = vbUnchecked Then
        chkEnd.Value = vbChecked
        dtpEndDate(0) = Now
        dtpEndDate(1) = Now
    End If

    With TWkRapid
        .wkResultDT = IIf(val(txtRapidSeq) = 0, MakeDate(DF_SHORT, dtpStartDate(0)), txtResultDT)
        If lstArray(7).Text = "염색" Or lstArray(7).Text = "정련" Then
            .wkProcID = txtProcID.Tag
            .wkRapidSeq = val(txtRapidSeq)
        Else
            .wkProcID = "4300"
            .wkRapidSeq = val(txtRapidSeq)
        End If
        .wkMachID = Left(lstArray(0), 2)
        .WorkClss = lstArray(7).Text
        .RapidClss = lstArray(8).Text
        .PatternID = Left(lstArray(1), 2)
        .InMethod = ""
        .UnitWght = 0
        .WkRoll = val(txtRoll)
        .WkQty = val(txtQty)
        .TeamID = Format(Left(lstArray(10), 1), "0#")
        .PersonID = Right(lstArray(9), 8)
        .StartDate = MakeDate(DF_SHORT, dtpStartDate(0))
        .StartTime = Mid(Format(dtpStartDate(1), "YYYYMMDDhhmm"), 9, 4)
        .EndDate = IIf(chkEnd = vbChecked, MakeDate(DF_SHORT, dtpEndDate(0)), "")
        .EndTime = IIf(chkEnd = vbChecked, Mid(Format(dtpEndDate(1), "YYYYMMDDhhmm"), 9, 4), "")
        .DyeSchID = val(lblSchID)
        .DyeSeq = val(lblSchSeq)
        .Remark = ""
        .HoldReason = ""
    End With



    ' 염색구분이 염색이 아닌 경우 Card없이 작업 함.

    JJ = 0
    If lstArray(7).Text = "염색" Or lstArray(7).Text = "정련" Then
        With lstArray(6)
            For II = 0 To .ListCount - 1
                JJ = JJ + 1
                ReDim Preserve TWkRapidSUB(JJ)

                TWkRapidSUB(JJ - 1).wkResultDT = MakeDate(DF_SHORT, dtpStartDate(0))
                TWkRapidSUB(JJ - 1).wkProcID = txtProcID.Tag
                TWkRapidSUB(JJ - 1).wkMachID = Left(lstArray(0), 2)
                TWkRapidSUB(JJ - 1).wkRapidSeq = 1

                vSchId = Split(.List(II), "-")
                If UBound(vSchId) = 1 Then
                    TWkRapidSUB(JJ - 1).CardID = vSchId(0)
                    TWkRapidSUB(JJ - 1).SplitID = vSchId(1)
                Else
                    TWkRapidSUB(JJ - 1).CardID = vSchId(0)
                    TWkRapidSUB(JJ - 1).SplitID = ""
                End If
                If lstArray(8).ListIndex > 1 Then
                    TWkRapidSUB(JJ - 1).ReWorkClss = "*"
                    TWkRapidSUB(JJ - 1).ReWorkID = lstArray(8).ItemData(lstArray(8).ListIndex)
                End If
            Next II
        End With
    Else
        ReDim TWkRapidSUB(1)
        TWkRapidSUB(0).wkResultDT = MakeDate(DF_SHORT, dtpStartDate(0))
        TWkRapidSUB(0).wkProcID = "4300"
        TWkRapidSUB(0).wkMachID = Left(lstArray(0), 2)
        TWkRapidSUB(0).wkRapidSeq = 0
        TWkRapidSUB(0).CardID = ""
        TWkRapidSUB(0).SplitID = ""
        JJ = 1
    End If


    
        If oRapid.InsertRapid(TWkRapid, TWkRapidSUB, JJ, JobMode) Then
            Set oRapid = Nothing
            If chkEnd.Value = vbChecked Then
                dMsg_Str = lstArray(7) & vbCrLf & vbCrLf & "작업이 완료(작성) 되었습니다."
            Else
                dMsg_Str = lstArray(7) & vbCrLf & vbCrLf & "작업이 시작 되었습니다."
            
            End If
            MsgBox dMsg_Str, vbOKOnly, "작성 성공"
            chkHold.Value = 0
            lstArray(6).Clear
            SaveData = True
        Else
            Set oRapid = Nothing
        End If

End Function

Private Function CheckWorkEnd() As Boolean
Dim iCount%
    
    
    If lstArray(1).ListCount > 0 Then
        If lstArray(1).SelCount = 0 Then
            MsgBox "염색패턴이 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    End If
    If lstArray(7).SelCount = 0 Then
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    If lstArray(6).ListCount = 0 And val(txtRoll) Then
        MsgBox "염색에서는 카드를 반드시 선택하십시오.", vbCritical, "작성 오류"
        Exit Function
    End If
    
    If lstArray(7).ListIndex > 0 Then
        If lstArray(8).SelCount > 0 Then
            MsgBox "염색구분이 선택되면 안됩니다", vbCritical, "작성 오류"
            Exit Function
        End If
    ElseIf lstArray(8).ListIndex = 0 Then
        If lstArray(8).SelCount = 0 Then
            MsgBox "염색구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
            Exit Function
        End If
    Else
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    If lstArray(9).SelCount = 0 Then
        MsgBox "작업자가 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    If lstArray(10).SelCount = 0 Then
        MsgBox "작업조가 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    
    If MakeDate(DF_SHORT, dtpStartDate(0)) & Mid(Format(dtpStartDate(1), "YYYYMMDDhhmm"), 9, 4) _
             < MakeDate(DF_SHORT, dtpEndDate(0)) & Mid(Format(dtpEndDate(1), "YYYYMMDDhhmm"), 9, 4) Then
        MsgBox "작업시작일시와 종료일시가 잘못 되었습니다", vbCritical, "작성 오류"
        Exit Function
             
    End If
    
    CheckWorkEnd = True
End Function

Private Sub cmdCard_Click()
    Dim stemp$, sCardID$, sSplitID$
    stemp = InputBox("수정작업을 할 카드번호를 입력하십시오", "수정작업")
    
    If stemp = "" Then Exit Sub
    
    stemp = Replace(stemp, "-", "")
    stemp = Replace(stemp, "(", "")
    stemp = Replace(stemp, ")", "")
    sCardID = Left(stemp, 8)
    sSplitID = Mid(stemp, 9, Len(stemp) - 8)
    
    If Len(sCardID) = 8 Then
        Call FillGridData_ReWork(sCardID, sSplitID)
        
        If grdList(3).Rows = grdList(3).FixedRows Then
            Exit Sub
        End If
        
        Call FillGridProcess
        Call FillGridPattern
        pnlPattern.Visible = True
    End If
End Sub

Private Sub cmdCardChange_Click()
    frmCardChange.chkSearch(4).Value = vbChecked
    frmCardChange.txtSearch(4).Text = grdList(grdTab.Tab).TextMatrix(grdList(grdTab.Tab).Row, 17)
    Call frmCardChange.cmdSearch_Click
End Sub

Private Sub cmdDel_Click()
    Dim i%
    
    With grdNewPattern
        If .TextMatrix(.Row, 3) = "*" Then
            MsgBox "작업이 완료된 공정은 공정을 삭제할 수 없습니다.", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        .RemoveItem .Row
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub cmdDown_Click()
    Dim i%
    Dim vTemp(6) As String
    
    With grdNewPattern
        If .Row = .Rows - 1 Then Exit Sub
        
        If .TextMatrix(.Row, 3) = "*" Then
            MsgBox "작업이 완료된 공정에 대해서는 순서를 변경할 수 없습니다.", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        For i = 1 To 6
            vTemp(i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = .TextMatrix(.Row + 1, i)
            .TextMatrix(.Row + 1, i) = vTemp(i)
        Next i
        .Select .Row + 1, 1
    End With

End Sub

Private Sub cmdRefresh_Click()
    Call AddLstBox
    
    If grdTab.Tab = 0 Then
        Call InitGrid(0)
        Call FillGridData_NOW(0)
    ElseIf grdTab.Tab = 1 Then
        Call InitGrid(1)
        Call FillGridData_NOW(1)
    ElseIf grdTab.Tab = 2 Then
        Call InitINGGrid(2)
        Call FillGridData_ING(2)
    ElseIf grdTab.Tab = 3 Then
        Call InitGrid(3)
        
    End If
    Call SetEditClear(0)
    If grdList(0).Rows < grdList(0).FixedRows Then
        grdList(0).SetFocus
    End If
'    grdTab.Tab = 0
End Sub



Private Sub LoadRapidWorkData(SchID As Long, Seq As Integer)
Dim oRapid As PlusLib2.CRapid
Dim rs As Recordset
Dim tRs As Recordset
Dim i%, iCount%
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName


    Set rs = oRapid.GetwiRapidData(SchID, Seq)

    If rs.RecordCount > 0 Then
        txtRoll = rs!wiRoll
        txtQty = Format(rs!wiQty, "###,##0")
        ' 염색패턴
        Set tRs = oRapid.GetDyePatternList(1, CInt(rs!wiMachID), 0)
        
        lstArray(6).Clear
        For iCount = 1 To tRs.RecordCount
            lstArray(6).AddItem Format(tRs!PtNo, "00") & ". " & tRs!PtName
            tRs.MoveNext
        Next iCount
        tRs.Close
        Set tRs = Nothing
        
        For i = 0 To lstArray(6).ListCount - 1
            If Left(lstArray(6).List(i), 2) = Format(CInt(rs!PatternID), "00") Then
                lstArray(6).Selected(i) = True
                Exit For
            End If
        Next i
        
        ' 작업구분
        For i = 0 To lstArray(7).ListCount - 1
            If lstArray(7).List(i) = rs!WorkClss Then
                lstArray(7).Selected(i) = True
                Exit For
            End If
        Next i
        ' 염색구분
        For i = 0 To lstArray(8).ListCount - 1
            If lstArray(8).List(i) = rs!RapidClss Then
                lstArray(8).Selected(i) = True
                Exit For
            End If
        Next i
        ' 작업자
        For i = 0 To lstArray(9).ListCount - 1
            If Right(lstArray(9).List(i), 8) = Format(rs!PersonID, "00000000") Then
                lstArray(9).Selected(i) = True
                Exit For
            End If
        Next i
        ' 작업조
        For i = 0 To lstArray(10).ListCount - 1
            If Left(lstArray(10).List(i), 1) = CStr(CInt(rs!TeamID)) Then
                lstArray(10).Selected(i) = True
                Exit For
            End If
        Next i
        
    End If
    rs.Close
    Set rs = Nothing
    Set oRapid = Nothing
''
''    txtEndDate = Format(Now, "YYYYMMDD")
''    txtEndTime = Format(time, "HHMM")
End Sub

Private Sub cmdSave_Click()
    Dim sCardID$, sSplitID$
    If UpdateCardPattern Then
        sCardID = MakeCardID(grdList(3).TextMatrix(grdList(3).Rows - 1, 9), OM_REDUCE)
        sSplitID = grdList(3).TextMatrix(grdList(3).Rows - 1, 10)
        Call FillGridData_ReWork(sCardID, sSplitID)
        pnlPattern.Visible = False
    End If
End Sub

Private Sub cmdUP_Click()
    Dim i%
    Dim vTemp(6) As String
    
    With grdNewPattern
        If .Row <= .FixedRows Then Exit Sub
        
        If .TextMatrix(.Row, 3) = "*" Or .TextMatrix(.Row - 1, 3) = "*" Then
            MsgBox "작업이 완료된 공정에 대해서는 순서를 변경할 수 없습니다.", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        For i = 1 To 6
            vTemp(i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = .TextMatrix(.Row - 1, i)
            .TextMatrix(.Row - 1, i) = vTemp(i)
        Next i
        .Select .Row - 1, 1
    End With

End Sub

Private Sub cmdWorkEnd_Click()
    If SaveData("U", cmdWorkEnd.Caption) Then Call cmdRefresh_Click
End Sub

Private Sub cmdWorkStart_Click()

    If SaveData("I", cmdWorkStart.Caption) Then Call cmdRefresh_Click
End Sub

Private Sub Form_Activate()
    m_sReWorkClss = ""
    PlusMDI.pnlMenu.Visible = False
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub
Private Sub SetButtonStatus()
    cmdWorkStart.Enabled = True
    cmdWorkEnd.Enabled = True
    cmdCancelStart.Enabled = True
    
    Select Case grdTab.Tab
        Case 0   '염색대기 카드
            cmdCancelStart.Enabled = False
            pnlEdit.Enabled = True
            pnlButton.Visible = False
            lstArray(0).Enabled = True
            lstArray(6).Enabled = True
        Case 1
            cmdCancelStart.Enabled = False
            pnlEdit.Enabled = True
            pnlButton.Visible = True
            lstArray(0).Enabled = True
            lstArray(6).Enabled = True
        Case 2
            cmdWorkStart.Enabled = False
            pnlEdit.Enabled = True
            pnlButton.Visible = False
            lstArray(0).Enabled = False
            lstArray(6).Enabled = False
        Case 3
            cmdCancelStart.Enabled = False
            pnlEdit.Enabled = True
            pnlButton.Visible = True
            lstArray(0).Enabled = True
            lstArray(6).Enabled = True
    End Select
End Sub
Private Sub Form_Load()
    Dim i%
    
    bEnableWork = True
    
     Me.Move 0, 0, 15360, 9840
    
    Call AddLstBox
        
    Call InitGrid(0)
    Call InitGrid(1)
    Call InitINGGrid(2)
    Call InitGrid(3)
    Call InitGridProcess
    
    Call FillGridData_NOW(0)
'    Call FillGridData_NOW(1)
'    Call FillGridData_ING(2)
'    Call FillGridData_ING(3)
    
    dtpStartDate(0) = Now
    dtpStartDate(1) = "15:13"
    
    dtpEndDate(0) = Now
    dtpEndDate(1) = Now
    Call SetEditClear(0)
    If grdList(0).Rows >= grdList(0).FixedRows Then
'        grdList(0).SetFocus
    End If
'    Call FillGridData_end
    
    nMachineID = 0
    pnlButton.Visible = False
    
    cmdAdd.Picture = LoadResPicture("BACK", vbResIcon)
    cmdDel.Picture = LoadResPicture("FRONT", vbResIcon)
    cmdUP.Picture = LoadResPicture("UP", vbResIcon)
    cmdDown.Picture = LoadResPicture("DOWN", vbResIcon)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub

Private Function CheckData() As Boolean
    Dim irow%, iCol%, iCount%, iChkCnt%
    
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
    If lstArray(5).SelCount = 0 Then
        MsgBox "작업구분이 선택되어 있지 않습니다", vbCritical, "작성 오류"
        Exit Function
    End If
    If lstArray(5).ListIndex > 0 Then
        If lstArray(2).SelCount > 0 Then
            MsgBox "염색구분이 선택되면 안됩니다", vbCritical, "작성 오류"
            Exit Function
        End If
    ElseIf lstArray(5).ListIndex = 0 Then
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
        
'    With grdList(4)
'        For iRow = 1 To .Rows - 2
'            If .Cell(flexcpChecked, iRow, 0, iRow, 0) = flexChecked Then
'                iCount = iCount + 1
'            End If
'            If .TextMatrix(iRow, 7) = "미확정" Then
'                iChkCnt = iChkCnt + 1
'            End If
'        Next iRow
'    End With
'
'    If iChkCnt > 0 Then
'        MsgBox "색상이 미확정인 카드는 염색지시를 내릴수 없습니다", vbCritical, "작성 오류"
'        Exit Function
'    End If
    
    CheckData = True
End Function

''Private Function AddData(TotRoll As Long, TotQty As Long) As Boolean
''    Dim oRapid As PlusLib2.CRapid
''    Dim tCardList() As PlusLib2.tRapidCard
''    Dim i%, iCount%, iCntChk%, iCol%, iRow%, iSeq%
''
''    Screen.MousePointer = vbHourglass
''    AddData = False
''
''    On Error GoTo ErrHandler
''
''    Set oRapid = New PlusLib2.CRapid
''    oRapid.Connection = g_adoCon
''    oRapid.UserName = g_sUserName
''
''    With grdList(4)
''        For i = .FixedRows To .Rows - 2
''            If .Cell(flexcpChecked, i, 0) = flexChecked Then
''                iCntChk = iCntChk + 1
''            End If
''        Next i
''        If lstArray(5).ListIndex > 0 Then
''            ReDim tCardList(1)
''            tCardList(iCount).sCardID = ""
''            tCardList(iCount).sSplitID = ""
''            tCardList(iCount).lDyeSchID = 0
''        Else
''            ReDim tCardList(iCntChk)
''            iCount = 0
''            For i = .FixedRows To .Rows - 2
''                If .Cell(flexcpChecked, i, 0) = flexChecked Then
''                    tCardList(iCount).sCardID = Trim(.TextMatrix(i, 17))
''                    tCardList(iCount).sSplitID = IIf(Trim(.TextMatrix(i, 18)) = "", " ", Trim(.TextMatrix(i, 18)))
''                    If lstArray(2).Text = "추가" Then
''                        tCardList(iCount).lDyeSchID = CLng(.TextMatrix(i, 23))
''                    Else
''                        tCardList(iCount).lDyeSchID = 0
''                    End If
''                    iCount = iCount + 1
''                End If
''            Next i
''        End If
''
''    End With
''
''    g_adoCon.BeginTrans
''
''    If Not oRapid.AddNewwiRapidItem(tCardList(), CLng(tCardList(0).lDyeSchID), "4300", Left(lstArray(0).Text, 2), _
''        0, lstArray(5).Text, lstArray(2).Text, Format(CInt("0" & Left(lstArray(1).Text, 2)), "000"), 0, TotRoll, _
''        TotQty, " ", Right(lstArray(3).Text, 8), CheckNull(txtRemark)) Then
''        Set oRapid = Nothing
''        AddData = False
''        Exit Function
''    End If
''
''    AddData = True
''    g_adoCon.CommitTrans
''
''    Set oRapid = Nothing
''
''    Screen.MousePointer = vbDefault
''
''    Exit Function
''
''ErrHandler:
''    Screen.MousePointer = vbDefault
''    AddData = False
''
''    Set oRapid = Nothing
''    Call ErrorBox(Err.Number, "frminstRapid.AddData", Err.Description)
''End Function


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub AddLstBox()
    Dim oRapid As PlusLib2.CRapid
    Dim oPerson  As PlusLib2.CPerson
    Dim rs As Recordset
    Dim iCount%, i%, iSeq%
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    

    cboHold.Clear
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
                
    ' 염색호기
    lstArray(0).Clear
    Set rs = oRapid.GetMachineNoList("Rapid")
    For iCount = 1 To rs.RecordCount
        lstArray(0).AddItem Format(rs!MachineNO, "00") & " 호기"
        rs.MoveNext
    Next iCount
    
    lstArray(0).ListIndex = 0
    rs.Close
    Set rs = Nothing
    
    Set oRapid = Nothing
    
' 진호염직의 염색구분 목록
    
    lstArray(8).Clear
    lstArray(8).AddItem "본염"
    lstArray(8).AddItem "추가"
    lstArray(8).AddItem "얼룩수정"
    lstArray(8).ItemData(2) = 11
    lstArray(8).AddItem "오염수정"
    lstArray(8).ItemData(3) = 12
    lstArray(8).AddItem "색수정"
    lstArray(8).ItemData(4) = 13
    lstArray(8).AddItem "시와수정"
    lstArray(8).ItemData(5) = 14
    lstArray(8).ListIndex = 0
   
   
    cboHold.AddItem "추가"
    cboHold.AddItem "얼룩수정"
    cboHold.AddItem "오염수정"
    cboHold.AddItem "색수정"
    cboHold.AddItem "시와수정"
      
    ' 진호염직의 작업구분
    lstArray(7).Clear
    
    lstArray(7).AddItem "염색"
    lstArray(7).AddItem "정련"
    lstArray(7).AddItem "BOX 탈색"
    lstArray(7).AddItem "BOX R/C"
    lstArray(7).AddItem "도포 Washing"
    lstArray(7).AddItem "Soaping"
    lstArray(7).AddItem "기계수리"
    lstArray(7).ListIndex = 0
    
    lstArray(9).Clear

    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    oPerson.UserName = g_sUserName
    Set rs = oPerson.GetWorkerList("05")     ' 작업자 ( 염색( '13' )
    For iCount = 1 To rs.RecordCount
        lstArray(9).AddItem rs!Name & Space(20) & Format(rs!PersonID, "00000000")
        rs.MoveNext
    Next iCount
    rs.Close
    Set rs = Nothing
    lstArray(9).ListIndex = 0
    
    lstArray(10).Clear

    ' 작업조
    Set rs = oPerson.GetWorkTeam()     '작업 조
    For iCount = 1 To rs.RecordCount
        lstArray(10).AddItem CStr(CInt(rs!TeamID)) & ". " & rs!Team
        rs.MoveNext
    Next iCount
    lstArray(10).ListIndex = 0
    rs.Close
    Set rs = Nothing
    
    Set oPerson = Nothing
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    Set oPerson = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.AddLstBox", Err.Description)
End Sub




'현재진행중인 작업
Private Sub InitINGGrid(ByVal Index As Integer)
' 호기, 작업조, 작업자, 거래처, 품명, 색상, 시작시각, 절수, 수량, 염색패턴, 작업구분, 염색구분, 스케즐번호, 스케즐차수, RESULTDT, WKPROCID, MACHID, RAPIDSEQ

    Call SetVSFlexGrid(grdList(Index))

    With grdList(Index)
        .WordWrap = False
        .Redraw = flexRDNone

        .Rows = 1
        .Cols = 25
        .RowHeight(0) = 300
        .FixedRows = 1
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightNever
        .ExplorerBar = flexExNone
        .FocusRect = flexFocusSolid
        
        
        .TextArray(0) = "":                      .ColWidth(0) = 300:           .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "호기":                  .ColWidth(1) = 600:         .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "작업조":                .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "작업자":                .ColWidth(3) = 1000:        .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처":                .ColWidth(4) = 2000:        .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "품명":                  .ColWidth(5) = 2300:        .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "색상":                  .ColWidth(6) = 2000:        .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "시작시각":              .ColWidth(7) = 1600:        .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "절수":                  .ColWidth(8) = 500:         .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "수량":                  .ColWidth(9) = 800:         .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "염색패턴":             .ColWidth(10) = 1000:       .ColAlignment(10) = flexAlignLeftCenter
        .TextArray(11) = "작업구분":             .ColWidth(11) = 1000:       .ColAlignment(11) = flexAlignLeftCenter
        .TextArray(12) = "염색구분":             .ColWidth(12) = 1000:       .ColAlignment(12) = flexAlignLeftCenter
        .TextArray(13) = "관리번호":             .ColWidth(13) = 1000:       .ColAlignment(13) = flexAlignLeftCenter
        .TextArray(14) = "OrderNO":              .ColWidth(14) = 0:          .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "DyeSchID":             .ColWidth(15) = 0:          .ColAlignment(15) = flexAlignLeftCenter
        .TextArray(16) = "DyeSeq":               .ColWidth(16) = 0:          .ColAlignment(16) = flexAlignLeftCenter
        .TextArray(17) = "ResultDT":             .ColWidth(17) = 0:          .ColAlignment(17) = flexAlignCenterCenter
        .TextArray(18) = "ProcID":               .ColWidth(18) = 0:          .ColAlignment(18) = flexAlignLeftCenter
        .TextArray(19) = "MachID":               .ColWidth(19) = 0:          .ColAlignment(19) = flexAlignLeftCenter
        .TextArray(20) = "RapidSeq":             .ColWidth(20) = 0:          .ColAlignment(20) = flexAlignLeftCenter
        .TextArray(21) = "종료일시":             .ColWidth(21) = 0:          .ColAlignment(21) = flexAlignLeftCenter
        .TextArray(22) = "공정명":               .ColWidth(22) = 0:          .ColAlignment(22) = flexAlignLeftCenter
        .TextArray(23) = "재작업구분":               .ColWidth(23) = 0:          .ColAlignment(23) = flexAlignLeftCenter
        .TextArray(24) = "재작업종류":               .ColWidth(24) = 0:          .ColAlignment(24) = flexAlignLeftCenter
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub InitGridProcess()
    With grdCardPattern
        .Redraw = flexRDNone
        .Cols = 7
        
        Call SetVSFlexGrid(grdCardPattern)
        .Rows = 1
        
        .TextArray(0) = "순서":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColHidden(1) = True
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "완료여부":     .ColHidden(3) = True
        .TextArray(4) = "요구폭":       .ColHidden(4) = True
        .TextArray(5) = "지시사항":     .ColHidden(5) = True
        .TextArray(6) = "비고":         .ColHidden(6) = True
        
        .HighLight = flexHighlightNever
        .Redraw = flexRDDirect
    End With
    
    With grdNewPattern
        .Redraw = flexRDNone
        .Cols = 8
        
        Call SetVSFlexGrid(grdNewPattern)
        .Rows = 1
        
        .TextArray(0) = "순서":         .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColWidth(1) = 0
        .TextArray(2) = "공정명":       .ColWidth(2) = 1600:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "완료여부":     .ColWidth(3) = 0
        .TextArray(4) = "요구폭":       .ColWidth(4) = "0":       .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "지시사항":     .ColWidth(5) = "0":      .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "비고":         .ColWidth(6) = "0":      .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "재작업":       .ColWidth(7) = "400":       .ColAlignment(7) = flexAlignLeftCenter:     .ColDataType(7) = flexDTBoolean

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusSolid
        .ScrollBars = flexScrollBarBoth
        .HighLight = flexHighlightAlways
        .ExtendLastCol = True
        
        .Redraw = flexRDDirect
    End With
    
    With grdProcess
        .Redraw = flexRDNone
        .Cols = 3
        
        Call SetVSFlexGrid(grdProcess)
        .Rows = 1
        
        .TextArray(0) = "":             .ColWidth(0) = 500:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "공정코드":     .ColWidth(1) = 0
        .TextArray(2) = "공정명":       .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignLeftCenter
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub InitGrid(ByVal i As Integer)
    
    Call SetVSFlexGrid(grdList(i))

    With grdList(i)
        .WordWrap = False
        .Redraw = flexRDNone

        .Rows = 1:      .Cols = 29
        .RowHeight(0) = 300
        .RowHeightMin = 400
        .FixedRows = 1:     .FixedCols = 0
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightNever
        .ExplorerBar = flexExNone
        .FocusRect = flexFocusSolid
        
        .FixedCols = 0
        
        .TextArray(0) = "":                     .ColWidth(0) = 250:         .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "밧자번호":             .ColWidth(1) = 0:           .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "밧자순위":             .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = " ":                 .ColWidth(3) = 300:         .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "No":                   .ColWidth(4) = 300:         .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "거래처":               .ColWidth(5) = 1100:        .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "품명":                 .ColWidth(6) = 2500:        .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "색상":                 .ColWidth(7) = 2000:        .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "관리번호":             .ColWidth(8) = 1200:        .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "카드번호":             .ColWidth(9) = 1000:        .ColAlignment(9) = flexAlignLeftCenter
        .TextArray(10) = "분할":                .ColWidth(10) = 500:        .ColAlignment(10) = flexAlignLeftCenter
        .TextArray(11) = "대기":                .ColWidth(11) = 800:        .ColAlignment(11) = flexAlignLeftCenter
        .TextArray(12) = "절수":                .ColWidth(12) = 600:        .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "수량":                .ColWidth(13) = 800:        .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "거래처코드":          .ColWidth(14) = 0:          .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "품명코드":            .ColWidth(15) = 0:          .ColAlignment(15) = flexAlignLeftCenter
        .TextArray(16) = "색상코드":            .ColWidth(16) = 0:          .ColAlignment(16) = flexAlignLeftCenter
        .TextArray(17) = "카드번호":            .ColWidth(17) = 0:          .ColAlignment(17) = flexAlignCenterCenter
        .TextArray(18) = "분할":                .ColWidth(18) = 0:          .ColAlignment(18) = flexAlignLeftCenter
        .TextArray(19) = "제직처":              .ColWidth(19) = 900:        .ColAlignment(19) = flexAlignLeftCenter
        .TextArray(20) = "관리번호":            .ColWidth(20) = 0:          .ColAlignment(20) = flexAlignLeftCenter
        .TextArray(21) = "OrderSeq":            .ColWidth(21) = 0:          .ColAlignment(21) = flexAlignLeftCenter
        .TextArray(22) = "이후공정":            .ColWidth(22) = 2000:       .ColAlignment(22) = flexAlignLeftCenter
        .TextArray(23) = "스케쥴번호":          .ColWidth(23) = 0:          .ColAlignment(23) = flexAlignLeftCenter
        .TextArray(24) = "차수":                .ColWidth(24) = 0:          .ColAlignment(24) = flexAlignLeftCenter
        .TextArray(25) = "WaitProcID":          .ColWidth(25) = 0:          .ColAlignment(25) = flexAlignLeftCenter
        .TextArray(26) = "재작업구분":          .ColWidth(26) = 0:          .ColAlignment(26) = flexAlignLeftCenter
        .TextArray(27) = "재작업종류":          .ColWidth(27) = 0:          .ColAlignment(27) = flexAlignLeftCenter
        .TextArray(28) = "공정순서":            .ColWidth(28) = 0:          .ColAlignment(28) = flexAlignLeftCenter
        
        Select Case i
            Case 0
                .ColDataType(0) = flexDTBoolean
        End Select
        
        .ScrollBars = flexScrollBarBoth
        .WordWrap = True
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub FillGridData_ING(ByVal Index As Integer)
    Dim oRapid As PlusLib2.CRapid_NEW
    Dim rs As Recordset
    Dim iCount%, k%, iSeq%, iNowRow%
    Dim sWorkUnitID$
    Dim bToggle As Boolean
    Dim sCustom$, sArticle$, sCheck As String
    
    
    Screen.MousePointer = vbHourglass

'    On Error GoTo ErrHandler
        
    Set oRapid = New PlusLib2.CRapid_NEW
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
    
    If Index = 2 Then
        sCheck = "0"
    Else
        sCheck = "1"
    End If
    

    Set rs = oRapid.GetRapidScheduling_ING(sCheck)
    Set oRapid = Nothing

    bToggle = False
    With grdList(Index)
        .Redraw = flexRDNone
        
        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        
        .Rows = .FixedRows
        .Redraw = flexRDDirect
        
        For iCount = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .Row = .Rows - 1
            .Col = 0
            
            .TextMatrix(.Rows - 1, 0) = IIf(Trim(rs!ReWorkClss) = "*", "■", "")
            .TextMatrix(.Rows - 1, 1) = rs!wkMachID
            .TextMatrix(.Rows - 1, 2) = rs!Team
            .TextMatrix(.Rows - 1, 3) = rs!Worker
            .TextMatrix(.Rows - 1, 4) = rs!kCustom
            .TextMatrix(.Rows - 1, 5) = rs!Article
            .TextMatrix(.Rows - 1, 6) = Trim(rs!Color)
            .TextMatrix(.Rows - 1, 7) = MakeDate(DF_LONG, rs!StartDate) & " " & Format(rs!StartTime, "00:00")
            .TextMatrix(.Rows - 1, 8) = rs!WkRoll
            .TextMatrix(.Rows - 1, 9) = rs!WkQty
            .TextMatrix(.Rows - 1, 10) = rs!PatternID
            .TextMatrix(.Rows - 1, 11) = rs!WorkClss
            .TextMatrix(.Rows - 1, 12) = rs!RapidClss
            .TextMatrix(.Rows - 1, 13) = MakeOrderID(rs!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 14) = rs!OrderNo
            .TextMatrix(.Rows - 1, 15) = rs!DyeSchID
            .TextMatrix(.Rows - 1, 16) = rs!DyeSeq
            .TextMatrix(.Rows - 1, 17) = rs!wkResultDT
            .TextMatrix(.Rows - 1, 18) = rs!wkProcID
            .TextMatrix(.Rows - 1, 19) = rs!wkMachID
            .TextMatrix(.Rows - 1, 20) = rs!wkRapidSeq
            .TextMatrix(.Rows - 1, 21) = rs!EndDate & rs!EndTime
            .TextMatrix(.Rows - 1, 22) = rs!Process
            .TextMatrix(.Rows - 1, 23) = rs!ReWorkClss
            .TextMatrix(.Rows - 1, 24) = rs!ReWorkID
            
            iSeq = iSeq + 1
            rs.MoveNext
        Next iCount
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If
        
        Dim SchID As String, bColor As Long, II As Integer
        bColor = &HFFC0C0
    End With
    Call wkRapid_ING(Index)
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.FillGridData_NOW", Err.Description)


End Sub

Private Sub FillGridData_NOW(ByVal i As Integer)
    Dim oRapid As PlusLib2.CRapid_NEW
    Dim rs As Recordset
    Dim iCount%, k%, iSeq%, iNowRow%
    Dim sWorkUnitID$
    Dim bToggle As Boolean
    Dim sCustom$, sArticle$
    
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
        
    Set oRapid = New PlusLib2.CRapid_NEW
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

'    If i = 3 Then
'        Set rs = oRapid.GetRapidScheduling_END
'    Else
        Set rs = oRapid.GetRapidScheduling_NOW(i)
'    End If
    
    Set oRapid = Nothing

    bToggle = False
    With grdList(i)
        .Redraw = flexRDNone
        
        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        
        .Rows = .FixedRows
        For iCount = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .Row = .Rows - 1
            .Col = 0
            If i < 3 Then
                If rs!SchID > 0 Then
                    .CellChecked = flexNoCheckbox
                Else
                    .CellChecked = flexUnchecked
                End If
            Else
                .CellChecked = flexUnchecked
            End If
            If iCount = 1 Then
                sWorkUnitID = rs!WorkUnitId
                iSeq = 0
            End If
            If sWorkUnitID <> rs!WorkUnitId Then
                bToggle = Not (bToggle)
                iSeq = 0
            End If
            .TextMatrix(.Rows - 1, 1) = rs!WorkUnitId
            .TextMatrix(.Rows - 1, 2) = rs!WorkUnitSeq
            .TextMatrix(.Rows - 1, 3) = IIf(Trim(rs!ReWorkClss) = "*", "■", "")
            .TextMatrix(.Rows - 1, 4) = CStr(iSeq + 1)
            .TextMatrix(.Rows - 1, 5) = Trim(rs!kCustom)
            .TextMatrix(.Rows - 1, 6) = Trim(rs!Article)
            .TextMatrix(.Rows - 1, 7) = Trim(rs!Color)
            .TextMatrix(.Rows - 1, 8) = MakeOrderID(rs!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 9) = Format(rs!CardID, "00-00-0000")
            .TextMatrix(.Rows - 1, 10) = rs!SplitID
            .TextMatrix(.Rows - 1, 11) = rs!WaitProc
            .TextMatrix(.Rows - 1, 12) = Format(rs!Roll, "#,##0")
            .TextMatrix(.Rows - 1, 13) = Format(rs!Qty, "#,###,##0")
            .TextMatrix(.Rows - 1, 14) = rs!CustomID
            .TextMatrix(.Rows - 1, 15) = rs!ArticleID
            .TextMatrix(.Rows - 1, 16) = rs!ColorID
            .TextMatrix(.Rows - 1, 17) = rs!CardID
            .TextMatrix(.Rows - 1, 18) = rs!SplitID
            .TextMatrix(.Rows - 1, 19) = rs!Custom
            .TextMatrix(.Rows - 1, 20) = rs!OrderID
            .TextMatrix(.Rows - 1, 21) = rs!OrderSeq
            .TextMatrix(.Rows - 1, 22) = rs!AfterProc
            .TextMatrix(.Rows - 1, 23) = rs!SchID
            .TextMatrix(.Rows - 1, 24) = rs!DyeSeq
            .TextMatrix(.Rows - 1, 25) = rs!waitprocid
            .TextMatrix(.Rows - 1, 26) = rs!ReWorkClss
            .TextMatrix(.Rows - 1, 27) = rs!ReWorkID
            .TextMatrix(.Rows - 1, 28) = rs!WaitProcSeq
           
            sWorkUnitID = rs!WorkUnitId
            
            iSeq = iSeq + 1
            rs.MoveNext
        Next iCount
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If
        
        Dim SchID As String, bColor As Long, II As Integer
        bColor = &HFFC0C0
        If i = 0 Then
            SchID = ""
            
            For II = .FixedRows To .Rows - 1
                If SchID = "" Then
                    SchID = Trim(.TextMatrix(II, 23)) & Trim(.TextMatrix(II, 24))
                    bColor = &HFFC0C0
                    
                ElseIf SchID <> Trim(.TextMatrix(II, 23)) & Trim(.TextMatrix(II, 24)) Then
                        SchID = Trim(.TextMatrix(II, 23)) & Trim(.TextMatrix(II, 24))
                        If bColor = &HFFC0C0 Then
                            bColor = 0
                        Else
                            bColor = &HFFC0C0
                        End If
                End If
                .Cell(flexcpBackColor, II, 0, II, .Cols - 1) = bColor
                
            Next II
        End If
        
        .Redraw = flexRDDirect
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.FillGridData_NOW", Err.Description)

End Sub

Private Sub FillGridData_ReWork(sCardID As String, sSplitID As String)
    Dim oRapid As PlusLib2.CRapid_NEW
    Dim rs As Recordset
    Dim iCount%, k%, iSeq%, iNowRow%
    Dim sWorkUnitID$
    Dim bToggle As Boolean
    Dim sCustom$, sArticle$
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
        
    Set oRapid = New PlusLib2.CRapid_NEW
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName

    Set rs = oRapid.GetRapidScheduling_ReWork(sCardID, sSplitID)
   
    Set oRapid = Nothing

    bToggle = False
    With grdList(3)
        .Redraw = flexRDNone
        
        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        
        .Rows = .FixedRows
        For iCount = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .Row = .Rows - 1
            .Col = 0
            .CellChecked = flexUnchecked

            If iCount = 1 Then
                sWorkUnitID = rs!WorkUnitId
                iSeq = 0
            End If
            If sWorkUnitID <> rs!WorkUnitId Then
                bToggle = Not (bToggle)
                iSeq = 0
            End If
            .TextMatrix(.Rows - 1, 1) = rs!WorkUnitId
            .TextMatrix(.Rows - 1, 2) = rs!WorkUnitSeq
            .TextMatrix(.Rows - 1, 3) = "■"
            .TextMatrix(.Rows - 1, 4) = CStr(iSeq + 1)
            .TextMatrix(.Rows - 1, 5) = Trim(rs!kCustom)
            .TextMatrix(.Rows - 1, 6) = Trim(rs!Article)
            .TextMatrix(.Rows - 1, 7) = Trim(rs!Color)
            .TextMatrix(.Rows - 1, 8) = MakeOrderID(rs!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 9) = Format(rs!CardID, "00-00-0000")
            .TextMatrix(.Rows - 1, 10) = rs!SplitID
            .TextMatrix(.Rows - 1, 11) = rs!WaitProc
            .TextMatrix(.Rows - 1, 12) = Format(rs!Roll, "#,##0")
            .TextMatrix(.Rows - 1, 13) = Format(rs!Qty, "#,###,##0")
            .TextMatrix(.Rows - 1, 14) = rs!CustomID
            .TextMatrix(.Rows - 1, 15) = rs!ArticleID
            .TextMatrix(.Rows - 1, 16) = rs!ColorID
            .TextMatrix(.Rows - 1, 17) = rs!CardID
            .TextMatrix(.Rows - 1, 18) = rs!SplitID
            .TextMatrix(.Rows - 1, 19) = rs!Custom
            .TextMatrix(.Rows - 1, 20) = rs!OrderID
            .TextMatrix(.Rows - 1, 21) = rs!OrderSeq
            .TextMatrix(.Rows - 1, 22) = rs!AfterProc
            .TextMatrix(.Rows - 1, 23) = rs!SchID
            .TextMatrix(.Rows - 1, 24) = rs!DyeSeq
            .TextMatrix(.Rows - 1, 25) = rs!waitprocid
            .TextMatrix(.Rows - 1, 26) = "*"
            .TextMatrix(.Rows - 1, 27) = ""
            .TextMatrix(.Rows - 1, 28) = rs!WaitProcSeq
           
            sWorkUnitID = rs!WorkUnitId
            
            iSeq = iSeq + 1
            rs.MoveNext
        Next iCount
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If
        .Redraw = flexRDDirect
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Set rs = Nothing
    Set oRapid = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid.FillGridData_ReWork", Err.Description)

End Sub

''Private Sub FillGridData()
''    Dim oRapid As PlusLib2.CRapid
''    Dim rs As Recordset
''    Dim iCount%, i%, k%, iSeq%, iNowRow%
''    Dim sWorkUnitID$
''    Dim bToggle As Boolean
''    Dim sCustom$, sArticle$
''
''
''    Screen.MousePointer = vbHourglass
''
''    On Error GoTo ErrHandler
''
''    For i = 0 To 3
''        Set oRapid = New PlusLib2.CRapid
''        oRapid.Connection = g_adoCon
''        oRapid.UserName = g_sUserName
''
''        Set rs = oRapid.GetRapidScheduling(i, 0)
''        Set oRapid = Nothing
''
''
''        bToggle = False
''        With grdList(i)
''            .Redraw = flexRDNone
''
''            iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
''
''            .Rows = .FixedRows
''            For iCount = 1 To rs.RecordCount
''                .Rows = .Rows + 1
''                .RowHeight(.Rows - 1) = 300
''                .Row = .Rows - 1
''                .Col = 0
''                If i < 3 Then
''                    If rs!SchID > 0 Then
''                        .CellChecked = flexNoCheckbox
''                    Else
''                        .CellChecked = flexUnchecked
''                    End If
''                Else
''                    .CellChecked = flexUnchecked
''                End If
''                If iCount = 1 Then
''                    sWorkUnitID = rs!WorkUnitId
''                    iSeq = 0
''                End If
''                If sWorkUnitID <> rs!WorkUnitId Then
''                    bToggle = Not (bToggle)
''                    iSeq = 0
''                End If
''                .TextMatrix(.Rows - 1, 1) = rs!WorkUnitId
''                .TextMatrix(.Rows - 1, 2) = rs!WorkUnitSeq
''                .TextMatrix(.Rows - 1, 3) = "" & rs!BatJaNO
''                .TextMatrix(.Rows - 1, 4) = CStr(iSeq + 1)
''                .TextMatrix(.Rows - 1, 5) = Trim(rs!kCustom)
''                .TextMatrix(.Rows - 1, 6) = Trim(rs!Article)
''                .TextMatrix(.Rows - 1, 7) = Trim(rs!Color)
''                .TextMatrix(.Rows - 1, 8) = MakeOrderID(rs!OrderID, OM_EXPAND)
''                .TextMatrix(.Rows - 1, 9) = Format(rs!CardID, "00-00-0000")
''                .TextMatrix(.Rows - 1, 10) = rs!SplitID
''                .TextMatrix(.Rows - 1, 11) = rs!WaitProc
''                .TextMatrix(.Rows - 1, 12) = Format(rs!Roll, "#,##0")
''                .TextMatrix(.Rows - 1, 13) = Format(rs!Qty, "#,###,##0")
''                .TextMatrix(.Rows - 1, 14) = rs!CustomID
''                .TextMatrix(.Rows - 1, 15) = rs!ArticleID
''                .TextMatrix(.Rows - 1, 16) = rs!ColorID
''                .TextMatrix(.Rows - 1, 17) = rs!CardID
''                .TextMatrix(.Rows - 1, 18) = rs!SplitID
''                .TextMatrix(.Rows - 1, 19) = rs!Custom
''                .TextMatrix(.Rows - 1, 20) = rs!OrderID
''                .TextMatrix(.Rows - 1, 21) = rs!OrderSeq
''                .TextMatrix(.Rows - 1, 22) = rs!AfterProc
''                .TextMatrix(.Rows - 1, 23) = rs!SchID
''                .TextMatrix(.Rows - 1, 24) = rs!DyeSeq
''
''                sWorkUnitID = rs!WorkUnitId
''
''                iSeq = iSeq + 1
''                rs.MoveNext
''            Next iCount
''            rs.Close
''            Set rs = Nothing
''
''            If .Rows > .FixedRows Then
''                .HighLight = flexHighlightAlways
''                .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
''                .TopRow = .Row
''                .Col = .FixedCols
''                .ColSel = .Cols - 1
''            End If
''
''            Dim SchID As String, bColor As Long
''            bColor = &HE0E0E0
''            If i = 0 Then
''                SchID = ""
''
''                For k = .FixedRows To .Rows - 1
''                    If SchID = "" Then
''                        SchID = .TextMatrix(II, 23) & .TextMatrix(II, 24)
''                    ElseIf SchID <> .TextMatrix(II, 23) & .TextMatrix(II, 24) Then
''                            If bColor = &HE0E0E0 Then
''                                bColor = 0
''                            Else
''                                bColor = &HE0E0E0
''                            End If
''                            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = bColor
''                    End If
''                Next k
''
''            End If
''
''            .Redraw = flexRDDirect
''        End With
''    Next i
''
''    Screen.MousePointer = vbDefault
''
''    Exit Sub
''
''ErrHandler:
''    Screen.MousePointer = vbDefault
''
''    Set rs = Nothing
''    Set oRapid = Nothing
''
''    Call ErrorBox(Err.Number, "frmInstRapid.FillGridData", Err.Description)
''End Sub



''Private Sub grdList_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''    If Index = 4 Then
''        Cancel = True
''    Else
''        If Col = 0 Then
''            Cancel = False
''        Else
''            Cancel = True
''        End If
''    End If
''End Sub

Private Sub grdList_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 3
            Call GetRollQty(Index)
            
        Case 2
            Call wkRapid_ING(Index)
    End Select
    
End Sub

Private Sub GetRollQty(ByVal Index As Integer)
    Dim II%, nTotRoll As Integer, nTotQty As Long, SchID As String, SchSeq As String
    Dim bChecked As Boolean
    Dim KK As Integer, JJ As Integer
    Dim sDyeSchID As Integer, nDyeSeq As Integer
    Dim vDyeSchID As Variant, bInDyeSch As Boolean
    Dim nReCnt%, nCnt%
    
    
    With grdList(Index)
        For II = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, II, 0) = flexChecked Then
                nCnt = nCnt + 1
            End If
            If .TextMatrix(II, 26) = "*" And .Cell(flexcpChecked, II, 0) = flexChecked Then
                nReCnt = nReCnt + 1
            End If
        Next II
    End With
    
    If nReCnt > 0 And nCnt <> nReCnt Then
        MsgBox "본작업과 수정작업을 같이할수 없습니다. 확인하여 주십시오", vbOKOnly
        grdList(Index).Cell(flexcpChecked, grdList(Index).Row, 0) = flexUnchecked
        Exit Sub
    End If
    
    nTotRoll = 0: nTotQty = 0
    lstArray(6).Clear
    lblSchID = ""
    lblSchSeq = ""
    dtpStartDate(0) = Now:    dtpStartDate(1) = Now

    If Index = 0 Then
        '염색스케즐이 있는 카드 선택 또는 선택 해제지
           
        With grdList(Index)
            
            If .Cell(flexcpChecked, .Row, 0) = flexChecked Then
                bChecked = True
                SchID = .TextMatrix(.Row, 23)
                SchSeq = .TextMatrix(.Row, 24)
            Else
                bChecked = False
            End If
            
            
            ' Check된것 해제
            For II = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, II, 0) = flexUnchecked
            Next II
            
            ' 염색스케즐에 있는 카드 선택된 것 해지
            With grdList(1)
                For II = .FixedRows To .Rows - 1
                    .Cell(flexcpChecked, II, 0) = flexUnchecked
                Next II
            End With
            
            
            '선택된거  선택처리
            If bChecked Then
                For II = .FixedRows To .Rows - 1
                    If .TextMatrix(II, 23) = SchID And .TextMatrix(II, 24) = SchSeq Then
                        .Cell(flexcpChecked, II, 0) = flexChecked
                        nTotRoll = nTotRoll + .ValueMatrix(II, 12)
                        nTotQty = nTotQty + .ValueMatrix(II, 13)
                        lstArray(6).AddItem MakeCardID(.TextMatrix(II, 9), OM_REDUCE, .TextMatrix(II, 10))
                        lstArray(6).ItemData(lstArray(6).ListCount - 1) = .TextMatrix(II, 28)
                        txtProcID.Text = .TextMatrix(II, 11)
                        txtProcID.Tag = .TextMatrix(II, 25)
                        If .TextMatrix(II, 3) = "■" Then
                            If grdTab.Tab = 3 Then
                                m_sReWorkClss = "수정"
                            Else
                                m_sReWorkClss = "재염"
                            End If
                        Else
                            m_sReWorkClss = ""
                        End If
                    End If
                Next II
            End If
        End With
    Else
        ' 염색스케즐에 있는 카드 선택된 것 해지
        With grdList(0)
            For II = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, II, 0) = flexUnchecked
            Next II
        End With
        
        With grdList(Index)
            For II = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, II, 0) = flexChecked Then
                    nTotRoll = nTotRoll + .ValueMatrix(II, 12)
                    nTotQty = nTotQty + .ValueMatrix(II, 13)
                    lstArray(6).AddItem MakeCardID(.TextMatrix(II, 9), OM_REDUCE, .TextMatrix(II, 10))
                    lstArray(6).ItemData(lstArray(6).ListCount - 1) = .TextMatrix(II, 28)
                    txtProcID.Text = .TextMatrix(II, 11)
                    txtProcID.Tag = .TextMatrix(II, 25)
                    If .TextMatrix(II, 3) = "■" Then
                        If grdTab.Tab = 3 Then
                            m_sReWorkClss = "수정"
                        Else
                            m_sReWorkClss = "재염"
                        End If
                    Else
                        m_sReWorkClss = ""
                    End If
                End If
            Next II
        End With
    End If
    
    
''    For KK = 0 To 1
''        With grdList(KK)
''            SchID = .TextMatrix(.Row, 23)
''            SchSeq = .TextMatrix(.Row, 24)
''
''            If .Cell(flexcpChecked, .Row, 0) = flexChecked Then
''                bChecked = True
''            Else
''                bChecked = False
''            End If
''
''            lblSchID = ""
''            lblSchSeq = ""
''
''            If Index = 0 Then
''                For II = .FixedRows To .Rows - 1
''                    If .TextMatrix(II, 23) = SchID And .TextMatrix(II, 24) = SchSeq Then
''                        If bChecked Then
''                            .Cell(flexcpChecked, II, 0) = flexChecked
''                            lblSchID = .TextMatrix(II, 23)
''                            lblSchSeq = .TextMatrix(II, 24)
''                        Else
''                            .Cell(flexcpChecked, II, 0) = flexUnchecked
''                        End If
''                    End If
''                Next II
''            End If
''
''
''            For II = .FixedRows To .Rows - 1
''                If .Cell(flexcpChecked, II, 0) = flexChecked Then
''                    nTotRoll = nTotRoll + .ValueMatrix(II, 12)
''                    nTotQty = nTotQty + .ValueMatrix(II, 13)
''                    lstArray(6).AddItem MakeCardID(.TextMatrix(II, 9), OM_REDUCE, .TextMatrix(II, 10))
''
''                    ' dyesch 번호가 있는지 확인 한다.
''                    bInDyeSch = False
''                End If
''            Next II
''        End With
''    Next KK
    
    txtRoll = nTotRoll
    txtQty = nTotQty
    lblSchID = SchID
    lblSchSeq = SchSeq
    
    If val(lblSchID) <> 0 Then
        Call SetWiRapid
    Else
        lstArray(7).ListIndex = 0
        chkEnd.Value = vbUnchecked
        cmdWorkEnd.Enabled = False
    End If
    
End Sub
Sub SetWiRapid()
    Dim oRapid As PlusLib2.CRapid_NEW
    Dim rs As Recordset

    Set oRapid = New PlusLib2.CRapid_NEW
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
        
    Set rs = oRapid.GetWiRapidSch(val(lblSchID), val(lblSchSeq))
    If rs.RecordCount = 1 Then
        lstArray(0).ListIndex = FindItem(lstArray(0), rs!wiMachID, 2)
        lstArray(8).ListIndex = FindItem(lstArray(8), rs!RapidClss)
        lstArray(1).ListIndex = FindItem(lstArray(1), Mid(rs!PatternID, 2), 2)
        lstArray(7).ListIndex = FindItem(lstArray(7), rs!WorkClss)
    End If
    rs.Close
    
    Set oRapid = Nothing

End Sub
 Sub SetGrdCheck(ByVal oFlex As VSFlexGrid, ByVal bCheck As Boolean)
    Dim II As Integer
    With oFlex
        If bCheck = True Then
            .Cell(flexcpChecked, II, 0) = flexUnchecked
        Else
            .Cell(flexcpChecked, II, 0) = flexChecked
        End If
    End With
End Sub

Private Sub grdList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Index = 4 Then
        With grdList(Index)
            If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
                .ToolTipText = .TextMatrix(.MouseRow, .Cols - 3)
            End If
        End With
'    End If
End Sub

Private Sub grdList_RowColChange(Index As Integer)
    Select Case Index
        Case 2
           ' Call wkRapid_ING( index)
    End Select
'    Call grdList_Click(Index)
End Sub
Private Sub wkRapid_END()
    Dim II As Integer, Index As Integer
    With grdList(2)
        If .Row < .FixedRows Then
            Exit Sub
        End If
        
        '작업호기
        For II = 0 To lstArray(0).ListCount - 1
            If Left(lstArray(0).List(II), 2) = .TextMatrix(.Row, 1) Then
                lstArray(0).ListIndex = II
                Exit For
            End If
        Next II
        
        '염색패턴
        For II = 0 To lstArray(1).ListCount - 1
            If Left(lstArray(1).List(II), 2) = .TextMatrix(.Row, 10) Then
                lstArray(1).ListIndex = II
                Exit For
            End If
        Next II
        
        '작업구분
        Index = 7
        For II = 0 To lstArray(Index).ListCount - 1
            If lstArray(Index).List(II) = .TextMatrix(.Row, 11) Then
                lstArray(Index).ListIndex = II
                Exit For
            End If
        Next II
        
        '염색구분
        Index = 8
        For II = 0 To lstArray(Index).ListCount - 1
            If lstArray(Index).List(II) = .TextMatrix(.Row, 12) Then
                lstArray(Index).ListIndex = II
                Exit For
            End If
        Next II
        
        '작업자
        Index = 9
        For II = 0 To lstArray(Index).ListCount - 1
            If lstArray(Index).List(II) = .TextMatrix(.Row, 3) Then
                lstArray(Index).ListIndex = II
                Exit For
            End If
        Next II
        
        '작업조
        Index = 10
        For II = 0 To lstArray(Index).ListCount - 1
            If lstArray(Index).List(II) = .TextMatrix(.Row, 2) Then
                lstArray(Index).ListIndex = II
                Exit For
            End If
        Next II
        
        
        txtRoll = .TextMatrix(.Row, 8)
        txtQty = .TextMatrix(.Row, 9)
'        .TextMatrix ( .Row , 7)    'dtpStartDate(0)


        dtpStartDate(0) = Left(.TextMatrix(.Row, 7), 10)
        dtpStartDate(1) = Right(Trim(.TextMatrix(.Row, 7)), 5)     '납기일자
        
        'dtpStartDate(0) = Left(.TextMatrix(.Row, 7), 10)   ' 종료일자
        'dtpStartDate(1) = Mid(.TextMatrix(.Row, 7), 11)    ' 종료시각
        '작업중인 카드
        Dim oRapid As PlusLib2.CRapid_NEW
        Dim rs As Recordset
        
        Set oRapid = New PlusLib2.CRapid_NEW
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
    
        txtResultDT = .TextMatrix(.Row, 17)
        txtRapidSeq = .TextMatrix(.Row, 20)
        
        Set rs = oRapid.GetRapidCardList(Trim(.TextMatrix(.Row, 17)), Trim(.TextMatrix(.Row, 18)), Trim(.TextMatrix(.Row, 19)), Trim(.TextMatrix(.Row, 20)))
        Set oRapid = Nothing
        
        lstArray(6).Clear
        Do Until rs.EOF
            lstArray(6).AddItem rs!CardID & rs!SplitID
            rs.MoveNext
        Loop
        rs.Close
        Set oRapid = Nothing
        
        lstArray(0).Enabled = False
        lstArray(6).Enabled = False
    End With
    pnlEdit.Enabled = False
    
End Sub

Private Sub wkRapid_ING(ByVal Index As Integer)
    Dim II As Integer
    With grdList(Index)
        If .Row < .FixedRows Then
            Exit Sub
        End If
        
        If .TextMatrix(.Row, 0) = "■" Then
            m_sReWorkClss = "수정"
        Else
            m_sReWorkClss = ""
        End If
        lstArray(0).ListIndex = FindItem(lstArray(0), Trim(.TextMatrix(.Row, 1)), 2)     '염색호기
        If Trim(.TextMatrix(.Row, 11)) = "염색" Then
            lstArray(7).ListIndex = FindItem(lstArray(7), Trim(.TextMatrix(.Row, 11)))       '작업구분
            lstArray(8).ListIndex = FindItem(lstArray(8), Trim(.TextMatrix(.Row, 12)))       '염색구분
        Else
            lstArray(7).ListIndex = FindItem(lstArray(7), Trim(.TextMatrix(.Row, 11)))       '작업구분
            lstArray(8).ListIndex = -1
        End If
        lstArray(1).ListIndex = FindItem(lstArray(1), Trim(.TextMatrix(.Row, 10)), 2)    '염색패턴
        lstArray(9).ListIndex = FindItem(lstArray(9), Trim(.TextMatrix(.Row, 3)))        '작업자
        lstArray(10).ListIndex = FindItem(lstArray(10), Trim(.TextMatrix(.Row, 2)))      '작업조
        lblSchID = .TextMatrix(.Row, 15)
        lblSchSeq = .TextMatrix(.Row, 16)
        
        txtRoll = .TextMatrix(.Row, 8)
        txtQty = .TextMatrix(.Row, 9)
'        .TextMatrix ( .Row , 7)    'dtpStartDate(0)
        txtProcID.Text = .TextMatrix(.Row, 22)
        txtProcID.Tag = .TextMatrix(.Row, 18)

        dtpStartDate(0) = Left(.TextMatrix(.Row, 7), 10)
        dtpStartDate(1) = Right(Trim(.TextMatrix(.Row, 7)), 5)     '납기일자
        
        If Index = 2 Then
            chkEnd.Value = vbUnchecked
            dtpEndDate(0).Enabled = False
            dtpEndDate(0).Enabled = False
            dtpEndDate(0) = Now
            dtpEndDate(1) = Now
            lstArray(0).Enabled = False
            lstArray(6).Enabled = False
        Else
            chkEnd.Value = vbChecked
            If Trim(.TextMatrix(.Row, 21)) <> "" Then
                dtpEndDate(0) = Format(Left(.TextMatrix(.Row, 21), 8), "####-##-##")
                dtpEndDate(1) = Format(Right(.TextMatrix(.Row, 21), 4), "00:00")         '납기일자
            End If
             pnlEdit.Enabled = False
            lstArray(0).Enabled = True
            lstArray(6).Enabled = True
        End If
        
        '작업중인 카드
        Dim oRapid As PlusLib2.CRapid_NEW
        Dim rs As Recordset
        
        Set oRapid = New PlusLib2.CRapid_NEW
        oRapid.Connection = g_adoCon
        oRapid.UserName = g_sUserName
    
        txtResultDT = .TextMatrix(.Row, 17)
        txtRapidSeq = .TextMatrix(.Row, 20)
        txtProcID.Tag = .TextMatrix(.Row, 18)
        
        Set rs = oRapid.GetRapidCardList(Trim(.TextMatrix(.Row, 17)), Trim(.TextMatrix(.Row, 18)), Trim(.TextMatrix(.Row, 19)), Trim(.TextMatrix(.Row, 20)))
        Set oRapid = Nothing
        
        lstArray(6).Clear
        II = 0
        Do Until rs.EOF
            lstArray(6).AddItem rs!CardID & rs!SplitID
            lstArray(6).ItemData(II) = rs!WaitProcSeq
            II = II + 1
            rs.MoveNext
        Loop
        rs.Close
        Set oRapid = Nothing
        
    End With
    If Index = 2 Then
        pnlEdit.Enabled = True
    Else
        pnlEdit.Enabled = False
    End If
End Sub


Private Sub grdProcess_DblClick()
    Dim i%
    With grdNewPattern
        If .Row < .Rows - 1 Then
            If .TextMatrix(.Row + 1, 3) = "*" Then
                MsgBox "작업이 완료된 공정안에는 공정을 추가할 수 없습니다.", vbInformation + vbOKOnly
                Exit Sub
            End If
        End If
        
        .Redraw = flexRDNone
        
        .AddItem vbTab & grdProcess.TextMatrix(grdProcess.Row, 1) & vbTab & grdProcess.TextMatrix(grdProcess.Row, 2) & vbTab & "" & vbTab & 0, .Row + 1
        
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub grdTab_Click(PreviousTab As Integer)
    Call SetButtonStatus
    
    If grdTab.Tab = 0 Then
        Call InitGrid(0)
        Call FillGridData_NOW(0)
        cmdCard.Visible = False
    ElseIf grdTab.Tab = 1 Then
        Call InitGrid(1)
        Call FillGridData_NOW(1)
        cmdCard.Visible = False
    ElseIf grdTab.Tab = 2 Then
        Call InitINGGrid(2)
        Call FillGridData_ING(2)
        Call SetEditClear(0)
        cmdCard.Visible = False
    ElseIf grdTab.Tab = 3 Then
        Call InitGrid(3)
        cmdCard.Visible = True
        Call SetEditClear(0)
    End If
End Sub



Private Sub SetEditClear(ByVal Index As Integer)
    lstArray(6).Clear
    
    lstArray(0).ListIndex = 0
    lstArray(0).ListIndex = 0
    lstArray(7).ListIndex = 0
    
    lstArray(8).Selected(0) = True
    
    If lstArray(9).ListCount > 0 Then
        lstArray(9).Selected(0) = True
    End If
    
    If lstArray(10).ListCount > 0 Then
        lstArray(10).Selected(0) = True
    End If
    
    If lstArray(6).ListCount > 0 Then
        lstArray(6).Selected(0) = True
    End If
    
    lblSchID = ""
    lblSchSeq = ""
    txtResultDT.Text = ""
    txtRapidSeq.Text = ""
    txtCardID.Text = ""
    txtRoll.Text = ""
    txtQty.Text = ""
    dtpStartDate(0) = Now:    dtpStartDate(1) = Now
    dtpEndDate(0) = Now:      dtpEndDate(1) = Now
    dtpEndDate(0).Enabled = False
    dtpEndDate(1).Enabled = False
    txtRemarkResult.Text = ""
    chkEnd.Value = vbUnchecked
    If Index = 3 Then
        pnlEdit.Enabled = False
    Else
        pnlEdit.Enabled = True
    End If
    
    txtProcID.Text = ""
    txtProcID.Tag = ""
    
End Sub





Private Sub lstArray_Click(Index As Integer)
    Dim oRapid As PlusLib2.CRapid
    Dim oPerson  As PlusLib2.CPerson
    Dim rs As Recordset
    Dim iCount%
    
    
    Select Case Index
        Case 0
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
                If lstArray(1).ListCount > 0 Then
                    lstArray(1).ListIndex = 0
                End If
            End If
    End Select
End Sub

Private Sub lstArray_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim vCard As Variant
    Dim sCardID As String, sSplitID As String
    
    Select Case Index
        
        Case 6     '--- 작성된 카드 삭제시
            If KeyCode = vbKeyDelete Then
                
                With lstArray(6)
                    If .ListCount > 0 Then
                        vCard = Split(.Text, "-")
                        If UBound(vCard) = 1 Then
                            sCardID = vCard(0)
                            sSplitID = vCard(1)
                            
                        Else
                            sCardID = vCard(0)
                            sSplitID = ""
                        End If
                    End If
                End With
            End If
    End Select
End Sub

Private Sub lstArray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 8:
            
            lstArray(7).ListIndex = 0
            
        Case 7:
            If lstArray(7).Text = "염색" Then
                lstArray(8).ListIndex = 0
            Else
                lstArray(8).ListIndex = -1
            End If
    End Select

End Sub

Private Sub FillGridPattern()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetCardPattern(MakeCardID(grdList(3).TextMatrix(grdList(3).Rows - 1, 9), OM_REDUCE), grdList(3).TextMatrix(grdList(3).Rows - 1, 10))
    
    Set oCard = Nothing
    
    If rs.EOF Then
        grdCardPattern.Rows = grdCardPattern.FixedRow9s
        grdNewPattern.Rows = grdNewPattern.FixedRows
        Set rs = Nothing
        Exit Sub
    End If
    
    With grdCardPattern
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        For i = 0 To rs.RecordCount - 1
            .AddItem rs!PlanSeq & vbTab & rs!ProcessID & vbTab & rs!Process & vbTab & rs!CompleteClss & vbTab & _
                rs!NeedWidth & vbTab & rs!InstRemark & vbTab & rs!Remark
            
            If rs!CompleteClss = "*" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
            End If
            
            rs.MoveNext
        Next i
        
        .Redraw = flexRDDirect
    End With
    
    rs.MoveFirst
    With grdNewPattern
        .Redraw = flexRDNone
        .Rows = .FixedRows

        For i = 0 To rs.RecordCount - 1
            .AddItem rs!PlanSeq & vbTab & rs!ProcessID & vbTab & rs!Process & vbTab & rs!CompleteClss & vbTab & _
                rs!NeedWidth & vbTab & rs!InstRemark & vbTab & rs!Remark & vbTab & ""

            If rs!ReWorkClss = "*" Then
                .Cell(flexcpChecked, .Rows - 1, 7) = flexChecked
            End If

            If rs!CompleteClss = "*" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
            End If

            rs.MoveNext
        Next i

        For i = .Rows - 1 To .FixedRows Step -1
            If Trim(.TextMatrix(i, 3)) <> "" Then
                .Row = i
                .Select i, 1, i, .Cols - 1
                Exit For
            End If
        Next i
        .Redraw = flexRDDirect
    End With
    rs.Close
    Set rs = Nothing
        
    Exit Sub
    
ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmInstRapid_New.FillGridPattern", Err.Description)
End Sub

Private Sub FillGridProcess()
    Dim oCard As PlusLib2.CCard
    Dim rs As Recordset
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon

    Set rs = oCard.GetProcess()
    Set oCard = Nothing

    With grdProcess
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!ProcessID & vbTab & rs!Process
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, "frmInstRaid_New.FillGridProcess", Err.Description)
End Sub




Private Function UpdateCardPattern() As Boolean
    Dim tItem As PlusLib2.TCard
    Dim tItemSub() As PlusLib2.TCardPattern
    Dim oCard As PlusLib2.CCard
    Dim i%
    
    On Error GoTo ErrHandler
    
    UpdateCardPattern = False
    

    tItem.sCardID = MakeCardID(grdList(3).TextMatrix(grdList(3).Rows - 1, 9), OM_REDUCE)
    tItem.sSplitID = grdList(3).TextMatrix(grdList(3).Rows - 1, 10)
    tItem.sPersonID = g_sUserName
    tItem.sUseClss = "대기"
    
    With grdCardPattern
        For i = .FixedRows To .Rows - .FixedRows
            tItem.sPrePlanProc = tItem.sPrePlanProc & .TextMatrix(i, 2) & "→"
        Next i
        tItem.sPrePlanProc = Left(tItem.sPrePlanProc, Len(tItem.sPrePlanProc) - 1)
    End With
    
    With grdNewPattern
        For i = .FixedRows To .Rows - .FixedRows
            tItem.sPostPlanProc = tItem.sPostPlanProc & .TextMatrix(i, 2) & "→"
        Next i
        tItem.sPostPlanProc = Left(tItem.sPostPlanProc, Len(tItem.sPostPlanProc) - 1)
    End With
    
    With grdNewPattern
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 3) <> "*" Then
                tItem.sAfterProc = tItem.sAfterProc & .TextMatrix(i, 2) & "→"
            End If
        Next i
        tItem.sAfterProc = Left(tItem.sAfterProc, Len(tItem.sAfterProc) - 1)
    End With
    
    With grdNewPattern
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 3) <> "*" Then
                tItem.sWaitProcID = .TextMatrix(i, 1)
                tItem.nWaitProcSeq = .TextMatrix(i, 0)
                Exit For
            End If
        Next i
    End With
    
    With grdNewPattern
        ReDim tItemSub(.Rows - .FixedRows - 1)
        
        For i = .FixedRows To .Rows - 1
            tItemSub(i - 1).sCardID = tItem.sCardID
            tItemSub(i - 1).sSplitID = tItem.sSplitID
            tItemSub(i - 1).nPlanSeq = .TextMatrix(i, 0)
            tItemSub(i - 1).sProcessID = .TextMatrix(i, 1)
            tItemSub(i - 1).sCompleteClss = .TextMatrix(i, 3)
            tItemSub(i - 1).nNeedWidth = .TextMatrix(i, 4)
            tItemSub(i - 1).sInstRemark = .TextMatrix(i, 5)
            tItemSub(i - 1).sRemark = .TextMatrix(i, 6)
            tItemSub(i - 1).sReWorkClss = IIf(.Cell(flexcpChecked, i, 7) = flexChecked, "*", "")
        Next i
    End With
   
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    oCard.UserName = g_sUserName
    
    If oCard.UpdateCardPattern(tItem, tItemSub()) Then
        UpdateCardPattern = True
    End If
    Set oCard = Nothing
    Exit Function
ErrHandler:
    Set oCard = Nothing
    UpdateCardPattern = False
    Call ErrorBox(Err.Number, "frmInstRapid_New.UpdateCardPattern", Err.Description)
End Function

