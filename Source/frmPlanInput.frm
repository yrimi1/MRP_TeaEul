VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanInput 
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   1710
   ClientWidth     =   15240
   Icon            =   "frmPlanInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15240
   Begin Threed.SSFrame frmToDay 
      Height          =   765
      Left            =   30
      TabIndex        =   49
      Top             =   8610
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1349
      _Version        =   196609
      Caption         =   "금일계획수량"
      Begin Threed.SSPanel pnlToday 
         Height          =   345
         Left            =   60
         TabIndex        =   50
         Top             =   300
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         _Version        =   196609
         Caption         =   "0 YDS"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.CheckBox chkExpand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "공정확장"
      Height          =   500
      Left            =   30
      Style           =   1  '그래픽
      TabIndex        =   48
      Top             =   960
      Width           =   495
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   2340
      TabIndex        =   44
      Top             =   3270
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
         TabIndex        =   45
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
         TabIndex        =   46
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSFrame frmPlanInput 
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   5070
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   1508
      _Version        =   196609
      Begin VB.CheckBox chkStuffClose 
         Caption         =   "투입완료"
         Height          =   315
         Left            =   12060
         TabIndex        =   51
         Top             =   90
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpExpectDate 
         Height          =   300
         Left            =   10170
         TabIndex        =   36
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Format          =   73334784
         CurrentDate     =   37950
      End
      Begin VB.ComboBox cboPattern 
         Height          =   300
         Left            =   1140
         Style           =   2  '드롭다운 목록
         TabIndex        =   38
         Top             =   450
         Width           =   12375
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   7530
         TabIndex        =   35
         Top             =   90
         Width           =   1485
      End
      Begin VB.ComboBox cboCmdColor 
         Height          =   300
         Left            =   4035
         Style           =   2  '드롭다운 목록
         TabIndex        =   30
         Top             =   90
         Width           =   2370
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   90
         TabIndex        =   31
         Top             =   105
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "지시 일자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpCmdDate 
         Height          =   300
         Left            =   1155
         TabIndex        =   32
         Top             =   105
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   393216
         Format          =   73334784
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   8
         Left            =   3000
         TabIndex        =   33
         Top             =   90
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "지시 색상"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   10
         Left            =   6480
         TabIndex        =   34
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "지시 수량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   7
         Left            =   90
         TabIndex        =   37
         Top             =   450
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "공정 패턴"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdPlus 
         Height          =   720
         Left            =   13650
         TabIndex        =   39
         Top             =   60
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         _Version        =   196609
         Caption         =   "지시"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   720
         Left            =   14430
         TabIndex        =   43
         Top             =   60
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         _Version        =   196609
         Caption         =   "취소"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   9
         Left            =   9090
         TabIndex        =   47
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "작업완료일"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   4125
      Left            =   0
      TabIndex        =   26
      Top             =   930
      Width           =   15225
      _cx             =   26855
      _cy             =   7276
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
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   1614
      _Version        =   196609
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   14370
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   23
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   10350
         TabIndex        =   20
         Top             =   90
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   6630
         TabIndex        =   16
         Top             =   465
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   6630
         TabIndex        =   12
         Top             =   75
         Width           =   1905
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   465
         Width           =   600
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   795
         Left            =   60
         TabIndex        =   1
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
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   480
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3450
         TabIndex        =   6
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   73334785
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3435
         TabIndex        =   7
         Top             =   465
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   73334785
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   2130
         TabIndex        =   8
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
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
            Caption         =   "수주 일자"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   9
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   5250
         TabIndex        =   13
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
            TabIndex        =   14
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   8595
         TabIndex        =   15
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
         Left            =   5250
         TabIndex        =   17
         Top             =   465
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
            TabIndex        =   18
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   8610
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   465
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
         Left            =   9030
         TabIndex        =   21
         Top             =   90
         Width           =   1260
         _ExtentX        =   2223
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
            TabIndex        =   22
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   12360
         TabIndex        =   24
         Top             =   90
         Width           =   1710
         _ExtentX        =   3016
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
            Caption         =   "마감분 포함"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   1275
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   12360
         TabIndex        =   27
         Top             =   465
         Width           =   1710
         _ExtentX        =   3016
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
            Caption         =   "투입완료분 포함"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   1635
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4755
         TabIndex        =   11
         Top             =   135
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   4755
         TabIndex        =   10
         Top             =   525
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdCommand 
      Height          =   690
      Left            =   11790
      TabIndex        =   41
      Tag             =   "PERM_ADDNEW"
      Top             =   8610
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      작업지시(&C)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13515
      TabIndex        =   42
      Top             =   8610
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdPlanData 
      Height          =   2565
      Left            =   0
      TabIndex        =   40
      Top             =   5940
      Width           =   15225
      _cx             =   26855
      _cy             =   4524
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
Attribute VB_Name = "frmPlanInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bLoading As Boolean
Private m_sOrderID As String

Private Sub chkExpand_Click()
    Dim i%
    With grdOrder
        For i = 13 To 23
            .ColHidden(i) = IIf(chkExpand.Value = vbChecked, False, True)
        Next i
        
        .ColHidden(14) = True
        .ColHidden(15) = True
        .ColHidden(17) = True
        .ColHidden(20) = True
        
        If chkExpand.Value = vbChecked Then
             .ScrollBars = flexScrollBarBoth
        Else
            .ScrollBars = flexScrollBarVertical
        End If
    End With
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else
        If chkSearch(Index).Value = vbChecked Then
            If Index = 1 Or Index = 2 Or Index = 3 Then
                txtSearch(Index).Enabled = True
                txtSearch(Index).SetFocus
            End If
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            If Index = 1 Or Index = 2 Or Index = 3 Then
                txtSearch(Index).Enabled = False
                cmdSearch.SetFocus
            End If

            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    End If

End Sub

Private Sub cmdCommand_Click()
    If grdPlanData.Rows = grdPlanData.FixedRows Then Exit Sub

'    If Not CheckData() Then Exit Sub

    If SaveData() Then
        grdPlanData.Rows = grdPlanData.FixedRows
        cmdPlus.Enabled = True
        cmdDel.Enabled = False
        cmdCommand.Enabled = False
        frmSearch.Enabled = True
        grdOrder.Enabled = True
        
        Call GetInstQtyByDate
        
        txtQty = ""
        dtpExpectDate = Now
        cboPattern.ListIndex = 0
        
        Call FillGridOrder
    End If
End Sub

Private Sub cmdDel_Click()
    With grdPlanData
        If .Rows = .FixedRows Or .Row < .FixedRows Then Exit Sub

        If MsgBox(LoadResString(207), vbQuestion + vbYesNo, "취소확인") = vbYes Then
            .Rows = .FixedRows

            cmdPlus.Enabled = True
            cmdDel.Enabled = False
            cmdCommand.Enabled = False
            frmSearch.Enabled = True
            grdOrder.Enabled = True
            
            txtQty = ""
            
            .HighLight = flexHighlightNever
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If
End Sub

Private Sub cmdPlus_Click()
    Dim i%, iRow&
    
    If grdOrder.IsSubtotal(grdOrder.Row) = False Or grdOrder.Rows = grdOrder.FixedRows Then Exit Sub
    
    If Not IsNumeric(txtQty.Text) Then
        MsgBox "지시량이 잘못 입력 되었습니다", vbInformation, "지시수량"
        Exit Sub
    ElseIf CDbl(txtQty.Text) <= 0 Then
        MsgBox "지시량을 정확히 입력하십시요", vbInformation, "지시수량"
        Exit Sub
    End If
    
    If grdOrder.TextMatrix(grdOrder.Row, 7) < grdOrder.TextMatrix(grdOrder.Row, 9) + txtQty Then
        If MsgBox("계획량이 입고량보다 많습니다." & vbCrLf & "그래도 지시하시겠습니까?", vbInformation + vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    
    
    If cboPattern.ListIndex < 0 Then
        MsgBox "공정 패턴을 선택하십시오.", vbInformation
        cboPattern.SetFocus
        Exit Sub
    End If

    cmdPlus.Enabled = False
    cmdDel.Enabled = True
    cmdCommand.Enabled = True
    
    frmSearch.Enabled = False
    grdOrder.Enabled = False
    
    Call FillGridPlanData
End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[1] 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = Date
    End If
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
    cmdDel.Enabled = False
    cmdCommand.Enabled = False
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15360, 9840
    
    Call SetOperate(Me)
    Call InitGrid
    Call MakePatternCombo
    
    For i = 0 To 1
        dtpDate(i) = Now
    Next i
    dtpCmdDate = Now
    dtpExpectDate = Now
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    
    cmdPlus.Picture = LoadResPicture("ADDNEW", vbResIcon)
    cmdDel.Picture = LoadResPicture("DELETE", vbResIcon)
    cmdCommand.Picture = LoadResPicture("COMMAND", vbResIcon)
    
    pnlProgress.Visible = False
    m_bLoading = False
    
    Call GetInstQtyByDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdOrder_DblClick()
    With grdOrder
        If .Row < .FixedRows Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub

Private Sub grdOrder_RowColChange()
    If m_bLoading Then Exit Sub

    If grdOrder.IsSubtotal(grdOrder.Row) = False Or grdOrder.Rows = grdOrder.FixedRows Then Exit Sub

    With grdOrder
        If .Rows = .FixedRows Or .Row < .FixedRows Then Exit Sub

        If .TextMatrix(.Row, 29) = "*" Then
            chkStuffClose.Value = vbChecked
        Else
            chkStuffClose.Value = vbUnchecked
        End If
        
        If Not IsNumeric(.TextMatrix(.Row, 27)) Then
            cboPattern.ListIndex = -1
        Else
            cboPattern.ListIndex = FindComboBox(cboPattern, CLng(.TextMatrix(.Row, 27)))
        End If
        
        Call MakeColorCombo
    End With
End Sub

Private Sub grdPlanData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdPlanData
        If Col = 4 Then
            If IsNumeric(.TextMatrix(Row, Col)) Then
                .Select Row, Col + 1
            Else
                .TextMatrix(Row, Col) = "0"
            End If
        ElseIf Col = 5 Then
            .Select Row, Col + 1
        ElseIf Col = 6 Then
            If Row < .Rows - 1 Then
            .Select Row + 1, 4
            End If
        End If
    End With
End Sub

Private Sub grdPlanData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 4 Then Cancel = True
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdOrder
        If optOrder(0).Value Then
            .ColWidth(4) = 1550
            .ColWidth(3) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(4) = 0
            .ColWidth(3) = 1550
            chkSearch(3).Caption = "관리번호"
        End If
    End With
End Sub

Private Sub txtQty_GotFocus()
    Call GotFocusText(txtQty)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQty = SetCurrency(txtQty)
    End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    ElseIf KeyAscii = vbKeyReturn And Index = 3 Then
        Call NextFocus
    End If
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdOrder
        .Cols = 30

        .Redraw = flexRDNone

        .Rows = 2
        .FixedRows = 2
        .FixedCols = 0
        .FrozenCols = 5
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        .TextArray(0) = " ":            .ColWidth(0) = 500
        .TextArray(1) = "거래처":       .ColWidth(1) = 1300:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "품명":         .ColWidth(2) = 2000:             .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호" & vbCrLf & "색  상  명":     .ColWidth(3) = 1550:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "Order No." & vbCrLf & "색  상  명":   .ColWidth(4) = 0:               .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "축율":         .ColWidth(5) = 700:            .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "수주량":       .ColWidth(6) = 900:            .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "입고량":       .ColWidth(7) = 900:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "미계획량":     .ColWidth(8) = 900:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "계획량":       .ColWidth(9) = 900:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "배색":        .ColWidth(10) = 900:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "배색":        .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "공정량":      .ColWidth(12) = 900:            .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "정련":        .ColWidth(13) = 900:            .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "수세":        .ColWidth(14) = 900:            .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "S/K":         .ColWidth(15) = 900:            .ColAlignment(15) = flexAlignRightCenter
        .TextArray(16) = "SETT":        .ColWidth(16) = 900:            .ColAlignment(16) = flexAlignRightCenter
        .TextArray(17) = "PEACH":       .ColWidth(17) = 900:            .ColAlignment(17) = flexAlignRightCenter
        .TextArray(18) = "CPB":         .ColWidth(18) = 900:            .ColAlignment(18) = flexAlignRightCenter
        .TextArray(19) = "염색":        .ColWidth(19) = 900:            .ColAlignment(19) = flexAlignRightCenter
        .TextArray(20) = "DRY":         .ColWidth(20) = 900:            .ColAlignment(20) = flexAlignRightCenter
        .TextArray(21) = "가공":        .ColWidth(21) = 900:            .ColAlignment(21) = flexAlignRightCenter
        .TextArray(22) = "검사":        .ColWidth(22) = 900:            .ColAlignment(22) = flexAlignRightCenter
        .TextArray(23) = "보류":        .ColWidth(23) = 900:            .ColAlignment(23) = flexAlignRightCenter
        .TextArray(24) = "검사":        .ColWidth(24) = 900:            .ColAlignment(24) = flexAlignRightCenter
        .TextArray(25) = "검사":        .ColWidth(25) = 900:            .ColAlignment(25) = flexAlignRightCenter
        .TextArray(26) = "출고량":      .ColWidth(26) = 1000:           .ColAlignment(26) = flexAlignRightCenter
        .TextArray(27) = "공정패턴코드":        .ColWidth(27) = 0
        .TextArray(28) = "가공폭":      .ColWidth(28) = 0
        .TextArray(29) = "투입구분":    .ColWidth(29) = 0
        
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
        .TextArray(.Cols + 14) = "수세"
        .TextArray(.Cols + 15) = "S/K"
        .TextArray(.Cols + 16) = "SETT"
        .TextArray(.Cols + 17) = "PEACH"
        .TextArray(.Cols + 18) = "CPB"
        .TextArray(.Cols + 19) = "염색"
        .TextArray(.Cols + 20) = "DRY"
        .TextArray(.Cols + 21) = "가공"
        .TextArray(.Cols + 22) = "검사"
        .TextArray(.Cols + 23) = "보류"
        .TextArray(.Cols + 24) = "합격"
        .TextArray(.Cols + 25) = "불합격"
        .TextArray(.Cols + 26) = "출고량"
        .TextArray(.Cols + 27) = "공정패턴코드"
        .TextArray(.Cols + 28) = "가공폭"
        .TextArray(.Cols + 29) = "투입구분"

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
        .ColFormat(25) = "#,##0"
        .ColFormat(26) = "#,##0"
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
        For i = 0 To 9
            .MergeCol(i) = True
        Next i
        
        For i = 12 To 23
            .MergeCol(i) = True
        Next i
        .MergeCol(26) = True
        .MergeCol(27) = True
        .MergeCol(28) = True
        .MergeCol(29) = True
       
        For i = 1 To .Cols - 1
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
        Next i
        
        For i = 13 To 23
            .ColHidden(i) = True
        Next i
    
        .ColHidden(14) = True
        .ColHidden(15) = True
        .ColHidden(17) = True
        .ColHidden(20) = True
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 0
        .Redraw = flexRDDirect
    End With
    
    With grdPlanData
        .Cols = 7
        Call SetVSFlexGrid(grdPlanData)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = " "
        .TextArray(1) = "공정코드":       .ColWidth(1) = 0:               .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "공정명":         .ColWidth(2) = 1500:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "가공폭":         .ColWidth(3) = 800:             .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "요구폭":         .ColWidth(4) = 800:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "지시사항":       .ColWidth(5) = 5000:            .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "비고사항":       .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignLeftCenter
        
        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub MakePatternCombo()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim rs As ADODB.Recordset
    Dim sPattern$, sPatternID$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    Set rs = oPlanInput.GetPattern()
    Set oPlanInput = Nothing
    
    With cboPattern
        .Clear

        sPatternID = rs!PatternID
        sPattern = rs!PatternID & ". " & rs!Pattern & " : "

        Do Until rs.EOF
            If sPatternID = rs!PatternID Then
                sPattern = sPattern & "[" & rs!Process & "]→ "
            Else
                .AddItem Left$(sPattern, Len(sPattern) - 2)
                .ItemData(.NewIndex) = CLng(sPatternID)

                sPatternID = rs!PatternID
                sPattern = rs!PatternID & ". " & rs!Pattern & " : "
            End If

            rs.MoveNext
        Loop
        If rs.RecordCount > 0 Then
            .AddItem Left$(sPattern, Len(sPattern) - 2)
            .ItemData(.NewIndex) = CLng(sPatternID)
        End If

        If .ListCount > 0 Then .ListIndex = 0
    End With
    rs.Close

    Set rs = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oPlanInput = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmPlanInput.MakePatternCombo", Err.Description)
End Sub

Private Sub MakeColorCombo()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim rs As ADODB.Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    Set rs = oPlanInput.GetOrderSub(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 3), OM_REDUCE))
    Set oPlanInput = Nothing
    
    With cboCmdColor
        .Clear

        Do Until rs.EOF
            .AddItem rs!Color
            .ItemData(.NewIndex) = CLng(rs!OrderSeq)
            
            rs.MoveNext
        Loop

        If .ListCount > 0 Then .ListIndex = 0
    End With
    rs.Close

    Set rs = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oPlanInput = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmPlanInput.MakeColorCombo", Err.Description)
End Sub

Private Sub FillGridOrder()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim rs As ADODB.Recordset
    Dim i%, nTop%, nCurRow%, sStuffClose$
    Dim nNoPlanQty#, nProceTotalQty#
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
    
    m_bLoading = True
    
    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    Set rs = oPlanInput.GetOrder(IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                 IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), IIf(chkSearch(5) = vbChecked, 1, 0))
    Set oPlanInput = Nothing
        
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            nNoPlanQty = IIf(rs!UnitClss = "0", rs!ColorQty, CLng(rs!ColorQty / 0.9144)) * (1 + rs!ChunkRate / 100) - rs!InstQty  '미계획량
            nProceTotalQty = rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty + rs!수세Qty + _
                            rs!SKQty + rs!셋팅Qty + rs!PeachQty + rs!C염색Qty + rs!염색Qty + rs!P염색Qty + rs!R수세Qty + _
                            rs!건조Qty + rs!가공Qty + rs!검사Qty + rs!PauseQty
                            
            If rs!StuffCloseClss = "*" Then
                sStuffClose = "■"
            Else
                sStuffClose = ""
            End If
                            
            If rs!OrderID <> MakeOrderID(.TextMatrix(nTop, 3), OM_REDUCE) Then
                .AddItem sStuffClose & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    rs!ChunkRate & vbTab & Format(rs!OrderQty, "#,###") & IIf(rs!UnitClss = "0", "", "M") & vbTab & rs!StuffInQty & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    CheckNull(rs!PatternID) & vbTab & rs!WorkWidth & vbTab & rs!StuffCloseClss
                
                
                Call DoFlexGridGroup(grdOrder, .Rows - 1, 1)
                Call GridCollapse(grdOrder, nTop)
                
                nTop = .Rows - 1
            End If
        
            .AddItem "" & vbTab & "" & vbTab & "" & vbTab & rs!Color & vbTab & rs!Color & vbTab & _
                "" & vbTab & rs!ColorQty & vbTab & "0" & vbTab & nNoPlanQty & vbTab & _
                rs!InstQty & vbTab & rs!InstQty - rs!배색TQty & vbTab & rs!배색TQty & vbTab & nProceTotalQty & vbTab & _
                rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty & vbTab & _
                rs!수세Qty & vbTab & rs!SKQty & vbTab & _
                rs!셋팅Qty & vbTab & rs!PeachQty & vbTab & rs!C염색Qty & vbTab & _
                rs!염색Qty + rs!P염색Qty + rs!R수세Qty & vbTab & rs!건조Qty & vbTab & rs!가공Qty & vbTab & _
                rs!검사Qty & vbTab & rs!PauseQty & vbTab & rs!PassQty & vbTab & rs!DefectQty & vbTab & rs!OutQty & vbTab & _
                CheckNull(rs!PatternID) & vbTab & rs!WorkWidth & vbTab & rs!StuffCloseClss
        
'            .TextMatrix(nTop, 7) = CLng(.TextMatrix(nTop, 7)) + rs!StuffInQty
            .TextMatrix(nTop, 8) = CLng(.TextMatrix(nTop, 8)) + nNoPlanQty
            .TextMatrix(nTop, 9) = CLng(.TextMatrix(nTop, 9)) + rs!InstQty
            .TextMatrix(nTop, 10) = CLng(.TextMatrix(nTop, 10)) + rs!InstQty - rs!배색TQty
            .TextMatrix(nTop, 11) = CLng(.TextMatrix(nTop, 11)) + rs!배색TQty
            .TextMatrix(nTop, 12) = CLng(.TextMatrix(nTop, 12)) + nProceTotalQty
            .TextMatrix(nTop, 13) = CLng(.TextMatrix(nTop, 13)) + rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty
            .TextMatrix(nTop, 14) = CLng(.TextMatrix(nTop, 14)) + rs!수세Qty
            .TextMatrix(nTop, 15) = CLng(.TextMatrix(nTop, 15)) + rs!SKQty
            .TextMatrix(nTop, 16) = CLng(.TextMatrix(nTop, 16)) + rs!셋팅Qty
            .TextMatrix(nTop, 17) = CLng(.TextMatrix(nTop, 17)) + rs!PeachQty
            .TextMatrix(nTop, 18) = CLng(.TextMatrix(nTop, 18)) + rs!C염색Qty
            .TextMatrix(nTop, 19) = CLng(.TextMatrix(nTop, 19)) + rs!염색Qty + rs!P염색Qty + rs!R수세Qty
            .TextMatrix(nTop, 20) = CLng(.TextMatrix(nTop, 20)) + rs!건조Qty
            .TextMatrix(nTop, 21) = CLng(.TextMatrix(nTop, 21)) + rs!가공Qty
            .TextMatrix(nTop, 22) = CLng(.TextMatrix(nTop, 22)) + rs!검사Qty
            .TextMatrix(nTop, 23) = CLng(.TextMatrix(nTop, 23)) + rs!PauseQty
            .TextMatrix(nTop, 24) = CLng(.TextMatrix(nTop, 24)) + rs!PassQty
            .TextMatrix(nTop, 25) = CLng(.TextMatrix(nTop, 25)) + rs!DefectQty
            .TextMatrix(nTop, 26) = CLng(.TextMatrix(nTop, 26)) + rs!OutQty
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            nCurRow = FindGridOrderID
            If nCurRow >= .Rows Then
                .Row = .FixedRows
            Else
                .Row = nCurRow
                .IsCollapsed(.Row) = flexOutlineExpanded
            End If
        Else
            .HighLight = flexHighlightNever
            grdPlanData.Rows = grdPlanData.FixedRows
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bLoading = False
    
    If grdOrder.Rows > grdOrder.FixedRows Then
        Call GridCollapse(grdOrder, nTop)
        Call MakeColorCombo
    End If
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    m_bLoading = False
    Set oPlanInput = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmPlanInput.FillGridOrder", Err.Description)
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
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = COLOR_GRIDROW
'            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HFFFFC0    '&HE0E0E0
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

Private Function SaveData() As Boolean
    Dim tPlan As PlusLib2.TPlanInput
    Dim tPlanSub() As PlusLib2.TPlanInputSub
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim i%

    Screen.MousePointer = vbHourglass
    SaveData = False

    On Error GoTo ErrHandler

    With grdOrder
        tPlan.sInstDate = MakeDate(DF_SHORT, dtpCmdDate)
        tPlan.sOrderID = MakeOrderID(.TextMatrix(.Row, 3), OM_REDUCE)
        tPlan.nOrderSeq = cboCmdColor.ItemData(cboCmdColor.ListIndex)
        tPlan.nInstQty = CheckNum(txtQty)
        tPlan.sExpectDate = MakeDate(DF_SHORT, dtpExpectDate)
        tPlan.sPersonID = g_sUserName
        tPlan.sPatternID = Format(cboPattern.ItemData(cboPattern.ListIndex), "00")
        tPlan.sStuffCloseClss = IIf(chkStuffClose.Value, "*", "")
    End With

    With grdPlanData
        ReDim tPlanSub(.Rows - 2)
    
        For i = 1 To .Rows - 1
            tPlanSub(i - 1).sInstDate = MakeDate(DF_SHORT, dtpCmdDate)
            tPlanSub(i - 1).nProcSeq = i
            tPlanSub(i - 1).sProcessID = .TextMatrix(i, 1)
            tPlanSub(i - 1).nNeedWidth = CheckNum(.TextMatrix(i, 4))
            tPlanSub(i - 1).sInstRemark = .TextMatrix(i, 5)
            tPlanSub(i - 1).sRemark = .TextMatrix(i, 6)
        Next i
    End With

    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    oPlanInput.UserName = g_sUserName
    
    If oPlanInput.AddNewPlanInPut(tPlan, tPlanSub()) Then
        SaveData = True
    Else
        SaveData = False
    End If
    Set oPlanInput = Nothing

    m_sOrderID = tPlan.sOrderID
    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    SaveData = False

    Set oPlanInput = Nothing
    Call ErrorBox(Err.Number, "frmPlanInput.SaveData", Err.Description)
End Function

Private Sub FillGridPlanData()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim rs As ADODB.Recordset
    Dim i%

    On Error GoTo ErrHandler

    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    Set rs = oPlanInput.GetPatternOne(Format(cboPattern.ItemData(cboPattern.ListIndex), "00"))
    Set oPlanInput = Nothing

    With grdPlanData
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!processid & vbTab & rs!Process & vbTab & grdOrder.TextMatrix(grdOrder.Row, 28)
            
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        .Cell(flexcpBackColor, 1, 2, .Rows - 1, 3) = COLOR_GRIDROW
        .Select 1, 3
        .SetFocus
    End With

    Exit Sub

ErrHandler:
    Set oPlanInput = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmPlanInput.FillGridPlanData", Err.Description)
End Sub

Private Function CheckData() As Boolean
    Dim i%
    
    CheckData = False
    With grdPlanData
        For i = .FixedRows To .Rows - 1
            If Len(.TextMatrix(i, 4)) = 0 Then
                MsgBox "요구폭을 입력하지 않으셨습니다." & vbCrLf & "요구폭을 입력해 주십시오", vbInformation + vbOKOnly
                Exit Function
            End If
        Next i
    End With
    CheckData = True
End Function

Private Sub GetInstQtyByDate()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim nInQty&
    
    On Error GoTo ErrHandler
    
    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    nInQty = oPlanInput.GetInstQtyByDate(MakeDate(DF_SHORT, dtpCmdDate))
    
    Set oPlanInput = Nothing
    
    pnlToday.Caption = Format(nInQty, "#,##0") & " YDS"
    Exit Sub
    
ErrHandler:
    Set oPlanInput = Nothing
    
    Call ErrorBox(Err.Number, "frmPlanInput.GetInstQtyByDate", Err.Description)
End Sub

Private Function FindGridOrderID() As Long
    Dim i%
    
    With grdOrder
        FindGridOrderID = .Rows
        For i = .FixedRows To .Rows - 1
            If m_sOrderID = MakeOrderID(.TextMatrix(i, 3), OM_REDUCE) And Len(m_sOrderID) > 0 Then
                FindGridOrderID = i
                Exit Function
            End If
        Next i
    End With
End Function

