VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCardDivide 
   ClientHeight    =   9255
   ClientLeft      =   2760
   ClientTop       =   1950
   ClientWidth     =   11850
   Icon            =   "frmCardDivide.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Crystal.CrystalReport cryReport 
      Left            =   330
      Top             =   8610
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlDivide 
      Height          =   4305
      Left            =   3630
      TabIndex        =   25
      Top             =   2820
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   7594
      _Version        =   196609
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboUseClss 
         Height          =   300
         Left            =   1410
         Style           =   2  '드롭다운 목록
         TabIndex        =   29
         Top             =   1530
         Width           =   1665
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '오른쪽 맞춤
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   35
         Text            =   "0"
         Top             =   810
         Width           =   1665
      End
      Begin VB.TextBox txtRoll 
         Alignment       =   1  '오른쪽 맞춤
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   33
         Text            =   "0"
         Top             =   450
         Width           =   1665
      End
      Begin VSFlex7LCtl.VSFlexGrid grdCard 
         Height          =   1455
         Left            =   60
         TabIndex        =   32
         Top             =   1950
         Width           =   3000
         _cx             =   5292
         _cy             =   2566
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
      Begin VB.TextBox txtDivide 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   1170
         Width           =   1665
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   690
         Left            =   1770
         TabIndex        =   31
         Top             =   3480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "취소"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   690
         Left            =   180
         TabIndex        =   30
         Top             =   3480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "저장"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   1170
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "분리할 카드 수"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   34
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "절수"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   36
         Top             =   810
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "수량"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   37
         Top             =   1530
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "카드 상태"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   30
         X2              =   4800
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   30
         X2              =   4800
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "공정카드분리"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   150
         TabIndex        =   26
         Top             =   135
         Width           =   1080
      End
      Begin VB.Shape Shape 
         BackColor       =   &H80000002&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   6  '내부 단색
         FillColor       =   &H00800000&
         Height          =   330
         Left            =   60
         Top             =   60
         Width           =   3045
      End
   End
   Begin Threed.SSCommand cmdDivide 
      Height          =   690
      Left            =   6600
      TabIndex        =   24
      Tag             =   "PERM_ADDNEW"
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "공정카드분리"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1614
      _Version        =   196609
      Begin VB.TextBox txtSearch 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   7830
         MaxLength       =   4
         TabIndex        =   42
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   6600
         MaxLength       =   8
         TabIndex        =   38
         Top             =   495
         Width           =   1185
      End
      Begin VB.ComboBox cboProcess 
         Height          =   300
         Left            =   8610
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   480
         Width           =   1320
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10980
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   4
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   6600
         TabIndex        =   3
         Top             =   75
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2820
         TabIndex        =   2
         Top             =   495
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   1
         Top             =   75
         Width           =   1905
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   795
         Left            =   60
         TabIndex        =   6
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
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   7
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
         TabIndex        =   9
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
            TabIndex        =   10
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   4785
         TabIndex        =   11
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
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   4770
         TabIndex        =   14
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
         TabIndex        =   15
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
            TabIndex        =   16
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   5220
         TabIndex        =   17
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
            TabIndex        =   18
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   8610
         TabIndex        =   39
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
            Caption         =   "대기공정"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   40
            Top             =   60
            Width           =   1185
         End
      End
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   450
      TabIndex        =   19
      Top             =   3480
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
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   120
         Width           =   270
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7485
      Left            =   0
      TabIndex        =   22
      Top             =   930
      Width           =   11835
      _cx             =   20876
      _cy             =   13203
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10110
      TabIndex        =   23
      Top             =   8490
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
      Left            =   8370
      TabIndex        =   41
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      발행(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmCardDivide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const REPORTFILE As String = "\Report\WorkCard.xls"
Private Const REPORTFILE1 As String = "\Report\TmpWorkCard.xls"

Private m_bLoading As Boolean

Private Sub cboUseClss_Click()
    With cboUseClss
        If cboUseClss = "보류" And cboUseClss.Tag = "대기" Then
            MsgBox "공정카드의 사용구분을 '보류'로 지정할 수 없습니다", vbInformation + vbOKOnly
            cboUseClss = cboUseClss.Tag
        End If
    End With
End Sub

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

Private Sub cmdCancel_Click()
    Call ModeChange(True)
    
    grdCard.Rows = 1
End Sub

Private Sub cmdDivide_Click()
Dim oRapid As Pluslib2.CRapid
Dim sRs As ADODB.Recordset

    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    If grdData.TextMatrix(grdData.Row, 12) = "작업" Then
        MsgBox "작업중인 카드는 분리작업을 할 수 없습니다.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    Set oRapid = New Pluslib2.CRapid
    oRapid.Connection = g_adoCon
    oRapid.UserName = g_sUserName
    
    Set sRs = oRapid.GetCheckDyeSch(MakeCardID(grdData.TextMatrix(grdData.Row, 6), OM_REDUCE), Trim(grdData.TextMatrix(grdData.Row, 7)))
    Set oRapid = Nothing
    
    If sRs.RecordCount > 0 Then
        If Trim(sRs!Complitclss) = "" Then
            Set sRs = Nothing
            MsgBox "염색작업지시가 내려진 카드는 카드분리를 할수 없습니다", vbInformation, "카드분리 불가"
            Exit Sub
        End If
    End If
    Set sRs = Nothing
    
    Call ModeChange(False)
    
    With grdData
        txtRoll = .TextMatrix(.Row, 9)
        txtQty = .TextMatrix(.Row, 10)
        cboUseClss = .TextMatrix(.Row, 13)
        cboUseClss.Tag = .TextMatrix(.Row, 13)
    End With
    
    grdCard.Rows = grdCard.FixedRows
    txtDivide.SetFocus
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

Private Sub cmdOK_Click()
    If grdCard.FixedRows = grdCard.Rows Then Exit Sub
    
    If Not CheckCardData Then Exit Sub
    
    If SaveData() Then
        Call ModeChange(True)
        Call FillGridData
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim sCardID$, sSplitID$, sPatternID$
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    With grdData
        sCardID = MakeCardID(.TextMatrix(.Row, 6), OM_REDUCE)
        sSplitID = .TextMatrix(.Row, 7)
        sPatternID = .TextMatrix(.Row, 16)
    End With
    
    Call PrintWorkCard(CryReport, sCardID, sSplitID, sPatternID, PlusMDI.PrintPreview)
End Sub

Public Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    Call InitGrid
    Call MakeProcessCombo
    Call ModeChange(True)
        
    With cboUseClss
        .AddItem "대기"
        .AddItem "보류"
        
        .ListIndex = -1
    End With
        
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    txtSearch(4).Enabled = False
    txtSearch(5).Enabled = False
    cboProcess.Enabled = False
    
    pnlProgress.Visible = False
End Sub

Private Sub grdCard_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdCard
        If Col = 0 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Or CheckNum(.TextMatrix(Row, Col)) < 1 Or CheckNum(.TextMatrix(Row, Col)) > 3 Then
                .TextMatrix(Row, Col) = "1"
            End If
            
            .Select Row, Col + 1
        ElseIf Col = 1 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Then
                .TextMatrix(Row, Col) = "0"
            End If
                                                            
            .Select Row, Col + 1
        ElseIf Col = 2 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Then
                .TextMatrix(Row, Col) = "0"
            End If
        
            If Row < .Rows - 1 Then
                .Select Row + 1, 0
                Call CalcLastCard
            End If
        End If
    End With
End Sub

Private Sub CalcLastCard()
    Dim i%
    
    With grdCard
        .TextMatrix(.Rows - 1, 1) = CheckNum(txtRoll)
        .TextMatrix(.Rows - 1, 2) = CheckNum(txtQty)
        For i = .FixedRows To .Rows - 1
            If i < .Rows - 1 Then
                .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1) - .TextMatrix(i, 1)
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) - .TextMatrix(i, 2)
            End If
        Next i
    End With
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

Private Sub txtDivide_KeyPress(KeyAscii As Integer)
    Dim i%
    
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(txtDivide) Then Exit Sub
        
        If txtDivide < 2 Then
            MsgBox "카드분리 수는 최소한 2개이상은 되어야 합니다.", vbInformation + vbOKOnly
            grdCard.Rows = grdCard.FixedRows
            txtDivide.SetFocus
            Exit Sub
        End If
        With grdCard
            .Rows = .FixedRows
            
            For i = 0 To txtDivide - 1
                .AddItem "1" & vbTab & 0 & vbTab & 0
            Next i
            .Row = .FixedRows
            .Select .Row, 0
            .SetFocus
        End With
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
    ElseIf KeyAscii = vbKeyReturn And Index >= 3 Then
        Call NextFocus
    End If
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Redraw = flexRDNone
        .Cols = 17
        
        Call SetVSFlexGrid(grdData)
        .Rows = 1
        .RowHeightMin = 390
        
        .TextArray(0) = " ":
        .TextArray(1) = " ":            .ColWidth(1) = 250:     .ColHidden(1) = True
        .TextArray(2) = "거래처":       .ColWidth(2) = 1000:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":         .ColWidth(3) = 1700:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":     .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNo":      .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "카드번호":     .ColWidth(6) = 1000:               .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "분할" & vbCrLf & "번호":     .ColWidth(7) = 500:            .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":         .ColWidth(8) = 1000:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "절수":         .ColWidth(9) = 500:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "수량":         .ColWidth(10) = 600:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "완료공정":    .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "대기공정":    .ColWidth(12) = 900:           .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "카드상태":    .ColWidth(13) = 900:           .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "계획공정":    .ColWidth(14) = 7000:           .ColAlignment(14) = flexAlignLeftCenter
        .TextArray(15) = "색상순번":    .ColWidth(15) = 0
        .TextArray(16) = "공정패턴":    .ColWidth(16) = 0
        
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
    
    With grdCard
        .Redraw = flexRDNone
        .Cols = 3
        
        Call SetVSFlexGrid(grdCard)
        .Rows = 1
        .FixedCols = 0
        
        .TextArray(0) = "튜브":     .ColWidth(0) = 500
        .TextArray(1) = "절수":     .ColWidth(1) = 1000:            .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "수량":     .ColWidth(2) = 1300:            .ColAlignment(2) = flexAlignRightCenter
        
        .ColFormat(1) = "#,##0"
        .ColFormat(2) = "#,##0"
        
        .Editable = flexEDKbdMouse
'        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub MakeProcessCombo()
    Dim oCard As Pluslib2.CCard
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bLoading = True
    
    Set oCard = New Pluslib2.CCard
    oCard.Connection = g_adoCon

    Set rs = oCard.GetProcess(1)
    Set oCard = Nothing

    With cboProcess
        .Clear

        Do Until rs.EOF
            .AddItem CStr(rs!Process)
            .ItemData(.NewIndex) = CLng(Left(rs!ProcessID, 2))
            
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing

        If .ListCount > 0 Then .ListIndex = 0
    End With

    m_bLoading = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    m_bLoading = False
    Call ErrorBox(Err.Number, "frmCardChange.MakeProcessCombo", Err.Description)
End Sub

Private Sub FillGridData()
    Dim oCard As Pluslib2.CCard
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    m_bLoading = True
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oCard = New Pluslib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetOrder(IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), txtSearch(4), txtSearch(5), _
                                 IIf(chkSearch(5) = vbChecked, 1, 0), Format(Left(cboProcess.ItemData(cboProcess.ListIndex), 2), "00"), 0)
    Set oCard = Nothing
        
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                    rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & rs!SplitID & vbTab & _
                    rs!Color & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                    rs!UseClss & vbTab & CheckNull(rs!AfterProc) & vbTab & rs!OrderSeq & vbTab & rs!PatternID
            
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
            
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bLoading = False
    Exit Sub

ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    pnlProgress.Visible = False
    m_bLoading = False
    Call ErrorBox(Err.Number, "frmCard.FillGridData", Err.Description)
End Sub

Private Sub ModeChange(bValue As Boolean)
    frmSearch.Enabled = bValue
    pnlDivide.Visible = Not bValue
    grdData.Enabled = bValue
    
    cmdDivide.Enabled = bValue
End Sub

Private Function SaveData() As Boolean
    Dim tItem() As Pluslib2.TCard
    Dim oCard As Pluslib2.CCard
    Dim i%
    
    On Error GoTo ErrHandler
    
    With grdCard
        ReDim tItem(.Rows - .FixedRows - 1)
        
        For i = 1 To .Rows - 1
            tItem(i - 1).sCardID = MakeCardID(grdData.TextMatrix(grdData.Row, 6), OM_REDUCE)
            tItem(i - 1).sSplitID = grdData.TextMatrix(grdData.Row, 7)
            tItem(i - 1).nRoll = .TextMatrix(i, 1)
            tItem(i - 1).nQty = .TextMatrix(i, 2)
            tItem(i - 1).sNewSplitID = RTrim(grdData.TextMatrix(grdData.Row, 7)) & i
            tItem(i - 1).sPersonID = g_sUserName
            tItem(i - 1).sModiClss = "카드분리"
            tItem(i - 1).sPatternID = grdData.TextMatrix(grdData.Row, 16)
            tItem(i - 1).sUseClss = cboUseClss
            tItem(i - 1).nChkUseClss = 0
            If cboUseClss <> cboUseClss.Tag And cboUseClss.Tag = "보류" Then
                tItem(i - 1).nChkUseClss = 1 '보류에서 대기로 변경될때 Hold Table 보류 취소 업데이트
            End If
            tItem(i - 1).nTubeNo = .TextMatrix(i, 0)
        Next i
    End With
    Set oCard = New Pluslib2.CCard
    oCard.Connection = g_adoCon
    oCard.UserName = g_sUserName
    
    If oCard.UpdateCardDivide(tItem) Then
        SaveData = True
    Else
        SaveData = False
    End If
    Set oCard = Nothing
    
    Exit Function
ErrHandler:
    Set oCard = Nothing
    SaveData = False
    Call ErrorBox(Err.Number, "frmCardChange.SaveData", Err.Description)
End Function

Private Function CheckCardData() As Boolean
    Dim i%, nRoll%, nQty%
    
    With grdCard
        For i = .FixedRows To .Rows - 1
            nRoll = nRoll + Abs(.TextMatrix(i, 1))
            nQty = nQty + Abs(.TextMatrix(i, 2))
        Next i
    End With
    
    If nRoll > CInt(txtRoll) Or nQty > CInt(txtQty) Then
        MsgBox "원래 카드의 수량보다 크게 분리 할수는 없습니다" & vbCrLf & "카드수량을 정정하여 주십시오", vbInformation
        CheckCardData = False
    Else
        CheckCardData = True
    End If
End Function

