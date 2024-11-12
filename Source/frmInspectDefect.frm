VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectDefect 
   Caption         =   "검사 현황"
   ClientHeight    =   9255
   ClientLeft      =   1725
   ClientTop       =   1650
   ClientWidth     =   11865
   Icon            =   "frmInspectDefect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.Frame fraPrint 
      Height          =   705
      Left            =   7485
      TabIndex        =   34
      Top             =   8475
      Width           =   930
      Begin VB.OptionButton optLang 
         Caption         =   "한 글"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   180
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.OptionButton optLang 
         Caption         =   "영 문"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   435
         Width           =   720
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdDefect 
      Height          =   1875
      Left            =   7620
      TabIndex        =   33
      Top             =   4500
      Visible         =   0   'False
      Width           =   4095
      _cx             =   7223
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
   Begin Threed.SSPanel pnlMsg 
      Height          =   675
      Left            =   5430
      TabIndex        =   32
      Top             =   60
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   1191
      _Version        =   196609
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "불량 위치 현황 리포트를 출력하실려면 색상명을 먼저 선택하셔야 합니다."
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   6585
      Left            =   0
      TabIndex        =   31
      Top             =   1890
      Width           =   3600
      _cx             =   6350
      _cy             =   11615
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
   Begin VB.Frame fraOrder 
      Height          =   795
      Left            =   30
      TabIndex        =   28
      Top             =   8430
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   210
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   495
         Width           =   1200
      End
   End
   Begin VB.Frame fraRange 
      Height          =   780
      Left            =   3630
      TabIndex        =   21
      Top             =   -30
      Width           =   1710
      Begin VB.OptionButton optRange 
         Caption         =   "100"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.OptionButton optRange 
         Caption         =   "200"
         Height          =   210
         Index           =   1
         Left            =   975
         TabIndex        =   22
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblName 
         Caption         =   "■ 범위 설정"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   180
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   1695
      Top             =   8820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraSearc 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   -60
      Width           =   3600
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   300
         Index           =   1
         Left            =   75
         MousePointer    =   99  '사용자 정의
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   495
         Width           =   510
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   300
         Index           =   2
         Left            =   75
         MousePointer    =   99  '사용자 정의
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   825
         Width           =   510
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   1185
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   765
         Left            =   2760
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   165
         Width           =   765
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   2
         Left            =   1380
         TabIndex        =   5
         Top             =   1545
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   630
         TabIndex        =   10
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   36271
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   630
         TabIndex        =   11
         Top             =   825
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   36271
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   165
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "검사일자"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   13
            Top             =   60
            Width           =   1050
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거래처"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   15
            Top             =   60
            Width           =   1050
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   2
         Left            =   75
         TabIndex        =   16
         Top             =   1545
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "오더번호"
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   17
            Top             =   60
            Width           =   1080
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Left            =   2925
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1185
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   1
         Left            =   1950
         TabIndex        =   20
         Top             =   555
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Left            =   1950
         TabIndex        =   19
         Top             =   885
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8445
      TabIndex        =   1
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   2
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdResult 
      Height          =   7710
      Left            =   3630
      TabIndex        =   3
      Top             =   780
      Width           =   8210
      _cx             =   14482
      _cy             =   13600
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
   Begin VB.Frame fraReport 
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   780
      Left            =   5625
      TabIndex        =   25
      Top             =   -30
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton optMain 
         Caption         =   "(대표) 불량현황"
         Height          =   315
         Index           =   1
         Left            =   225
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   450
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optMain 
         Caption         =   "불량위치 현황"
         Height          =   315
         Index           =   0
         Left            =   225
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label lblCount 
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
      Height          =   225
      Left            =   3435
      TabIndex        =   0
      Top             =   8820
      Width           =   3120
   End
End
Attribute VB_Name = "frmInspectDefect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH1 = 1300   '1470
Private Const LIMIT_WIDTH2 = 1535
Private Const LIMIT_WIDTH3 = 1250

Private Const LIMIT_ROW1 = 23
Private Const LIMIT_ROW2 = 26
Private Const LIMIT_ROW3 = 11

Private Const REPORTFILE1 = "\Report\InspectDefect.rpt"
Private Const REPORTFILE2 = "\Report\InspectDefect_e.rpt"

Private Const BASE_X       As Integer = 150
Private Const BASE_Y       As Integer = 1300
Private Const DEFECT_COUNT As Integer = 50

Private Type TDefect
    Korean  As String
    English As String
    Defect  As String
End Type

Dim m_bLoading     As Boolean
Dim m_bSortForward As Boolean

Dim m_sTotalField(6)  As String             ' 리포트 Title
Dim m_nDefectName(DEFECT_COUNT) As TDefect
Dim m_nSelected%

Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 11985, 9660

    Show

    Call SetOperate(Me)
    Call InitGrid
    Call InitGroup

    cmdFind.Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    For i = 0 To DEFECT_COUNT
        m_nDefectName(i).Korean = ""
        m_nDefectName(i).English = ""
        m_nDefectName(i).Defect = ""
    Next i

    dtpDate(0).Enabled = False
    dtpDate(1).Enabled = False
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
    cmdFind.Enabled = False

    dtpDate(0) = Now
    dtpDate(1) = Now
    chkSearch(0).Value = vbChecked
    m_nSelected = 0
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
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
    Else
        If chkSearch(Index) Then
            If Index = 1 Then cmdFind.Enabled = True
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
        Else
            If Index = 1 Then cmdFind.Enabled = False
            txtSearch(Index).Enabled = False
            cmdSearch.SetFocus
        End If
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
        If Index = 1 Then Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdFind_Click()
    Call ReturnCode(LG_CUSTOM, 0, True, txtSearch(1))
End Sub

Private Sub cmdSearch_Click()
    Call FillGrid
End Sub

Private Sub grdData_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With grdData
        If .Rows = .FixedRows Or .MouseRow < 0 Or .MouseRow >= .FixedRows Then Exit Sub

        Call SortGrid(grdData, .MouseCol, m_bSortForward)
        m_bSortForward = Not m_bSortForward

        Call FillGridResult
    End With
End Sub

Private Sub grdData_RowColChange()
    If m_bLoading Then Exit Sub

    Call FillGridResult
End Sub

Private Sub grdResult_RowColChange()
    If grdResult.Rows = grdResult.FixedRows Then Exit Sub

    If FillGridDefect() Then
        If (grdResult.Row - grdResult.TopRow + 2) > (LIMIT_ROW2 / 2 - 4) Then
            grdDefect.Move 7600, 1800
        Else
            grdDefect.Move 7600, 5200
        End If
        grdDefect.Visible = True
    Else
        grdDefect.Visible = False
    End If
End Sub

Private Sub grdResult_DblClick()
    With grdResult
        If .Row < .FixedRows Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub

Private Sub grdResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        With grdResult
            If .Rows = .FixedRows Then Exit Sub
        End With

        Call CheckResultRow
    End If
End Sub

Private Sub grdResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With grdResult
        If .Rows = .FixedRows Or .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
    End With

    Call CheckResultRow
End Sub

Private Sub optOrder_Click(Index As Integer)
    If optOrder(1).Value = True Then
        With grdData
            .ColWidth(2) = 1485
            .ColWidth(3) = 0
        End With

        chkSearch(2).Caption = "관리번호"
    Else
        With grdData
            .ColWidth(2) = 0
            .ColWidth(3) = 1485
        End With

        chkSearch(2).Caption = "OrderNo."
    End If
End Sub

Private Sub optMain_Click(Index As Integer)
    fraRange.Visible = IIf(optMain(0) = vbChecked, True, False)
End Sub

Private Sub cmdPrint_Click()
    If grdData.Rows = grdData.FixedRows Then
        Call MessageBox(LoadResString(111))
        Exit Sub
    End If

    Call PrintDefectPosition
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 7
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .TextArray(1) = "완" & vbCrLf & "료":       .ColWidth(1) = 300:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "관리번호":                 .ColWidth(2) = 1485:            .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "OrderNo.":                 .ColWidth(3) = 0:               .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처명":                 .ColWidth(4) = LIMIT_WIDTH1:    .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "영문거래처명":             .ColWidth(5) = 0
        .TextArray(6) = "품명":                     .ColWidth(6) = 0

        .Redraw = flexRDDirect
    End With

    With grdDefect
        .Cols = 6
        Call SetVSFlexGrid(grdDefect)

        .Redraw = flexRDNone

        .Rows = .FixedRows
        .RowHeightMin = 275
        .Width = 3660
        .TextArray(1) = "불량명":       .ColWidth(1) = LIMIT_WIDTH3:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "영문명":       .ColHidden(2) = True
        .TextArray(3) = "Tag":          .ColWidth(3) = 700:             .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "위치":         .ColWidth(4) = 700:             .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "벌점":         .ColWidth(5) = 700:             .ColAlignment(5) = flexAlignRightCenter
        
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0.0"

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub InitGroup()
    With grdResult
        .Redraw = flexRDNone
        .Cols = 15
        
      '  Call SetVSFlexGrid(grdResult)
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1

        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarVertical
        .ExplorerBar = flexExSortShow
        .ScrollTrack = True
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 450
        .ColWidth(0) = 360
        .RowHeightMin = 300


        .TextArray(0) = "":                         .ColAlignment(0) = flexAlignCenterCenter:       .ColWidth(0) = 300
        .TextArray(1) = "색상명" & vbCrLf & "LotNO-RollNo":  .ColAlignment(1) = flexAlignLeftCenter:     .ColWidth(1) = LIMIT_WIDTH2 '1335
        .TextArray(2) = "검사일자":                 .ColAlignment(2) = flexAlignCenterCenter:       .ColWidth(2) = 1100
        .TextArray(3) = "수주량":                   .ColAlignment(3) = flexAlignRightCenter:        .ColWidth(3) = 800:     .ColFormat(3) = GetFormat()
        .TextArray(4) = "검사량":                   .ColAlignment(4) = flexAlignRightCenter:        .ColWidth(4) = 900:     .ColFormat(4) = GetFormat(g_nPointPos)
        .TextArray(5) = "합격량":                   .ColAlignment(5) = flexAlignRightCenter:        .ColWidth(5) = 800:     .ColFormat(5) = GetFormat(g_nPointPos)
        .TextArray(6) = "불량갯수":                 .ColAlignment(6) = flexAlignRightCenter:        .ColWidth(6) = 800:     .ColFormat(6) = GetFormat()
        .TextArray(7) = "보상":                     .ColAlignment(7) = flexAlignRightCenter:        .ColWidth(7) = 500:     .ColFormat(7) = GetFormat(1)
        .TextArray(8) = "견본":                     .ColAlignment(8) = flexAlignRightCenter:        .ColWidth(8) = 500:     .ColFormat(8) = GetFormat(g_nPointPos)
        .TextArray(9) = "난단":                     .ColAlignment(9) = flexAlignRightCenter:        .ColWidth(9) = 500:     .ColFormat(9) = GetFormat(g_nPointPos)
        .TextArray(10) = "검사" & vbCrLf & "절수":  .ColAlignment(10) = flexAlignRightCenter:       .ColWidth(10) = 660:    .ColFormat(10) = GetFormat()
        .TextArray(11) = "Sort ColorID":            .ColWidth(11) = 0
        .TextArray(12) = "Sort 검사일자":           .ColWidth(12) = 0
        .TextArray(13) = "RollNo":                  .ColWidth(13) = 0
        .TextArray(14) = "ReollID":                .ColWidth(14) = 0
        
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 5) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 6) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 7) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 8) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 9) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 10) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 11) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 12) = flexAlignCenterCenter

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGrid()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim i%, lNowRow&, lWithDash&

    On Error GoTo ErrHandler

    m_bLoading = True

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetOrder(IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, IIf(optOrder(0), 2, 1), 0), txtSearch(2), 0, 0, "")
    Set oInspect = Nothing

    With grdData
        .Redraw = flexRDNone

        lNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & IIf(IsNull(rs!CloseDate), "", "*") & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!OrderNo & vbTab & rs!KCustom & vbTab & CheckNull(rs!ECustom) & vbTab & rs!Article

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            lblCount.Caption = LoadResString(250) & grdData.Rows - 1 & " 건"

            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            lblCount.Caption = LoadResString(250)

            .HighLight = flexHighlightNever
        End If

        Call ChangeScrollData

        .Redraw = flexRDDirect
    End With

    m_bLoading = False

    Call FillGridResult

    Exit Sub

ErrHandler:
    m_bLoading = False

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGridResult()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim iCol(8) As Integer
    Dim i%, iTop%
    Dim nBeforeTop%
    Dim nPass%, nUnPass%
    
    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetDefect(MakeOrderID(grdData.TextMatrix(grdData.Row, 2), OM_REDUCE), _
                            IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)))
    Set oInspect = Nothing

    With grdResult
        .Redraw = flexRDNone

        .Rows = 1
        iTop = 1

        For i = 0 To 7
            iCol(i) = i + 3
        Next i

        Do Until rs.EOF
            If rs!ColorID <> .TextMatrix(.Rows - 1, 12) Then
                .AddItem "" & vbTab & CheckNull(rs!ColorID) & ". " & CheckNull(rs!Color) & vbTab & _
                    "" & vbTab & Format(rs!ColorQty, "#,##0") & vbTab & "0" & vbTab & "0" & vbTab & " " & vbTab & _
                    "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & _
                    CheckNull(rs!ColorID) & vbTab & "" & vbTab & rs!RollNO & vbTab & "0"

                Call DoFlexGridGroup(grdResult, .Rows - 1, 1)
                Call GridCollapse(nBeforeTop)       ' 서브토탈 row를 접힌 상태로 출력
                
                nBeforeTop = .Rows - 1
                
                iTop = .Rows - 1
            End If
        
            .AddItem "" & vbTab & CheckNull(rs!LotNo) & "-" & CheckNull(rs!RollNO) & vbTab & MakeDate(DF_LONG, rs!ExamDate) & vbTab & _
                rs!Grade & vbTab & rs!CtrlQty & vbTab & CStr(rs!CtrlQty - rs!CutQty) & vbTab & _
                rs!DefectQty & vbTab & rs!LossQty & vbTab & rs!SampleQty & vbTab & rs!CutQty & vbTab & vbTab & _
                rs!ColorID & vbTab & rs!ExamDate & vbTab & rs!RollNO & vbTab & rs!RollID

            .TextMatrix(iTop, iCol(1)) = CStr(CSng(.TextMatrix(iTop, iCol(1))) + rs!CtrlQty)
            .TextMatrix(iTop, iCol(2)) = CStr(CSng(.TextMatrix(iTop, iCol(2))) + (rs!CtrlQty - rs!CutQty))
            .TextMatrix(iTop, iCol(4)) = CStr(CSng(.TextMatrix(iTop, iCol(4))) + rs!LossQty)
            .TextMatrix(iTop, iCol(5)) = CStr(CSng(.TextMatrix(iTop, iCol(5))) + rs!SampleQty)
            .TextMatrix(iTop, iCol(6)) = CStr(CSng(.TextMatrix(iTop, iCol(6))) + rs!CutQty)
            .TextMatrix(iTop, iCol(7)) = CStr(CSng(.TextMatrix(iTop, iCol(7)))) + 1

            rs.MoveNext
        Loop
        
        Call GridCollapse(nBeforeTop)
        
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            pnlMsg.Visible = True
        Else
            pnlMsg.Visible = False
        End If

        Call ChangeScrollResult

        .Redraw = flexRDDirect
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Function FillGridDefect() As Boolean
    Dim oInspect As PlusLib2.CInspect
    Dim rs As ADODB.Recordset
    Dim i%, iNowRow%

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetInspectSub(MakeOrderID(grdData.TextMatrix(grdData.Row, 2), OM_REDUCE), grdResult.TextMatrix(grdResult.Row, 14))
    Set oInspect = Nothing

    With grdDefect
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        For i = 1 To rs.RecordCount
            .AddItem CStr(.Rows) & vbTab & rs!KDefect & vbTab & rs!EDefect & vbTab & rs!TagName & vbTab & rs!YPos & vbTab & _
                rs!Demerit

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            If .Rows < LIMIT_ROW3 Then
                .Height = (.RowHeight(.FixedRows) + 45) * .Rows + 340
                .ScrollBars = flexScrollBarNone
            Else
                .Height = 3000
                .ScrollBars = flexScrollBarVertical
            End If

            If .Rows > iNowRow Then
                .Row = iNowRow
            Else
                .Row = .Rows - 1
            End If

            .HighLight = flexHighlightAlways

            .Col = .FixedCols
            .ColSel = .Cols - 1

            FillGridDefect = True
        Else
            FillGridDefect = False
        End If

        Call ChangeScrollDefect

        .Redraw = flexRDDirect
    End With

    Exit Function
    
ErrHandler:
    FillGridDefect = False

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Function

Private Sub ChangeScrollData()
    With grdData
        .ColWidth(4) = LIMIT_WIDTH1 - IIf(.Rows > LIMIT_ROW1, 240, 0)
    End With
End Sub

Private Sub ChangeScrollResult()
    With grdResult
        .ColWidth(1) = LIMIT_WIDTH2 - IIf(.Rows > LIMIT_ROW2, 240, 0)
    End With
End Sub

Private Sub ChangeScrollDefect()
    With grdDefect
        .ColWidth(1) = LIMIT_WIDTH3 - IIf(.Rows > LIMIT_ROW3, 240, 0)
    End With
End Sub

Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(iRow) = True

        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = vbBlue
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = COLOR_GRIDROW
            .Cell(flexcpChecked, iRow, 0) = flexUnchecked
            '.ColDataType(0) = flexDTBoolean
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        End Select
    End With
End Sub



Private Sub GridCollapse(Row As Integer)
    
    With grdResult
    
        If Row >= .FixedRows Then
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub


Private Sub CheckResultRow()
    With grdResult
        If .IsSubtotal(.Row) = True Then
            If .Cell(flexcpChecked, .Row, 0) = flexUnchecked Then
               .Cell(flexcpChecked, .Row, 0) = flexChecked
               m_nSelected = m_nSelected + 1
            Else
               .Cell(flexcpChecked, .Row, 0) = flexUnchecked
               m_nSelected = m_nSelected - 1
            End If
       End If
    End With
End Sub

Private Sub PrintDefectPosition()
    Dim oOrder   As PlusLib2.COrder
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim rsData   As ADODB.Recordset
    Dim sOrderID$, sGrade$, sTemp As String * 6
    Dim nRoll%, nPenalty%
    Dim i%, j%, k%
    Dim sColor() As String
    Dim sColorID() As String
    Dim sRollNo() As String

    On Error GoTo ErrHandler

    If m_nSelected <= 0 Then
        Call MessageBox("선택한 색상이 없습니다.")
        m_nSelected = 0
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    ReDim sColorID(m_nSelected)
    ReDim sColor(m_nSelected)
    ReDim sRollNo(m_nSelected)
    j = 0
    With grdResult
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = vbChecked Then
                sColor(j) = Mid(.TextMatrix(i, 1), 5)      ' Color
                sColorID(j) = .TextMatrix(i, 11)
                sRollNo(j) = .TextMatrix(i, 13)
                j = j + 1
            End If
        Next i
    End With
    
    Call GetDefect


    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    sOrderID = MakeOrderID(grdData.TextMatrix(grdData.Row, 2), OM_REDUCE)
    Set rs = oOrder.GetOrderOne(sOrderID)
    Set oOrder = Nothing
    
    m_sTotalField(0) = grdData.TextMatrix(grdData.Row, 2)           ' 관리번호
    If optLang(0).Value = True Then
       m_sTotalField(1) = grdData.TextMatrix(grdData.Row, 4)        ' 한글 거래처
    Else
       m_sTotalField(1) = grdData.TextMatrix(grdData.Row, 5)        ' 영문 거래처
    End If
    m_sTotalField(3) = grdData.TextMatrix(grdData.Row, 3)           ' Order No
    m_sTotalField(4) = rs!Article
    m_sTotalField(5) = IIf(rs!OrderUnit = "0", "Yards", "Meters")
    m_sTotalField(6) = rs!Width

    rs.Close
    Set rs = Nothing

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    For i = 0 To m_nSelected - 1
        If Len(sColor(i)) = 0 Then
            Exit For
        End If
        m_sTotalField(2) = sColor(i)
    
        ' Report 양식 출력
        Call ReportForm(IIf(optRange(0).Value = True, 1, 2), IIf(optLang(0).Value = True, 1, 2))

        nRoll = 0

        Set rs = oInspect.GetInspect(sOrderID, sColorID(i))
        If rs.EOF Then
            Call MessageBox("해당 검사실적이 없습니다.")
            Printer.KillDoc
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If

        With rs
            Do Until .EOF
                ' NewPage
                If nRoll = 10 Then
                    nRoll = 0
                    Printer.NewPage
                    Call ReportForm(IIf(optRange(0).Value = True, 1, 2), IIf(optLang(0).Value = True, 1, 2))
                End If

                Printer.Font.Size = 7

                ' Roll No
                Printer.CurrentX = BASE_X + 1300 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 700
                Printer.Print !RollNO
                
                '투입량
                Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11000
                RSet sTemp = CStr(!StuffQty)
                Printer.Print sTemp
                '실폭
                Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11000
                If Not IsNull(!Width) Then
                    RSet sTemp = Format(Val(!Width), "#0.0")
                Else
                    RSet sTemp = Format(0, "#0.0")
                End If
                Printer.Print sTemp
                
                '계측량(실제검사 수량)
                Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11300
                RSet sTemp = Str(IIf(!Realqty = 0, 0, !Realqty))
                Printer.Print sTemp
                '실수량(조정수량)
                Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11300
                RSet sTemp = Str(!CtrlQty)
                Printer.Print sTemp
                
                'LOSS
                Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11600
                RSet sTemp = Str(!LossQty)
                Printer.Print sTemp
                '불량수
                Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11600
                RSet sTemp = Str(IIf(IsNull(!DefectQty), 0, !DefectQty)) '불량수
                Printer.Print sTemp
                
                '견본
                Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 12200
                RSet sTemp = Str(!SampleQty)
                Printer.Print sTemp
                '난단
                Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 12200
                RSet sTemp = Str(!CutQty)
                Printer.Print sTemp
                
                
                ' 등급
                'POINT
                Select Case !GradeID
                    Case "1":   sGrade = "A"
                    Case "2":   sGrade = "B"
                    Case "3":   sGrade = "C"
                    Case "4":   sGrade = "D"
                    Case "5":   sGrade = "E"
                    Case "6":   sGrade = "F"
                End Select
                
                Printer.CurrentX = BASE_X + (1000 * (nRoll + 1)) + 450
                Printer.CurrentY = BASE_Y + 12500
                Printer.Print sGrade
                
                '대표불량
                Printer.CurrentX = BASE_X + 1160 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 12800

                If Not IsNull(!DefectID) Then
                    Printer.Print rs!Defect
                End If
                
                ' 불량 Detail
                nPenalty = 0

                Set rsData = oInspect.GetInspectSub(sOrderID, rs!RollID) 'sRollNo(i))
                Do Until rsData.EOF
                    Printer.Font.Size = 5
                    Printer.CurrentX = BASE_X + 1200 + (1000 * nRoll)

                    If optRange(0).Value = True Then
                        Printer.CurrentY = BASE_Y + 900 + (Val(rsData!YPos) * 100)
                    Else
                        Printer.CurrentY = BASE_Y + 900 + (Val(rsData!YPos) * 200)
                    End If
                    Printer.Print m_nDefectName(Val(Right(rsData!DefectID, 2))).Defect & Str(Val(rsData!Demerit) / 10) & ""
                    
                    nPenalty = nPenalty + rsData!Demerit
                    rsData.MoveNext
                Loop
                rsData.Close
                Set rsData = Nothing


                '불량수  벌점
                Printer.Font.Size = 7
                Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11900
                RSet sTemp = Str(IIf(IsNull(!DefectQty), 0, !DefectQty))
                Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
                Printer.CurrentY = BASE_Y + 11900
                RSet sTemp = CheckNull(!LotNo) 'LOTNo
                Printer.Print sTemp
          
'
'                Printer.CurrentX = BASE_X + 1100 + (1000 * nRoll)
'                Printer.CurrentY = BASE_Y + 12500
'                RSet sTemp = Str(nPenalty / 10)
'                Printer.Print sTemp
'
'                RSet sTemp = (Format(!CalcValue1, "0.00"))
'                Printer.CurrentX = BASE_X + 1570 + (1000 * nRoll)
'                Printer.CurrentY = BASE_Y + 12500
'                Printer.Print sTemp

                nRoll = nRoll + 1
                .MoveNext
            Loop
            .Close
        End With
        Set rs = Nothing
        Printer.EndDoc
    Next i

    Set oOrder = Nothing
    Set oInspect = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    
    Printer.KillDoc
    Set oInspect = Nothing
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub ReportForm(NewRange As Integer, NewLang As Integer)
    Dim i%, j%, cnt%
    
    On Error GoTo ErrHandler

    Printer.DrawWidth = 3
    Printer.Font.Name = "바탕"

    ' Title명
    Printer.Font.Size = 20
    Printer.CurrentX = BASE_X + 3450
    Printer.CurrentY = BASE_Y - 800

    If NewLang = 1 Then
         Printer.Print "       검사결과표"
    Else
         Printer.Print "INSPECTION REPORT"
    End If

    ' 출력일
    Printer.Font.Size = 10
    Printer.CurrentY = BASE_Y - 300
    If NewLang = 1 Then
        Printer.CurrentX = BASE_X + 9000
        Printer.Print "출력일 : " & Format(Date, "yyyy/MM/dd")
    Else
        Printer.CurrentX = BASE_X + 8400
        Printer.Print "Printed Date : " & Format(Date, "yyyy/MM/dd")
    End If

    Printer.Font.Size = 8
    Printer.CurrentX = BASE_X + 230
    Printer.CurrentY = BASE_Y + 100
    If NewLang = 1 Then
        Printer.Print "접수번호                                                       거래처                                                                     칼라명  "
    Else
        Printer.Print "INNER NO                                                     BUYER                                                                  COLOR   "
    End If

    Printer.CurrentX = BASE_X + 230
    Printer.CurrentY = BASE_Y + 400
    If NewLang = 1 Then
        Printer.Print "Order No.                                                      품 명                                                                      단위"
    Else
        Printer.Print "Order No.                                                      ITEM                                                                      UNIT"
    End If

    Printer.CurrentX = BASE_X + 230
    Printer.CurrentY = BASE_Y + 700
    Printer.Print "Roll No"

    ' Box
    Printer.Line (BASE_X, BASE_Y)-(BASE_X, BASE_Y + 14000)
    Printer.Line (BASE_X, BASE_Y)-(BASE_X + 11000, BASE_Y)
    Printer.Line (BASE_X + 11000, BASE_Y)-(BASE_X + 11000, BASE_Y + 14000)
    Printer.Line (BASE_X, BASE_Y + 14000)-(BASE_X + 11000, BASE_Y + 14000)

    ' 항목의 Y축 Line
    Printer.Line (BASE_X + 1100, BASE_Y)-(BASE_X + 1100, BASE_Y + 600)
    Printer.Line (BASE_X + 3500, BASE_Y)-(BASE_X + 3500, BASE_Y + 600)
    Printer.Line (BASE_X + 4700, BASE_Y)-(BASE_X + 4700, BASE_Y + 600)
    Printer.Line (BASE_X + 7600, BASE_Y)-(BASE_X + 7600, BASE_Y + 600)
    Printer.Line (BASE_X + 8600, BASE_Y)-(BASE_X + 8600, BASE_Y + 600)

    ' X, Y축의 Line
    Printer.Line (BASE_X, BASE_Y + 300)-(BASE_X + 11000, BASE_Y + 300)
    Printer.Line (BASE_X, BASE_Y + 600)-(BASE_X + 11000, BASE_Y + 600)
    Printer.Line (BASE_X, BASE_Y + 900)-(BASE_X + 11000, BASE_Y + 900)
    For i = 0 To 10
        Printer.Line (BASE_X + (1000 * i), BASE_Y + 600)-(BASE_X + (1000 * i), BASE_Y + 13000)
    Next i
    Printer.Line (BASE_X, BASE_Y + 10900)-(BASE_X + 11000, BASE_Y + 10900)
    Printer.Line (BASE_X, BASE_Y + 11200)-(BASE_X + 11000, BASE_Y + 11200)
    Printer.Line (BASE_X, BASE_Y + 11500)-(BASE_X + 11000, BASE_Y + 11500)
    Printer.Line (BASE_X, BASE_Y + 11800)-(BASE_X + 11000, BASE_Y + 11800)
    Printer.Line (BASE_X, BASE_Y + 12100)-(BASE_X + 11000, BASE_Y + 12100)
    Printer.Line (BASE_X, BASE_Y + 12400)-(BASE_X + 11000, BASE_Y + 12400)
    Printer.Line (BASE_X, BASE_Y + 12700)-(BASE_X + 11000, BASE_Y + 12700)
    Printer.Line (BASE_X, BASE_Y + 13000)-(BASE_X + 11000, BASE_Y + 13000)

    ' 항목명
    Printer.Font.Size = 7
    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 11000
    
    If NewLang = 1 Then
        Printer.Print "투입량  실폭"
    Else
        Printer.Print "IN Q'Y  Width"
    End If

    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 11300
    If NewLang = 1 Then
        Printer.Print "계측량 실수량"
    Else
        Printer.Print " Q'TY   NET "
    End If

    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 11600
    If NewLang = 1 Then
        Printer.Print " LOSS 불량수"
    Else
        Printer.Print " LOSS Defect"
    End If
    ' LOT
    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 11900
    If NewLang = 1 Then
        Printer.Print "          LOT"
    Else
        Printer.Print "          LOT"
    End If

    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 12200
    If NewLang = 1 Then
        Printer.Print "  견본   난단"
    Else
        Printer.Print "Sample Short"
    End If

    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 12500
    If NewLang = 1 Then
        Printer.Print "     등 급   "
    Else
        Printer.Print "     POINT   "
    End If

    Printer.CurrentX = BASE_X + 60
    Printer.CurrentY = BASE_Y + 12800
    If NewLang = 1 Then
        Printer.Print "   대표불량"
    Else
        Printer.Print " Main Defect"
    End If
    
    Printer.CurrentX = BASE_X + 4650
    Printer.CurrentY = BASE_Y + 14100
    Printer.FontSize = 10
        
    Printer.Print "선광 염직"

    Printer.Font.Size = 8
    ' 눈금
    For i = 1 To 10
        For j = 1 To 99
            If j Mod 5 = 0 Then
                If i = 1 Then
                    If j * NewRange < 100 Then
                        Printer.CurrentX = BASE_X + 700
                    Else
                        Printer.CurrentX = BASE_X + 600
                    End If
                    Printer.CurrentY = BASE_Y + 840 + (100 * j)
                    Printer.Print j * NewRange
                End If
                Printer.Line (BASE_X + (1000 * i), BASE_Y + 900 + (100 * j))-(BASE_X + (1000 * i) + 80, BASE_Y + 900 + (100 * j))
            Else
                Printer.Line (BASE_X + (1000 * i), BASE_Y + 900 + (100 * j))-(BASE_X + (1000 * i) + 40, BASE_Y + 900 + (100 * j))
            End If
        Next j
    Next i

    ' 점선
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbDot
    For i = 0 To 10
        Printer.Line (BASE_X + (1000 * i) + 500, BASE_Y + 10900)-(BASE_X + (1000 * i) + 500, BASE_Y + 12400)
'        If i <> 0 Then
'            Printer.CurrentX = BASE_X + (1000 * i) + 500
'            Printer.CurrentY = BASE_Y + 12500
'            Printer.Print "/"
'        End If
    Next i
    Printer.DrawStyle = vbSolid

    ' Title 항목 데이타 (관리번호)
    Printer.CurrentX = BASE_X + 1500
    Printer.CurrentY = BASE_Y + 100
    Printer.Print m_sTotalField(0)
    
    ' Title 항목 데이타 (거래처)
    Printer.CurrentX = BASE_X + 5100
    Printer.CurrentY = BASE_Y + 100
    Printer.Print m_sTotalField(1)
    
    ' Title 항목 데이타 (색상명)
    Printer.CurrentX = BASE_X + 9000
    Printer.CurrentY = BASE_Y + 100
    Printer.Print m_sTotalField(2)
    
    ' Title 항목 데이타 (Order No)
    Printer.CurrentX = BASE_X + 1500
    Printer.CurrentY = BASE_Y + 400
    Printer.Print m_sTotalField(3)
    
    ' Title 항목 데이타 (품명)
    Printer.CurrentX = BASE_X + 5100
    Printer.CurrentY = BASE_Y + 400
    Printer.Print m_sTotalField(4)
    
    ' Title 항목 데이타 (단위)
    Printer.CurrentX = BASE_X + 9000
    Printer.CurrentY = BASE_Y + 400
    Printer.Print m_sTotalField(5)
    
    ' 불량명 인쇄
    Printer.Font.Size = 7
    ' 제직불량
    For i = 1 To 20
        
        If m_nDefectName(i).Defect = "" Then Exit For
        Printer.CurrentX = BASE_X + 200 + (1550 * ((i - 1) Mod 7))
        Printer.CurrentY = BASE_Y + 13100 + (160 * Int((i - 1) / 7))
        If NewLang = 1 Then
            Printer.Print m_nDefectName(i).Defect & "-" & m_nDefectName(i).Korean
        Else
            Printer.Print m_nDefectName(i).Defect & "-" & m_nDefectName(i).English
        End If
    Next i
    ' 가공불량
    cnt = i
    For i = 21 To DEFECT_COUNT
        If cnt > DEFECT_COUNT Then Exit For
        If m_nDefectName(i).Defect = "" Then Exit For
        Printer.CurrentX = BASE_X + 200 + (1550 * ((cnt - 1) Mod 7))
        Printer.CurrentY = BASE_Y + 13100 + (160 * Int((cnt - 1) / 7))
        If NewLang = 1 Then
            Printer.Print m_nDefectName(i).Defect & "-" & m_nDefectName(i).Korean
        Else
            Printer.Print m_nDefectName(i).Defect & "-" & m_nDefectName(i).English
        End If
        cnt = cnt + 1
    Next i

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub GetDefect()
    Dim oCode As PlusLib2.CCode
    Dim rs    As ADODB.Recordset
    Dim nDefectNo%

    On Error GoTo ErrHandle

    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
    oCode.CodeType = CD_DEFECT

    Set rs = oCode.GetCode()
    Set oCode = Nothing

    Do Until rs.EOF
        nDefectNo = Val(Right(rs!DefectID, 2))
        If Val(nDefectNo) < DEFECT_COUNT Then
            m_nDefectName(nDefectNo).Korean = Trim(rs!KDefect)
            m_nDefectName(nDefectNo).English = CheckNull(rs!EDefect)
            m_nDefectName(nDefectNo).Defect = CheckNull(rs!TagName)
        End If

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub

ErrHandle:
    Set rs = Nothing
    Set oCode = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub
