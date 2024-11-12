VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanCPBView 
   Caption         =   "C.P.B 염색계획 조회"
   ClientHeight    =   9255
   ClientLeft      =   -240
   ClientTop       =   645
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15255
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7545
      Left            =   30
      TabIndex        =   24
      Top             =   900
      Width           =   15225
      _cx             =   26855
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
      Begin Threed.SSPanel pnlProgress 
         Height          =   870
         Left            =   1800
         TabIndex        =   26
         Top             =   1530
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
            TabIndex        =   27
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
            TabIndex        =   28
            Top             =   120
            Width           =   270
         End
      End
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1588
      _Version        =   196609
      Begin VB.ComboBox cboProcessID 
         Height          =   300
         Left            =   7305
         Style           =   2  '드롭다운 목록
         TabIndex        =   30
         Top             =   120
         Width           =   1905
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   5955
         TabIndex        =   29
         Top             =   120
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "공정"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1695
         MousePointer    =   99  '사용자 정의
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   435
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1695
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   600
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   11325
         TabIndex        =   4
         Top             =   105
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   11325
         TabIndex        =   3
         Top             =   435
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   7305
         TabIndex        =   2
         Top             =   465
         Width           =   1905
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   14325
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 저장"
         Top             =   45
         Width           =   780
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   795
         Left            =   180
         TabIndex        =   7
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1402
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   480
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3660
         TabIndex        =   10
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117309441
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3660
         TabIndex        =   11
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   117309441
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   2340
         TabIndex        =   12
         Top             =   150
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
            Caption         =   "계획일자"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   9945
         TabIndex        =   14
         Top             =   105
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
            TabIndex        =   15
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   13260
         TabIndex        =   16
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
         Index           =   3
         Left            =   9945
         TabIndex        =   17
         Top             =   435
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
         Left            =   13260
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   435
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
         Left            =   5955
         TabIndex        =   20
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
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1185
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   4965
         TabIndex        =   23
         Top             =   570
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4965
         TabIndex        =   22
         Top             =   210
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13590
      TabIndex        =   25
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
      Left            =   11880
      TabIndex        =   31
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   360
      Left            =   30
      TabIndex        =   32
      Top             =   8490
      Width           =   4560
      _cx             =   8043
      _cy             =   635
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
Attribute VB_Name = "frmPlanCPBView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ProcessID As String

Public Sub LoadCPBView(ByVal pProcID As Integer)
    Me.Show
    cboProcessID.ListIndex = pProcID  '0: 4000(c염색), 1:Rapid 염색으로 설정
    dtpDate(0).SetFocus

End Sub




Private Sub cboProcessID_Click()
    Select Case cboProcessID.ListIndex
        Case 0: m_ProcessID = "4000"
            grdData.ColHidden(6) = False
            grdData.ColHidden(7) = False
        Case 1: m_ProcessID = "4300"
            grdData.ColHidden(6) = True
            grdData.ColHidden(7) = True
    End Select
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
    End If
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

Private Sub cmdPrint_Click()
    Call ColResize("-")
    With grdData
        .Redraw = flexRDBuffered
    
        .GridLinesFixed = flexGridNone
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "염 색   계 획(" & cboProcessID.Text & ")"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 1, 1, 1, 3) = "▶ 계획일 : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD")
        .Cell(flexcpAlignment, 1, 1, 1, 3) = flexAlignLeftCenter
        .Cell(flexcpText, 1, 9, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD HH:SS")
        .Cell(flexcpAlignment, 1, 9, 1, .Cols - 1) = flexAlignRightCenter
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        .ColWidth(3) = 1500
        .ColWidth(4) = 1700
        
        .ColWidth(1) = 1200
        
        .PrintGrid "태을염직", True, 2, 100, 500

        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True

        .ColWidth(3) = 1200
        .ColWidth(4) = 1400


        .Redraw = flexRDDirect
    End With
    Call ColResize("+")
End Sub

Sub ColResize(ByVal pType As String)
    Dim II%, JJ As Integer
    
    If pType = "-" Then
        With grdData
            For II = 0 To .Cols - 1
                .ColWidth(II) = .ColWidth(II) * 0.8
            Next II
        End With
    Else
        With grdData
            For II = 0 To .Cols - 1
                .ColWidth(II) = .ColWidth(II) / 0.8
            Next II
        End With
    End If
    
''    With grdData
''        .Redraw = flexRDBuffered
''        .ColWidth(0) = .ColWidth(0) + 360 * JJ
''        .ColWidth(1) = .ColWidth(1) + 50 * JJ
''        .ColWidth(2) = .ColWidth(2) + 200 * JJ
''        .ColWidth(3) = .ColWidth(3) + 200 * JJ
''        .ColWidth(4) = .ColWidth(4) + 200 * JJ
''        .ColWidth(5) = .ColWidth(5) + 50 * JJ
''        .ColWidth(6) = .ColWidth(6) + 50 * JJ
''        .ColWidth(7) = .ColWidth(7) + 150 * JJ
''        .ColWidth(8) = .ColWidth(8) + 100 * JJ
''        .ColWidth(9) = .ColWidth(9) + 50 * JJ
''
''
'''        .TextMatrix(3, 1) = "일자":             .ColWidth(1) = 600:                 .ColAlignment(1) = flexAlignCenterCenter
'''        .TextMatrix(3, 2) = "거래처명":         .ColWidth(2) = 1800:                .ColAlignment(2) = flexAlignLeftCenter
'''        .TextMatrix(3, 3) = "실 입고처":        .ColWidth(3) = 1200:                .ColAlignment(3) = flexAlignLeftCenter
'''        .TextMatrix(3, 4) = "품명":             .ColWidth(4) = 2400:                .ColAlignment(4) = flexAlignLeftCenter
'''        .TextMatrix(3, 5) = "관리번호":         .ColWidth(5) = 1300:                .ColAlignment(5) = flexAlignCenterCenter
'''        .TextMatrix(3, 6) = "OrderNO":          .ColWidth(6) = 1300:                .ColAlignment(6) = flexAlignLeftCenter
'''        .TextMatrix(3, 7) = "가공":             .ColWidth(7) = 1000:                 .ColAlignment(7) = flexAlignCenterCenter
'''        .TextMatrix(3, 8) = "절 수":            .ColWidth(8) = 800:                 .ColAlignment(8) = flexAlignRightCenter
'''        .TextMatrix(3, 9) = "수   량":          .ColWidth(9) = 900:                 .ColAlignment(9) = flexAlignRightCenter
''
''        .Redraw = flexRDDirect
''    End With

End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
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
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15360, 9840

    Call SetOperate(Me)
    Call InitGrid

    For i = 0 To 1
        dtpDate(i) = Now
    Next i
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    
    pnlProgress.Visible = False
    
    '공정선택 콤보박스에 C염색(C.P.B염색), 염색(Rapid염색)으로 설정하는 프로시저 호출
    Call SetProcessID(cboProcessID, "'4000', '4300'")
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub optOrder_Click(Index As Integer)
    If optOrder(0).Value Then
        chkSearch(3).Caption = "Order No"
    Else
        chkSearch(3).Caption = "관리번호"
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
    With grdData
        .Cols = 11
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 3
        .FixedRows = 3

        .RowHidden(0) = True
        .RowHidden(1) = True

        .TextMatrix(2, 0) = " "
        .TextMatrix(2, 1) = "계획일자":     .ColWidth(1) = 1200:             .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(2, 2) = "거래처":       .ColWidth(2) = 1500:             .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(2, 3) = "관리번호":     .ColWidth(3) = 1200:             .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(2, 4) = "OrderNo":      .ColWidth(4) = 1400:             .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(2, 5) = "품명":         .ColWidth(5) = 2700:             .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(2, 6) = "색상명":       .ColWidth(6) = 2000:             .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(2, 7) = "수량(YD)":     .ColWidth(7) = 800:              .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(2, 8) = "계획":         .ColWidth(8) = 600:              .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(2, 9) = "긴급":         .ColWidth(9) = 600:              .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(2, 10) = "내역":        .ColWidth(10) = 2000:            .ColAlignment(10) = flexAlignLeftCenter
        
        .Redraw = flexRDDirect
    End With
    
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 2
        .ExtendLastCol = True
        
        .RowHeight(0) = 300
        .TextArray(0) = "합계":           .ColWidth(0) = 2000:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "YD:              .ColWidth(1) = 700:    .ColAlignment(1) = flexAlignRightCenter"
        
        .RowHeight(0) = 300
        .Redraw = flexRDDirect
    End With
    
End Sub


Private Sub FillGridData()
    Dim oPlanCPB As PlusLib2.CPlanCPB
    Dim rs As ADODB.Recordset
    Dim i%, nNoPlanQty#, nTotQty As Long
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oPlanCPB = New PlusLib2.CPlanCPB
    oPlanCPB.Connection = g_adoCon
    
    Set rs = oPlanCPB.GetPlanCPBView(m_ProcessID, IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                 IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3))
    Set oPlanCPB = Nothing
        
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(i + 1) & vbTab & MakeDate(DF_LONG, rs!Plandate) & vbTab & rs!kCustom & vbTab & _
                    MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    Trim(rs!Article) & vbTab & Trim(rs!ColorName) & vbTab & rs!Qty & vbTab & Trim(rs!EmerClss) & vbTab & _
                    Trim(rs!PlanClss) & vbTab & Trim(rs!Remark)
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            nTotQty = nTotQty + rs!Qty
            rs.MoveNext
        Next i
        rs.Close
        grdTotal.TextMatrix(0, 1) = Format(nTotQty, "##,###,##0 YD")
        
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
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oPlanCPB = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmPlanCPBView.FillGridData", Err.Description)
End Sub


