VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanInputView 
   ClientHeight    =   9255
   ClientLeft      =   240
   ClientTop       =   1365
   ClientWidth     =   11850
   Icon            =   "frmPlanInputView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin VSFlex7LCtl.VSFlexGrid grdSum 
      Height          =   315
      Left            =   0
      TabIndex        =   30
      Top             =   8070
      Width           =   11835
      _cx             =   20876
      _cy             =   556
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
      Height          =   7095
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   11835
      _cx             =   20876
      _cy             =   12515
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
         Left            =   390
         TabIndex        =   23
         Top             =   2820
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
            TabIndex        =   24
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
            TabIndex        =   25
            Top             =   120
            Width           =   270
         End
      End
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1667
      _Version        =   196609
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   465
         Width           =   600
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
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   4
         Top             =   435
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   7170
         TabIndex        =   3
         Top             =   435
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   9120
         TabIndex        =   2
         Top             =   435
         Width           =   1485
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10950
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3450
         TabIndex        =   7
         Top             =   75
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
         Left            =   3450
         TabIndex        =   8
         Top             =   465
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
         Left            =   2130
         TabIndex        =   9
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
            Caption         =   "지시 일자"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   11
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   12
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   6690
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   435
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
         Left            =   7140
         TabIndex        =   14
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
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
            TabIndex        =   15
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   8700
         TabIndex        =   16
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
         Left            =   9120
         TabIndex        =   17
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
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
            TabIndex        =   18
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   735
         Left            =   90
         TabIndex        =   26
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1296
         _Version        =   196609
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   27
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
         Left            =   4755
         TabIndex        =   20
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4755
         TabIndex        =   19
         Top             =   135
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   22
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   690
      Left            =   8445
      TabIndex        =   29
      Tag             =   "PERM_DELETE"
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      삭제(&D)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   6750
      TabIndex        =   31
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmPlanInputView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdDelete_Click()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim sInstDate$, nInstSeq%
    
    On Error GoTo ErrHandler
    
    If MsgBox(LoadResString(201), vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    With grdData
        sInstDate = .TextMatrix(.Row, 14)
        nInstSeq = .TextMatrix(.Row, 2)
    End With
    If oPlanInput.DeletePlanInput(sInstDate, nInstSeq) Then
        Call FillGridData
    Else
        MsgBox "자료 삭제에 실패하였습니다", vbInformation + vbOKOnly
    End If
    Set oPlanInput = Nothing
    Exit Sub
    
ErrHandler:
    Set oPlanInput = Nothing
    Call ErrorBox(Err.Number, "frmPlanInputview.cmdDelete_Click", Err.Description)
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
    With grdData
        .Redraw = flexRDBuffered
    
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = vbWhite
'        .GridLinesFixed = flexGridNone
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(.Rows - 1) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "생지 투입 계획"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 1, 1, 1, 6) = "▶ 지시일 : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD")
        .Cell(flexcpAlignment, 1, 1, 1, 6) = flexAlignLeftCenter
        .Cell(flexcpText, 1, 9, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpAlignment, 1, 9, 1, .Cols - 1) = flexAlignRightCenter
        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        .ColWidth(3) = 1500
        .ColWidth(5) = 1700
        .ColWidth(6) = 2800
        .ColWidth(7) = 1400
        .ColWidth(9) = 1100
        .ColWidth(10) = 1100
        .ColWidth(11) = 1100
        .ColWidth(12) = 1100
        
        .PrintGrid "태을염직", True, 2, 100, 500

        .GridLinesFixed = flexGridInset
        .GridColorFixed = &H80000010
'        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(.Rows - 1) = True

        .ColWidth(3) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 2000
        .ColWidth(7) = 1400
        .ColWidth(9) = 800
        .ColWidth(10) = 800
        .ColWidth(11) = 800
        .ColWidth(12) = 800
        
        
        
        
'        .TextMatrix(2, 1) = "지시일":       .ColWidth(1) = 700:         .ColAlignment(1) = flexAlignCenterCenter
'        .TextMatrix(2, 2) = "지시순위":     .ColWidth(2) = 0
'        .TextMatrix(2, 3) = "관리번호":     .ColWidth(3) = 1300:        .ColAlignment(3) = flexAlignCenterCenter
'        .TextMatrix(2, 4) = "Order No.":    .ColWidth(4) = 0:           .ColAlignment(4) = flexAlignLeftCenter
'        .TextMatrix(2, 5) = "거래처":       .ColWidth(5) = 1300:        .ColAlignment(5) = flexAlignLeftCenter
'        .TextMatrix(2, 6) = "품명":         .ColWidth(6) = 2000:        .ColAlignment(6) = flexAlignLeftCenter
'        .TextMatrix(2, 7) = "색상명":       .ColWidth(7) = 1400:        .ColAlignment(7) = flexAlignLeftCenter
'        .TextMatrix(2, 8) = "축율":         .ColWidth(8) = 650:         .ColAlignment(8) = flexAlignCenterCenter
'        .TextMatrix(2, 9) = "수주량":       .ColWidth(9) = 800:        .ColAlignment(9) = flexAlignRightCenter
'        .TextMatrix(2, 10) = "입고량":      .ColWidth(10) = 800:       .ColAlignment(10) = flexAlignRightCenter
'        .TextMatrix(2, 11) = "미계획량":    .ColWidth(11) = 800:       .ColAlignment(11) = flexAlignRightCenter
'        .TextMatrix(2, 12) = "계획량":      .ColWidth(12) = 800:       .ColAlignment(12) = flexAlignRightCenter
'        .TextMatrix(2, 13) = "배색량":      .ColWidth(13) = 800:       .ColAlignment(13) = flexAlignRightCenter
        
        .Redraw = flexRDDirect
    End With
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

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
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
    cmdDelete.Picture = LoadResPicture("DELETE", vbResIcon)
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
        If optOrder(0).Value Then
            .ColWidth(4) = 1350
            .ColWidth(3) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(4) = 0
            .ColWidth(3) = 1350
            chkSearch(3).Caption = "관리번호"
        End If
    End With
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
        .Cols = 15
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 3
        .FixedRows = 3
        

        .RowHidden(0) = True
        .RowHidden(1) = True

        .TextMatrix(2, 0) = " "
        .TextMatrix(2, 1) = "지시일":       .ColWidth(1) = 700:         .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(2, 2) = "지시순위":     .ColWidth(2) = 0
        .TextMatrix(2, 3) = "관리번호":     .ColWidth(3) = 1300:        .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(2, 4) = "Order No.":    .ColWidth(4) = 0:           .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(2, 5) = "거래처":       .ColWidth(5) = 1100:        .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(2, 6) = "품명":         .ColWidth(6) = 2000:        .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(2, 7) = "색상명":       .ColWidth(7) = 1400:        .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(2, 8) = "축율":         .ColWidth(8) = 550:         .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(2, 9) = "수주량":       .ColWidth(9) = 1200:        .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(2, 10) = "입고량":      .ColWidth(10) = 800:       .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(2, 11) = "미계획량":    .ColWidth(11) = 800:       .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(2, 12) = "계획량":      .ColWidth(12) = 800:       .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(2, 13) = "배색량":      .ColWidth(13) = 800:       .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(2, 14) = "지시일자":     .ColWidth(14) = 0:       .ColAlignment(14) = flexAlignRightCenter
        
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        .ColFormat(11) = "#,##0"
        .ColFormat(12) = "#,##0"
        .ColFormat(13) = "#,##0"
        
        .Redraw = flexRDDirect
    End With
    
    With grdSum
        .Redraw = flexRDNone
        
        .Cols = 3
        Call SetVSFlexGrid(grdSum)
        
        .RowHeight(0) = 300
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 0
        
        .TextArray(0) = "합계":     .ColWidth(0) = 9710:    .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "0":        .ColWidth(1) = 1000:    .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "0":        .ColWidth(2) = 1000:    .ColAlignment(2) = flexAlignRightCenter
        
        .ColFormat(1) = "#,##0"
        .ColFormat(2) = "#,##0"
        .HighLight = flexHighlightNever
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub FillGridData()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim rs As ADODB.Recordset
    Dim i%, nNoPlanQty As Double, nTotQty(2) As Double
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    Set rs = oPlanInput.GetPlanInput(IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                 IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3))
    Set oPlanInput = Nothing
        
    nTotQty(0) = 0
    nTotQty(1) = 0
    With grdData
        .Redraw = flexRDDirect
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
             nNoPlanQty = IIf(rs!UnitClss = "0", rs!OrderQty, CLng(rs!OrderQty / 0.9144)) * (1 + rs!ChunkRate / 100) - rs!InstQty  '미계획량
             
            .AddItem CStr(i + 1) & vbTab & MakeDate(DF_MD, rs!InstDate) & vbTab & rs!InstSeq & vbTab & _
                    MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    rs!kCustom & vbTab & rs!Article & vbTab & rs!Color & vbTab & _
                    rs!ChunkRate & vbTab & Format(rs!OrderQty, "#,###") & IIf(rs!UnitClss = "0", "", "M") & vbTab & rs!StuffInQty & vbTab & nNoPlanQty & vbTab & _
                    rs!InstQty & vbTab & rs!matchqty & vbTab & rs!InstDate
            
            nTotQty(0) = nTotQty(0) + rs!InstQty
            nTotQty(1) = nTotQty(1) + rs!matchqty
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        
        .Rows = .Rows + 1
        .RowHidden(.Rows - 1) = True
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 11) = "합계"
        .Cell(flexcpText, .Rows - 1, 12) = nTotQty(0)
        .Cell(flexcpText, .Rows - 1, 13) = nTotQty(1)
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
        .MergeCells = flexMergeFree
        .MergeRow(.Rows - 1) = True
        .MergeCol(1) = True
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            grdSum.TextMatrix(0, 1) = 0
            grdSum.TextMatrix(0, 2) = 0
            MsgBox LoadResString(203), vbInformation
            
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    
    grdSum.TextMatrix(0, 1) = nTotQty(0)
    grdSum.TextMatrix(0, 2) = nTotQty(1)
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oPlanInput = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmPlanInputView.FillGridData", Err.Description)
End Sub


