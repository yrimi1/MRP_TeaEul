VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStockReport 
   Caption         =   "재고 명세서"
   ClientHeight    =   9270
   ClientLeft      =   2730
   ClientTop       =   3585
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.ComboBox cboTaxClss 
      Height          =   300
      Left            =   6570
      Style           =   2  '드롭다운 목록
      TabIndex        =   20
      Top             =   360
      Width           =   1395
   End
   Begin Threed.SSPanel pnlPrint 
      Height          =   3195
      Left            =   2850
      TabIndex        =   11
      Top             =   2580
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5636
      _Version        =   196609
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame Frame1 
         Caption         =   "인쇄구분"
         Height          =   1065
         Left            =   990
         TabIndex        =   17
         Top             =   510
         Width           =   3555
         Begin VB.OptionButton opPRN 
            Caption         =   "업체별 인쇄"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   19
            Top             =   720
            Width           =   1995
         End
         Begin VB.OptionButton opPRN 
            Caption         =   "현황인쇄"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   18
            Top             =   330
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   405
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   714
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "재고 명세서 인쇄"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.ComboBox cboCustom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2220
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   1800
         Width           =   2265
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   2970
         TabIndex        =   14
         Top             =   2430
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   1020
         TabIndex        =   15
         Top             =   2430
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "인쇄"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   930
         TabIndex        =   16
         Top             =   1800
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄범위"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   6570
      TabIndex        =   2
      Top             =   30
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   8430
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   3
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   30
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월 dd일"
      Format          =   116785155
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   5250
      TabIndex        =   4
      Top             =   360
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "사용구분"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   5
      Top             =   8520
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7770
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   690
      Width           =   11790
      _cx             =   20796
      _cy             =   13705
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8490
      TabIndex        =   7
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   0
      Left            =   8040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   9
      Left            =   5250
      TabIndex        =   9
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   975
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   840
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   4350
      _cx             =   7673
      _cy             =   1482
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   11.25
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
      Index           =   0
      Left            =   30
      TabIndex        =   21
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "재고일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   3180
      TabIndex        =   22
      Top             =   30
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월 dd일"
      Format          =   116785155
      CurrentDate     =   36871
   End
End
Attribute VB_Name = "frmStockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSearch_Click(Index As Integer)
    Select Case Index

        Case 1    '거래처
            If chkSearch(Index) = vbChecked Then
                txtCustom(1).Enabled = True
                txtCustom(1).SetFocus
                cmdFind(0).Enabled = True
            Else
                txtCustom(1).Enabled = False
                cmdFind(0).Enabled = False
                txtCustom(1).Tag = ""
            End If
            
    End Select
End Sub

Private Sub chkSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

'Private Sub chkSearch_Click()
'    If chkSearch.Value = vbChecked Then
'        dtpDate(0).Enabled = True
'        dtpDate(1).Enabled = True
'    Else
'        dtpDate(0).Enabled = False
'        dtpDate(1).Enabled = False
'    End If
'End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0                '[1] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustom(1))
    End Select
End Sub

Private Sub cmdPrint_Click()

    pnlPrint.Visible = True

End Sub

Private Sub cmdPrnCancel_Click()
    pnlPrint.Visible = False
End Sub

Private Sub cmdPrnOK_Click()
    Dim II As Integer
    If opPrn(0).Value = True Then
        Call FillGrdPrint
    Else
        If cboCustom.Text = AllStr Then
            For II = 1 To cboCustom.ListCount - 1
                Call SetDataToPrn(cboCustom.List(II))
            Next II
        Else
            Call SetDataToPrn(cboCustom.Text)
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub

Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String
    Dim nRowHeight As Integer
    Dim nBackColor As Long
    Dim nPageHV As Integer

    
    With grdData(0)
        .Redraw = flexRDBuffered
        .ExtendLastCol = True
        
        Call SetPrintMode(grdData(0), 1, True)

        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "재고명세서"
        .RowHeight(0) = 1000
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "▶ 재고기간 : " & MakeDate(DF_FULL, dtpDate(0)) & " ~ " & MakeDate(DF_FULL, dtpDate(1))
        .Cell(flexcpText, 2, 1, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpBackColor, 0, 0, .FixedRows - 2, .Cols - 1) = vbWhite
        .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignRightCenter
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
'        For i = .FixedRows To .Rows - 1
'            .RowHeight(i) = 400
'            ' 일계, 총계의 금액은 BackColor을 설정 한다.
'            If (.TextMatrix(i, 11) = "Z4" Or .TextMatrix(i, 11) = "Z5") And .ValueMatrix(i, 10) <> 0 Then
'                .Cell(flexcpBackColor, i, 6, i, .Cols - 1) = PRNHeaderColor
'            End If
'        Next i
        
        .PrintGrid "태을염직", True, 1, 100, 500
        
 '----  인쇄하기 이전으로 원상복귀
        Call SetPrintMode(grdData(0), 1, False)

        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True

        .ExtendLastCol = True
        
'        For i = .FixedRows To .Rows - 1
'             Call SetGrdColor(grdData, Mid(.TextMatrix(i, 11), 2), i, 0, i, .Cols - 1)
'        Next i
        .Redraw = flexRDDirect
        
    End With
    
    
    
End Sub

Sub FillGrdPrintHeader(ByVal kCustom As String)
    Dim i%
    Dim sDate As String
    
    sDate = Format(dtpDate(0), "YYYY/MM/DD")
    
    With grdData(1)
        .Redraw = flexRDBuffered
        .Rows = .FixedRows
        
        .ExtendLastCol = True
        
   '     Call SetPrintMode(grdData(1), 1, True)
        
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .RowHeight(0) = 1000
        .RowHeight(1) = 350
        .RowHeight(2) = 350
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "원단 수불내역"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "▶ 거 래 처 : " & kCustom
        .Cell(flexcpText, 2, 1, 2, .Cols - 1) = "▶ 재고일자 : " & MakeDate(DF_FULL, dtpDate(0)) & " ~ " & MakeDate(DF_FULL, dtpDate(1))
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter

        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    
End Sub

Sub SetDataToPrn(ByVal kCustom As String)
    Dim II%, JJ%
    Call FillGrdPrintHeader(kCustom)
    With grdData(1)
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            If grdData(0).TextMatrix(II, 1) = kCustom And grdData(0).TextMatrix(II, 8) = "Z0" Then
                .AddItem ""
                For JJ = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
                .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
                .Redraw = flexRDDirect
            End If
        Next II
        
        .MergeCells = flexMergeFree
        For II = 1 To 2
            .MergeCol(II) = True
        Next II
        .ColHidden(0) = True
        .ColHidden(1) = True
        
        .ColWidth(2) = 3580  '품명
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        .ColWidth(5) = 1800
        .ColWidth(6) = 1800
        .ColWidth(7) = 1800
        
        .ExtendLastCol = False

        Call SetPrintMode(grdData(1), 1, True)
        .PrintGrid "태을염직", True, 2, 700, 500
        Call SetPrintMode(grdData(1), 1, False)
    End With
End Sub

'Sub FillGrdPrint()
'    Dim i%
'    Dim sDate As String, eDate As String
'    Dim JJ%, nRow%
'
'    If chkSearch(0).Value Then
'        sDate = Format(dtpDate(0), "YYYY/MM/DD")
'        eDate = Format(dtpDate(1), "YYYY/MM/DD")
'    Else
'        sDate = ""
'        eDate = ""
'    End If
'    '----------------
'    ' 인쇄시 관리번호 제외하고 인쇄
'    ' chkReport.value = vbchecked 일때 ( 결재용 ) : Design 부분 제외한 후 인쇄
'    '-------------------
'    With grdData
'        .Redraw = flexRDBuffered
'        .ExtendLastCol = False
'
'        .GridLinesFixed = flexGridNone
'        .GridLines = flexGridFlat
'        .RowHidden(0) = False
'        .RowHidden(1) = False
'        .RowHeight(0) = 500
'        .RowHeight(1) = 350
'
''        For i = 0 To 3
''           .MergeRow(i) = True
''        Next i
'
'
'        .FontSize = 7
'        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "생지입고명세서"
'        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
'        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
'        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
'
'        .Cell(flexcpText, 1, 1, 1, 3) = "▶ 입고일자 : " & sDate & " ~ " & eDate
'        .Cell(flexcpText, 1, .Cols - 4, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
'
'        .Cell(flexcpText, 1, 4, 1, 4) = "▶ 입고구분 : " & IIf(chkSearch(4).Value, CboStuffClss2.Text, "(전체)")
'        .Cell(flexcpText, 1, 5, 1, 7) = "▶ 확정구분 : " & IIf(chkSearch(5).Value, cboOrderID.Text, "(전체)")
'        .Cell(flexcpText, 2, 7, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
'        .Cell(flexcpBackColor, 0, 0, .FixedRows - 2, .Cols - 1) = vbWhite
'        .Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1, .Cols - 1) = vbWhite
'
'        For i = 0 To 3
'           .MergeRow(i) = True
'        Next i
'
'        .ColHidden(5) = True
'
'        '-- 결재용인 경우 design 부분 제외 처리
'        If chkReport.Value = vbChecked Then
'            For i = .FixedRows To .Rows - 1
'                If .TextMatrix(i, 12) = "DD" Then
'                    nRow = i
'                    Exit For
'                End If
'            Next i
'            For JJ = nRow To .Rows - 1
'                .RowHidden(JJ) = True
'            Next JJ
'        End If
'
'        .ColHidden(0) = True
'
'        .PrintGrid "인쇄명: 생지입고명세서", False, 1, 400, 500
'
'        ' 인쇄완료 후 조회 모들 되돌려 놓기
'        .GridLinesFixed = flexGridInset
'        .RowHidden(0) = True
'        .RowHidden(1) = True
'        .RowHidden(2) = True
'
'
'        For i = .FixedRows To .Rows - 1
'            Call SetGrdColor(grdData, Mid(.TextMatrix(i, 12), 2), i, 1, i, .Cols - 1)
'        Next i
'
'                '-- 결재용인 경우 design 부분 제외 처리
'        If chkReport.Value = ssCBChecked Then
'            For JJ = nRow + 1 To .Rows - 1
'                .RowHidden(JJ) = False
'            Next JJ
'        End If
'
'        .ColHidden(5) = False
'        .ColHidden(0) = False
'
'
'        .FontSize = 9
'
'        .ExtendLastCol = True
'        .Redraw = flexRDDirect
'    End With
'
'End Sub


Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11970, 9660

    Call InitGrid(0)
    Call InitGrid(1)
    
    Call SetOperate(Me)
    
    '----- 날짜설정
    dtpDate(0) = Now
    dtpDate(1) = Now
    
'    CboStuffClss2.ListIndex = 0
    

    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    

    cmdFind(0).Enabled = False
    
    txtCustom(1).Enabled = False
    pnlPrint.Visible = False
    
    With cboTaxClss
        .AddItem "9.전체"
        .AddItem "0.비사용"
        .AddItem "1.사용"
        .ListIndex = 0
    End With
    

End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim II%, nRows As Integer
    
    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Cols = 9
        .Rows = 4
        .FixedRows = 4
        .FixedCols = 1
        
        .RowHeightMin = 300
        
        
        For II = 0 To nRows - 1
            .RowHidden(II) = True
        
        Next II
        
        
        nRows = 3
        
        .RowHeight(nRows) = 400
        
        .TextMatrix(nRows, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "거래처":           .ColWidth(1) = 2000:                .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(nRows, 2) = "품명":             .ColWidth(2) = 2600:                .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(nRows, 3) = "전월이월":         .ColWidth(3) = 1300:                .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(nRows, 4) = "당기입고":         .ColWidth(4) = 1300:                .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(nRows, 5) = "가공납품":         .ColWidth(5) = 1300:                .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(nRows, 6) = "소요량":           .ColWidth(6) = 1300:                .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(nRows, 7) = "차기이월":         .ColWidth(7) = 1300:                .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(nRows, 8) = "Depth":            .ColWidth(8) = 0:                   .ColAlignment(8) = flexAlignCenterCenter
        
        .MergeCells = flexMergeFree
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .Redraw = flexRDDirect
    End With

End Sub

Sub FillgrdData()
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    
    Set rs = oStuffIn.GetStockReport(MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)) _
                                   , IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag, Left(cboTaxClss, 1))

    Set oStuffIn = Nothing
    cboCustom.Clear
    cboCustom.AddItem AllStr
    With grdData(0)
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        If rs.RecordCount < 1 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
            Do Until rs.EOF
                
                .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & IIf(rs!Depth = "Z0", " " & Trim(rs!StuffWidth) & "“", "") & vbTab & _
                SetCurrency(rs!StockQty) & vbTab & SetCurrency(rs!StuffQty) & vbTab & SetCurrency(rs!OutQty) & vbTab & SetCurrency(rs!OutRealQty) & vbTab & SetCurrency(rs!StockQtyNOW) & vbTab & rs!Depth
                If rs!StockQty = 0 And rs!StuffQty = 0 And rs!OutQty = 0 And rs!OutRealQty = 0 And rs!StockQtyNOW = 0 Then
                    .RemoveItem (.Rows - 1)
                End If
                
                Select Case rs!Depth
                    Case "Z2":   .TextMatrix(.Rows - 1, 1) = "합계"
                                 .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                    Case "Z1":   cboCustom.AddItem Trim(rs!kCustom)
                                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                        If rs!nCount = 1 Then
                            .RowHidden(.Rows - 1) = True
                        End If
                        .AddItem ""
                        .RowHidden(.Rows - 1) = True
                End Select
                rs.MoveNext
            Loop
        End If
        
        .MergeCells = flexMergeFree
        For i = 1 To 1
            .MergeCol(i) = True
        Next i
        
        .Redraw = flexRDDirect
    End With
    cboCustom.ListIndex = 0
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "FrmDeliveryReport.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub



Private Sub opPrn_Click(Index As Integer)
    Select Case Index
        Case 0: cboCustom.Enabled = False
        Case 1: cboCustom.Enabled = True
    End Select
    
End Sub

Private Sub txtCustom_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 1
            Call MoveFocus(KeyCode)
    End Select

End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call cmdFind_Click(0)
            End If
    End Select
End Sub

