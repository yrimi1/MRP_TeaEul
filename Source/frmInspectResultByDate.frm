VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectResultByDate 
   ClientHeight    =   9255
   ClientLeft      =   1755
   ClientTop       =   3600
   ClientWidth     =   11865
   Icon            =   "frmInspectResultByDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.ComboBox cboExamNo 
      Height          =   300
      Left            =   10095
      Style           =   2  '드롭다운 목록
      TabIndex        =   26
      Top             =   450
      Width           =   870
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7605
      Left            =   0
      TabIndex        =   25
      Top             =   840
      Width           =   11865
      _cx             =   20929
      _cy             =   13414
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
      Left            =   45
      TabIndex        =   22
      Top             =   -15
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   495
         Value           =   -1  'True
         Width           =   1110
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   7095
      TabIndex        =   15
      Top             =   105
      Width           =   1500
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   300
      Index           =   2
      Left            =   10095
      MaxLength       =   4
      TabIndex        =   14
      Top             =   105
      Width           =   870
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2070
      MousePointer    =   99  '사용자 정의
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1425
      MousePointer    =   99  '사용자 정의
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2070
      MousePointer    =   99  '사용자 정의
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1425
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   780
      Left            =   11055
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "자료 검색"
      Top             =   30
      Width           =   780
   End
   Begin Threed.SSPanel pnlLanguage 
      Height          =   690
      Left            =   15
      TabIndex        =   1
      Top             =   8520
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1217
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optPrint 
         Caption         =   "영문"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   690
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "한글"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   3930
      TabIndex        =   8
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   21430273
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   3930
      TabIndex        =   9
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   21430273
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2715
      TabIndex        =   10
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "검사일자"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Index           =   0
      Left            =   8445
      TabIndex        =   12
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
      TabIndex        =   13
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   8805
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   3
      Left            =   5880
      TabIndex        =   16
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "관리번호"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   4
      Left            =   8820
      TabIndex        =   18
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "LOT No."
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Top             =   45
         Width           =   1095
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   8820
      TabIndex        =   27
      Top             =   450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "검사호기"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   28
         Top             =   45
         Width           =   1020
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Index           =   1
      Left            =   6300
      TabIndex        =   29
      Top             =   8520
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      위치별결과표(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Frame fraRange 
      Height          =   780
      Left            =   900
      TabIndex        =   30
      Top             =   8430
      Visible         =   0   'False
      Width           =   1710
      Begin VB.OptionButton optRange 
         Caption         =   "200"
         Height          =   210
         Index           =   1
         Left            =   975
         TabIndex        =   32
         Top             =   480
         Width           =   690
      End
      Begin VB.OptionButton optRange 
         Caption         =   "100"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.Label lblName 
         Caption         =   "■ 범위 설정"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   33
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   3
      Left            =   5205
      TabIndex        =   21
      Top             =   165
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   2
      Left            =   5205
      TabIndex        =   20
      Top             =   510
      Width           =   360
   End
End
Attribute VB_Name = "frmInspectResultByDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE_1 = "\Report\InspectResultByDate_K.rpt"
Private Const REPORTFILE_2 = "\Report\InspectResultByDate_E.rpt"

Private Const BASE_X       As Integer = 150
Private Const BASE_Y       As Integer = 1300
Private Const DEFECT_COUNT As Integer = 50

Private Type TDefect
    Korean  As String
    English As String
    Defect  As String
End Type

Private m_iSortType As Integer
Private m_nSelected As Integer
Dim m_sTotalField(7)  As String             ' 리포트 Title
Dim m_nDefectName(DEFECT_COUNT) As TDefect
Dim m_nPageCnt(1) As Integer


Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)
    
    cmdPrint(0).Picture = LoadResPicture("PRINT", vbResIcon)
    cmdPrint(1).Picture = LoadResPicture("PRINT", vbResIcon)
    
    With cboExamNo
        For i = 1 To 10
            .AddItem Format(i, "00") & "호기"
        Next i
        .ListIndex = 0
    End With
    dtpDate(0) = Now
    dtpDate(1) = Now

    Call InitGrid

    Show

    chkSearch(0).Value = vbChecked
    m_nSelected = 0
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If chkSearch(Index) Then
        If Index = 0 Then
            dtpDate(0).SetFocus
        ElseIf Index = 2 Then
            txtSearch(2).SetFocus
        ElseIf Index = 3 Then
            cboExamNo.SetFocus
        End If
    Else
        cmdSearch.SetFocus
    End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub cmdSearch_Click()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%, iNowRow%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetResultByDate(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1), IIf(optOrder(0), 2, 1), 0), IIf(optOrder(0), txtSearch(1), MakeOrderID(txtSearch(1), OM_REDUCE)), _
        IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    Set oInspect = Nothing

    With grdData
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!kCustom & vbTab & rs!Article & vbTab & MakeDate(DF_LONG, rs!DvlyDate) & vbTab & _
                rs!WorkName & vbTab & rs!Color & vbTab & CheckNum(rs!ColorQty) & vbTab & _
                CheckNum(rs!PassRoll) & vbTab & CheckNum(rs!PassQty) & vbTab & CheckNull(rs!ECustom) & vbTab & _
                rs!UnitClss & vbTab & rs!WorkWidth & vbTab & rs!OrderSeq & vbTab & CheckNull(rs!DesignNO) & vbTab & _
                CheckNum(rs!LossQty) & vbTab & CheckNum(rs!SampleQty) & vbTab & CheckNum(rs!CutQty)

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
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        m_nSelected = 0

        .Redraw = flexRDDirect

        .SetFocus
    End With

    Screen.MousePointer = vbArrow

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub grdData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdData
        If .Rows = .FixedRows Or .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub

        If .Cell(flexcpChecked, .Row, 1) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, 1) = flexChecked
            m_nSelected = m_nSelected + 1
        Else
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
            m_nSelected = m_nSelected - 1
        End If
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkSearch(1).Caption = "Order No"
    Else
        chkSearch(1).Caption = "관리번호"
    End If

    cmdSearch.SetFocus
End Sub

Private Sub optPrint_Click(Index As Integer)
    cmdPrint(0).SetFocus
End Sub

Private Sub cmdPrint_Click(Index As Integer)
    Dim i%

    If grdData.Rows = grdData.FixedRows Then Exit Sub


    If m_nSelected = 1 Then
        With grdData
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then Exit For
            Next i
            If Index = 0 Then
                PopupMenu PlusMDI.mnuPopup
                Call ReportPrint(PlusMDI.PrintPreview, MakeOrderID(.TextMatrix(i, 3), OM_REDUCE), .TextMatrix(i, 15))
            Else
                Call PrintDefectPosition(MakeOrderID(.TextMatrix(i, 3), OM_REDUCE), .TextMatrix(i, 15))
            End If
        End With
    ElseIf m_nSelected > 1 Then
        With grdData
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then
                    .Row = i
                    If Index = 0 Then
                        Call ReportPrint(False, MakeOrderID(.TextMatrix(i, 3), OM_REDUCE), .TextMatrix(i, 15))
                    Else
                        Call PrintDefectPosition(MakeOrderID(.TextMatrix(i, 3), OM_REDUCE), .TextMatrix(i, 15))
                    End If
                End If
            Next i
        End With
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 20
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .TextArray(1) = "선택":         .ColWidth(1) = 300:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order No.":    .ColWidth(2) = 1350:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호":     .ColWidth(3) = 1225:            .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "거래처":       .ColWidth(4) = 1170:            .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "품명":         .ColWidth(5) = 2000:     .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "납기일자":     .ColWidth(6) = 990:             .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "가공구분":     .ColWidth(7) = 450:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":       .ColWidth(8) = 1530:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "수주수량":     .ColWidth(9) = 830:             .ColAlignment(9) = flexAlignRightCenter:    .ColFormat(9) = GetFormat()
        .TextArray(10) = "검사절수":    .ColWidth(10) = 450:            .ColAlignment(10) = flexAlignRightCenter:   .ColFormat(10) = GetFormat()
        .TextArray(11) = "검사수량":    .ColWidth(11) = 830:            .ColAlignment(11) = flexAlignRightCenter:   .ColFormat(11) = GetFormat(g_nPointPos)
        .TextArray(12) = "거래처(영)":  .ColWidth(12) = 0
        .TextArray(13) = "수량단위":    .ColWidth(13) = 0
        .TextArray(14) = "생지폭":      .ColWidth(14) = 0
        .TextArray(15) = "색상순위":    .ColWidth(15) = 0
        .TextArray(16) = "DesignNo":    .ColWidth(16) = 0
        .TextArray(17) = "보상수량":    .ColWidth(17) = 0:   .ColFormat(17) = GetFormat(1)
        .TextArray(18) = "견본수량":    .ColWidth(18) = 0:   .ColFormat(18) = GetFormat(g_nPointPos)
        .TextArray(19) = "난단수량":    .ColWidth(19) = 0:   .ColFormat(19) = GetFormat(g_nPointPos)

        .ColDataType(1) = flexDTBoolean

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub ReportPrint(bDirect As Boolean, sOrderID As String, nOrderSeq As Integer)
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim sParam() As String
    Dim i%, iPoint%, iLoop%

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    ' footer
    Set rs = oInspect.GetDefectByLang(IIf(optPrint(0), 1, 2))
    Set oInspect = Nothing

    ReDim sParam(39)

    For i = 0 To rs.RecordCount - 1
        sParam(i) = CStr(rs!Tag) & "-" & CStr(CheckNull(rs!Display))

        rs.MoveNext
    Next i
    rs.Close
    Set rs = Nothing

    Do While i <= 39
        sParam(i) = " "
        i = i + 1
    Loop

    ReDim Preserve sParam(53)

    ' MarkClss
    If optPrint(0) Then
        If dtpDate(0) = dtpDate(1) Then
            sParam(i) = "검사일자 : " & Format(dtpDate(0), "YYYY년 MM월 DD일") & Space(5) & _
                IIf(chkSearch(3), "호기 : " & cboExamNo, "전체 호기")
        Else
            sParam(i) = "검사일자 : " & Format(dtpDate(0), "YYYY년 MM월 DD일") & " ~ " & Format(dtpDate(1), "YYYY년 MM월 DD일") & Space(5) & _
                IIf(chkSearch(3), "호기 : " & cboExamNo, "전체 호기")
        End If
    Else
        If dtpDate(0) = dtpDate(1) Then
            sParam(i) = "INSPECTION DATE : " & Format(dtpDate(0), "YYYY/MM/DD") & Space(5) & _
                IIf(chkSearch(3), "Machine No : " & Left(cboExamNo, 2) & "Mc", "Total Machine")
        Else
            sParam(i) = "INSPECTION DATE : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD") & Space(5) & _
                IIf(chkSearch(3), "Machine No : " & Left(cboExamNo, 2) & "Mc", "Total Machine")
        End If
    End If

    i = i + 1

    ' Header
    iPoint = grdData.Cols * grdData.Row
    sParam(i) = MakeOrderID(sOrderID, OM_EXPAND)
    sParam(i + 1) = grdData.TextArray(iPoint + IIf(optPrint(0), 4, 12))
    sParam(i + 2) = grdData.TextArray(iPoint + 8)
    sParam(i + 3) = IIf(grdData.TextArray(iPoint + 13) = "0", "YARD", "METER")
    sParam(i + 4) = grdData.TextArray(iPoint + 2)
    sParam(i + 5) = grdData.TextArray(iPoint + 5)
    sParam(i + 6) = grdData.TextArray(iPoint + 7)
    sParam(i + 7) = grdData.TextArray(iPoint + 14)
    sParam(i + 8) = Format(grdData.TextArray(iPoint + 11), "#,##0") & IIf(grdData.TextArray(iPoint + 13) = "0", "Y", "M")
    sParam(i + 9) = Format(grdData.TextArray(iPoint + 17), "#,##0")
    sParam(i + 10) = grdData.TextArray(iPoint + 10)
    sParam(i + 11) = Format(grdData.TextArray(iPoint + 18), "#,##0")
    sParam(i + 12) = Format(grdData.TextArray(iPoint + 19), "#,##0")
    
    i = i + 13
    
    ReDim Preserve sParam(61)
    ' title
    If optPrint(0) Then
        sParam(i) = "검사 결과표"
    Else
        sParam(i) = "INSPECTION REPORT"
    End If
    ' companyname
    sParam(i + 1) = CompanyName

    i = i + 2

    ' GradeCount
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetGradeQtyByColor(sOrderID, nOrderSeq, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), _
        MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(2), 1, 0), txtSearch(2), _
        IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))

    For iLoop = 0 To rs.RecordCount - 1
        sParam(i + iLoop) = IIf(IsNull(rs!GradeCount), "0", rs!GradeCount)
        
        rs.MoveNext
    Next iLoop
    rs.Close
    Set rs = Nothing
    If optPrint(0) Then
    Set rs = oInspect.PrintResultByColor(sOrderID, nOrderSeq, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), _
        MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(2), 1, 0), txtSearch(2), _
        IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    Else: Set rs = oInspect.PrintResultByColor_eng(sOrderID, nOrderSeq, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)) _
        , MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(2), 1, 0), txtSearch(2) _
        , IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    End If
    Set oInspect = Nothing

    Call PrintReport(IIf(optPrint(0), REPORTFILE_1, REPORTFILE_2), rs, sParam, bDirect)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub PrintDefectPosition(sOrderID As String, sColorID As String)
    Dim oOrder   As PlusLib2.COrder
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim rsData   As ADODB.Recordset
    Dim sGrade$, sTemp As String * 6
    Dim nRoll%, nPenalty%
    Dim i%, j%, k%
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
       
    Call GetDefect

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    sOrderID = MakeOrderID(grdData.TextMatrix(grdData.Row, 3), OM_REDUCE)
    Set rs = oOrder.GetOrderOne(sOrderID)
    Set oOrder = Nothing
    
    m_sTotalField(0) = grdData.TextMatrix(grdData.Row, 3)           ' 관리번호
    If optPrint(0).Value = True Then
       m_sTotalField(1) = grdData.TextMatrix(grdData.Row, 4)        ' 한글 거래처
    Else
       m_sTotalField(1) = grdData.TextMatrix(grdData.Row, 12)        ' 영문 거래처
    End If
    m_sTotalField(2) = grdData.TextMatrix(grdData.Row, 8)           ' 색상명
    m_sTotalField(3) = grdData.TextMatrix(grdData.Row, 2)           ' Order No
    m_sTotalField(4) = rs!Article
    m_sTotalField(5) = IIf(rs!UnitClss = "0", "Yards", "Meters")
    m_sTotalField(6) = rs!WorkWidth

    rs.Close
    Set rs = Nothing

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetInspectRollNo(sOrderID, sColorID, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), _
        MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(2), 1, 0), txtSearch(2), _
        IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    If rs.EOF Then
        Call MessageBox("해당 검사실적이 없습니다.")
        Printer.KillDoc
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    m_nPageCnt(0) = 1
    m_nPageCnt(1) = Int(rs.RecordCount / 10) + 1
    ' Report 양식 출력
    Call ReportForm(IIf(optRange(0).Value = True, 1, 2), IIf(optPrint(0).Value = True, 1, 2))

    nRoll = 0

    With rs
        Do Until .EOF
            ' NewPage
            If nRoll = 10 Then
                nRoll = 0
                m_nPageCnt(0) = m_nPageCnt(0) + 1
                Printer.NewPage
                Call ReportForm(IIf(optRange(0).Value = True, 1, 2), IIf(optPrint(0).Value = True, 1, 2))
            End If

            Printer.Font.Size = 7

            ' Roll No
            Printer.CurrentX = BASE_X + 1300 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 700
            Printer.Print !RollNo
            
            '투입량
            Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11000
            RSet sTemp = CStr(!StuffQty)
            Printer.Print sTemp
            '실폭
            Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11000
            If Not IsNull(!StuffWidth) Then
                RSet sTemp = Format(val(!StuffWidth), "#0.0")
            Else
                RSet sTemp = Format(0, "#0.0")
            End If
            Printer.Print sTemp
            
            '계측량(실제검사 수량)
            Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11300
            RSet sTemp = str(IIf(!RealQty = 0, 0, !RealQty))
            Printer.Print sTemp
            '실수량(조정수량)
            Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11300
            RSet sTemp = str(!CtrlQty)
            Printer.Print sTemp
            
            'LOSS
            Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11600
            RSet sTemp = str(!LossQty)
            Printer.Print sTemp
            '불량수
            Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11600
            RSet sTemp = str(IIf(IsNull(!DefectQty), 0, !DefectQty)) '불량수
            Printer.Print sTemp
            
            '견본
            Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 12200
            RSet sTemp = str(!SampleQty)
            Printer.Print sTemp
            '난단
            Printer.CurrentX = BASE_X + 1560 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 12200
            RSet sTemp = str(!CutQty)
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

            If optPrint(0).Value Then
                If Not IsNull(!DefectID) Then
                    
                    Printer.Print CheckNull(rs!KDefect)
                End If
            Else
                If Not IsNull(!DefectID) Then       '영문일때는 간격때문에 태그명을 찍어준다.
                    Printer.Print CheckNull(rs!TagName)
                End If
            End If
            
            ' 불량 Detail
            nPenalty = 0

            Set rsData = oInspect.GetInspectSub(sOrderID, rs!RollSeq)
            Do Until rsData.EOF
                Printer.Font.Size = 5
                Printer.CurrentX = BASE_X + 1200 + (1000 * nRoll)

                If optRange(0).Value = True Then
                    Printer.CurrentY = BASE_Y + 900 + (val(rsData!yPos) * 100)
                Else
                    Printer.CurrentY = BASE_Y + 900 + (val(rsData!yPos) * 200)
                End If
                Printer.Print m_nDefectName(val(Right(rsData!DefectID, 2))).Defect & str(val(rsData!Demerit) / 10) & ""
                
                nPenalty = nPenalty + rsData!Demerit
                rsData.MoveNext
            Loop
            rsData.Close
            Set rsData = Nothing


            '불량수  벌점
            Printer.Font.Size = 7
            Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
            Printer.CurrentY = BASE_Y + 11900
            RSet sTemp = str(IIf(IsNull(!DefectQty), 0, !DefectQty))
            
            
'            Printer.CurrentX = BASE_X + 1060 + (1000 * nRoll)
'            Printer.CurrentY = BASE_Y + 11900
'            RSet sTemp = CheckNull(!BatNo) 'Bat No
'            Printer.Print sTemp
            
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
    Printer.CurrentY = BASE_Y - 500
    Printer.CurrentX = BASE_X + 9000
    Printer.Print "Page No : " & m_nPageCnt(0) & " / " & m_nPageCnt(1)
    
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
    Printer.Line (BASE_X, BASE_Y)-(BASE_X, BASE_Y + 14300)
    Printer.Line (BASE_X, BASE_Y)-(BASE_X + 11000, BASE_Y)
    Printer.Line (BASE_X + 11000, BASE_Y)-(BASE_X + 11000, BASE_Y + 14300)
    Printer.Line (BASE_X, BASE_Y + 14300)-(BASE_X + 11000, BASE_Y + 14300)

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
        Printer.Print "        LOT"
    Else
        Printer.Print "        LOT"
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
        
    Printer.Print ""

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

    Set rs = oCode.GetCode("1")
    Set oCode = Nothing

    Do Until rs.EOF
        nDefectNo = val(Right(rs!DefectID, 2))
        If val(nDefectNo) < DEFECT_COUNT Then
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

