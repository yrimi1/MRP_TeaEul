VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectDefectTotal 
   ClientHeight    =   9255
   ClientLeft      =   1440
   ClientTop       =   1380
   ClientWidth     =   11865
   Icon            =   "frmInspectDefectTotal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   2
      Left            =   7980
      TabIndex        =   25
      Top             =   450
      Width           =   1935
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7275
      Left            =   0
      TabIndex        =   17
      Top             =   1170
      Width           =   11835
      _cx             =   20876
      _cy             =   12832
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   50
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
   Begin Threed.SSCommand cmdHtml 
      Height          =   690
      Left            =   6720
      TabIndex        =   18
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      HTML(&H)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   3
      Left            =   7980
      TabIndex        =   12
      Top             =   810
      Width           =   1935
   End
   Begin VB.Frame fraOrder 
      Height          =   795
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   495
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1710
      MousePointer    =   99  '사용자 정의
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2355
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1710
      MousePointer    =   99  '사용자 정의
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2355
      MousePointer    =   99  '사용자 정의
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   780
      Left            =   11055
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   16
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   7980
      TabIndex        =   14
      Top             =   105
      Width           =   1935
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   8460
      TabIndex        =   19
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   270
      Top             =   8730
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   20
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   4470
      TabIndex        =   7
      Top             =   105
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy-MM-dd (ddd)"
      Format          =   54788099
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   4470
      TabIndex        =   9
      Top             =   450
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy-MM-dd (ddd)"
      Format          =   54788099
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   6705
      TabIndex        =   21
      Top             =   105
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   9945
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   105
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlOrder 
      Height          =   300
      Left            =   6705
      TabIndex        =   22
      Top             =   810
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "Order No."
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "관리번호"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   11
         Top             =   45
         Width           =   1140
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   3120
      TabIndex        =   23
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "검사일자"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   24
         Top             =   45
         Value           =   1  '확인
         Width           =   1035
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   6705
      TabIndex        =   26
      Top             =   450
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   27
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   2
      Left            =   9945
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   450
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   1
      Left            =   5775
      TabIndex        =   10
      Top             =   510
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   0
      Left            =   5775
      TabIndex        =   8
      Top             =   180
      Width           =   360
   End
End
Attribute VB_Name = "frmInspectDefectTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\InspectDefectTotal.rpt"

Private m_nDefectCount As Integer

Private Sub Form_Load()
    Dim i%
    Me.Move 0, 0, 11970, 9660

    dtpDate(0) = Now
    dtpDate(1) = Now

    Call SetOperate(Me)
    
    Call InitGrid

    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
    txtSearch(3).Enabled = False
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
        If optOrder(0).Value Then
            .ColWidth(1) = 0
            .ColWidth(2) = 1290
            chkSearch(1).Caption = "Order No."
        Else
            .ColWidth(1) = 1290
            .ColWidth(2) = 0
            chkSearch(1).Caption = "관리번호"
        End If
    End With
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then '[0] 수주일자 선택
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else '[1, 2] 거래처, 관리번호 선택
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

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub

    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, 0, False, txtSearch(Index))
    End If
    cmdSearch.SetFocus
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If
    cmdSearch.SetFocus
End Sub

Private Sub cmdSearch_Click()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim i%

    On Error GoTo ErrHandle

    Call InitGrid

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
    oInspect.UserName = g_sUserName
    Set rs = oInspect.GetDefectTotal(IIf(chkSearch(0).Value = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
        IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
        IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0), 2, 1), 0), txtSearch(3))
    
    Set oInspect = Nothing

    With grdData
        .Redraw = False

        .Rows = .FixedRows
        Do While Not rs.EOF
            i = i + 1
            .AddItem CStr(i) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & rs!KCustom & vbTab & _
                rs!Article & vbTab & rs!Color & vbTab & _
                rs!D1 & vbTab & rs!D2 & vbTab & rs!D3 & vbTab & rs!D4 & vbTab & rs!D5 & vbTab & rs!D6 & vbTab & rs!D7 & vbTab & _
                rs!D8 & vbTab & rs!D9 & vbTab & rs!D10 & vbTab & rs!D11 & vbTab & rs!D12 & vbTab & rs!D13 & vbTab & rs!D14 & vbTab & _
                rs!D15 & vbTab & rs!D16 & vbTab & rs!D17 & vbTab & rs!D18 & vbTab & rs!D19 & vbTab & rs!D20 & vbTab & rs!D21 & vbTab & _
                rs!D22 & vbTab & rs!D23 & vbTab & rs!D24 & vbTab & rs!D25 & vbTab & rs!D26 & vbTab & rs!D27 & vbTab & rs!D28 & vbTab & _
                rs!D29 & vbTab & rs!D30 & vbTab & rs!D31 & vbTab & rs!D32 & vbTab & rs!D33 & vbTab & rs!D34 & vbTab & rs!D35 & vbTab & _
                rs!D36 & vbTab & rs!D37 & vbTab & rs!D38 & vbTab & rs!D39 & vbTab & rs!D40 & vbTab & rs!D41 & vbTab & rs!D42 & vbTab & _
                rs!D43 & vbTab & rs!D44 & vbTab & rs!D45 & vbTab & rs!D46 & vbTab & rs!D47
                
'            If (i Mod 2) = 0 Then
'                .Row = .FixedRows + i - 1
'                .Col = .FixedCols
'                .ColSel = .Cols - 1
'                .CellBackColor = COLOR_GRIDROW
'            End If

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            cmdHtml.Enabled = True
            cmdExcel.Enabled = True
        Else
            cmdHtml.Enabled = False
            cmdExcel.Enabled = False
            
            MsgBox LoadResString(203), vbInformation
        End If

        .Redraw = True
    End With

    Exit Sub

ErrHandle:
    Set rs = Nothing
    Set oInspect = Nothing
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Sub

Private Sub cmdHTML_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    grdData.Cols = m_nDefectCount
    If MakeHtmlGrid(grdData, "C:\" & Me.Caption & ".html") Then
        Call RelateOpen(Me.hwnd, "C:\" & Me.Caption & ".html")
    End If
End Sub

Private Sub cmdExcel_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    grdData.Cols = m_nDefectCount
    Call MakeExcelGrid(grdData)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim oInspect As PlusLib2.CInspect
    Dim rs    As Recordset
    Dim i%, j%

    On Error GoTo ErrHandler

    With grdData
        .Redraw = flexRDDirect
        .Cols = 54
        
        Call SetVSFlexGrid(grdData)
        .Rows = 2

        .FixedCols = 1
        .FixedRows = 2
        .FrozenCols = 5
        .ScrollBars = flexScrollBarBoth

        .RowHeight(0) = 250
        .RowHeight(1) = 250

        .TextArray(0) = "순위":         .ColWidth(0) = 400
        .TextArray(1) = "관리번호":     .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order No":     .ColWidth(2) = 0:       .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "거래처":       .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "품명":         .ColWidth(4) = 2000:    .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "색상명":       .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignLeftCenter

        .TextArray(.Cols) = "순위"
        .TextArray(.Cols + 1) = "관리번호"
        .TextArray(.Cols + 2) = "Order No"
        .TextArray(.Cols + 3) = "거래처"
        .TextArray(.Cols + 4) = "품명"
        .TextArray(.Cols + 5) = "색상명"

        Set oInspect = New PlusLib2.CInspect
        oInspect.Connection = g_adoCon

        Set rs = oInspect.GetDefectByLang(1)
        Set oInspect = Nothing

        m_nDefectCount = 6 + rs.RecordCount
        .Cols = m_nDefectCount
        i = 6
        Do While Not rs.EOF
            
            .TextArray(i) = IIf(rs!DefectClss = "1", "가공", "제직") & " 불량"
            .TextMatrix(.Rows - 1, i) = rs!Display
            .ColWidth(i) = Max(1000, TextWidth(Trim(rs!Display)) + 250)
            .ColDataType(i) = flexDTDecimal
            .ColFormat(i) = "#,##0"

            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        For j = i To .Cols - 1
            .ColHidden(j) = True
        Next j

        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i

        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True

        .MergeRow(0) = True

        .WordWrap = False
        .Redraw = flexRDDirect
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, "frmInspectDefectTotal.InitGrid", Err.Description)
End Sub

Private Function Max(Value1, Value2)
    If Value1 > Value2 Then
        Max = Value1
    Else
        Max = Value2
    End If
End Function

