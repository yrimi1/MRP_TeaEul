VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectTotal 
   ClientHeight    =   9255
   ClientLeft      =   4500
   ClientTop       =   1455
   ClientWidth     =   11865
   Icon            =   "frmInspectTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   0
      Left            =   7215
      TabIndex        =   18
      Top             =   105
      Width           =   1935
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7605
      Left            =   0
      TabIndex        =   15
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
      TabIndex        =   12
      Top             =   -15
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   495
         Value           =   -1  'True
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2070
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   3930
      TabIndex        =   5
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23658497
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   3930
      TabIndex        =   6
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23658497
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2715
      TabIndex        =   7
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "검사 일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Index           =   0
      Left            =   8370
      TabIndex        =   8
      Top             =   8520
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      집계표(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   9
      Top             =   8520
      Width           =   1680
      _ExtentX        =   2963
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Index           =   1
      Left            =   6570
      TabIndex        =   16
      Top             =   8520
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      업무결산(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   4800
      TabIndex        =   17
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   5940
      TabIndex        =   19
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
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Left            =   9180
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   105
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      Enabled         =   0   'False
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   3
      Left            =   5205
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   510
      Width           =   360
   End
End
Attribute VB_Name = "frmInspectTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE_1 = "\Report\InspectTotal.rpt"
Private Const REPORTFILE_2 = "\Report\InspectTotalByDate.rpt"


Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(Index).Value Then
            cmdFind.Enabled = True
            txtSearch(0).SetFocus
        Else
            cmdFind.Enabled = False
        End If
    End If
End Sub

Private Sub cmdExcel_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    Call MakeExcelGrid(grdData)
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)
    
    cmdPrint(0).Picture = LoadResPicture("PRINT", vbResIcon)
    cmdPrint(1).Picture = LoadResPicture("PRINT", vbResIcon)
    cmdFind.Picture = LoadResPicture("FIND", vbResIcon)
    
    dtpDate(0) = Now
    dtpDate(1) = Now

    Call InitGrid

    Show

End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub cmdSearch_Click()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%, iNowRow%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetInspectTotal(1, MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(0), 1, 0), txtSearch(0).Tag)
    Set oInspect = Nothing

    With grdData
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!kCustom & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Article & vbTab & rs!Color & vbTab & _
                rs!ColorQty & vbTab & rs!Cost & vbTab & rs!INQty & vbTab & _
                rs!SetQty & vbTab & IIf(rs!OrderUnit = "0", rs!SetQty * rs!Cost, Int(rs!SetQty / 0.9144) * rs!Cost) & vbTab & _
                IIf(rs!OrderUnit = "0", rs!InspectQty, rs!InspectQtyY) & vbTab & IIf(rs!OrderUnit = "0", rs!InspectQty * rs!Cost, Int(rs!InspectQty / 0.9144) * rs!Cost) & vbTab & _
                rs!OutQty & vbTab & IIf(rs!OrderUnit = "0", rs!OutQty * rs!Cost, Int(rs!OutQty / 0.9144) * rs!Cost)

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

'        Call ChangeScroll

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

Private Sub cmdPrint_Click(Index As Integer)
    If Index = 0 Then
        Call PrintInspectTotal
    Else
        Call PrintWorkEndByDate
    End If
End Sub

Private Sub PrintWorkEndByDate()
    Dim oInspect As PlusLib2.CInspect
    Dim rs As ADODB.Recordset
    Dim i%
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    ReDim sParam(0)
    
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
   
    Set rs = oInspect.GetInspectTotal(1, MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(0)), IIf(chkSearch(0), 1, 0), txtSearch(0).Tag)
    Set oInspect = Nothing

    sParam(0) = MakeDate(DF_LONG, dtpDate(0))
    
    Call PrintReport(REPORTFILE_1, rs, sParam, PlusMDI.PrintPreview)
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    Set oInspect = Nothing
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "PrintWorkEndByDate", Err.Description)
End Sub

Private Sub PrintInspectTotal()
    Dim oInspect As PlusLib2.CInspect
    Dim rs As ADODB.Recordset
    Dim i%, sSDate$
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    ReDim sParam(8)
    
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
    
    sSDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6) & "01"
    Set rs = oInspect.GetInspectTotalMonth(sSDate, MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(0), 1, 0), txtSearch(0).Tag)
    
    sParam(1) = Format(rs!INQty, "#,###")
    sParam(2) = Format(rs!InCost, "#,###")
    sParam(3) = Format(rs!SetQty, "#,###")
    sParam(4) = Format(rs!SetCost, "#,###")
    sParam(5) = Format(rs!InspectQty, "#,###")
    sParam(6) = Format(rs!InspectCost, "#,###")
    sParam(7) = Format(rs!OutQty, "#,###")
    sParam(8) = Format(rs!OutCost, "#,###")
    rs.Close
    
    Set rs = oInspect.GetInspectTotal(1, MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), IIf(chkSearch(0), 1, 0), txtSearch(0).Tag)
    Set oInspect = Nothing
    

    sParam(0) = MakeDate(DF_LONG, dtpDate(0)) & " ~ " & MakeDate(DF_LONG, dtpDate(1))
    
    Call PrintReport(REPORTFILE_2, rs, sParam, PlusMDI.PrintPreview)
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    Set oInspect = Nothing
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "PrintInspectTotal", Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Cols = 15
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .TextArray(1) = "거래처":       .ColWidth(1) = 1200:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "Order No.":    .ColWidth(2) = 0:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호":     .ColWidth(3) = 1350:           .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "품명":         .ColWidth(4) = 1500:        .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "색상명":       .ColWidth(5) = 1200:        .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "수주량":       .ColWidth(6) = 800:         .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "단가":         .ColWidth(7) = 450:         .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "입고량":       .ColWidth(8) = 800:         .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "투입량":       .ColWidth(9) = 800:         .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "투입금액":    .ColWidth(10) = 1000:       .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "생산량":      .ColWidth(11) = 800:        .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "생산금액":    .ColWidth(12) = 1000:       .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "출고량":      .ColWidth(13) = 800:        .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "출고금액":    .ColWidth(14) = 1000:       .ColAlignment(14) = flexAlignRightCenter
        
        For i = 6 To 14
            .ColFormat(i) = "#,###"
        Next i
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub

    If Index = 0 Then Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    cmdSearch.SetFocus
End Sub

Private Sub cmdFind_Click()
    Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(0))
    cmdSearch.SetFocus
End Sub

