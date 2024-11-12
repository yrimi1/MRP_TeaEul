VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccountByCustom 
   ClientHeight    =   9255
   ClientLeft      =   1665
   ClientTop       =   1470
   ClientWidth     =   15180
   Icon            =   "frmAccountByCustom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15180
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   5445
      TabIndex        =   11
      Top             =   75
      Width           =   1410
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7680
      Left            =   0
      TabIndex        =   8
      Top             =   780
      Width           =   15165
      _cx             =   26749
      _cy             =   13547
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
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   8580
      TabIndex        =   5
      Top             =   75
      Width           =   1410
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   720
      Left            =   14100
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "자료 검색"
      Top             =   30
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   480
      Left            =   1485
      TabIndex        =   1
      Top             =   75
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23724033
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   480
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   75
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   847
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "수불일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11775
      TabIndex        =   3
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
      Left            =   13500
      TabIndex        =   4
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
      Index           =   4
      Left            =   7365
      TabIndex        =   6
      Top             =   75
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품   명"
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   7
         Top             =   45
         Width           =   1095
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   4230
      TabIndex        =   9
      Top             =   75
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거래처"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   0
      Left            =   6885
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   75
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      Enabled         =   0   'False
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   10020
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   75
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      Enabled         =   0   'False
      ButtonStyle     =   3
      Outline         =   0   'False
   End
End
Attribute VB_Name = "frmAccountByCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\AccountByCustom.rpt"


Private Sub cmdFind_Click(Index As Integer)
    If Index = 0 Then
        Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
    Else
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If
End Sub

Private Sub cmdPrint_Click()
    Call ReportPrint
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub


Private Sub Form_Load()
    Dim i%
    
    'me.Move 0, 0, 11985, 9660
    Me.Move 0, 0, 15300, 9660

    Call SetOperate(Me)
    
    cmdPrint.Picture = LoadResPicture("PRINT", vbResIcon)
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    
    dtpDate = Now
    
    Call InitGrid

    Show

End Sub



Private Sub chkSearch_Click(Index As Integer)
    If chkSearch(Index) Then
        txtSearch(Index).Enabled = True
        txtSearch(Index).SetFocus
        
        If Index = 0 Then
            cmdFind(0).Enabled = True
        ElseIf Index = 2 Then
            cmdFind(1).Enabled = True
        End If
    Else
        txtSearch(Index).Enabled = False
        
        If Index = 0 Then
            cmdFind(0).Enabled = False
        ElseIf Index = 2 Then
            cmdFind(1).Enabled = False
        End If
    End If
End Sub



Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub


Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
        End If
    End If
End Sub


Private Sub cmdSearch_Click()
    Dim oAccount As PlusLib2.CSubul
    Dim rs       As Recordset
    Dim i%, iNowRow%
    Dim sDate$
    Dim nChkCustom%, sCustomID$
    Dim nChkArticle%, sArticleID$
    Dim nPreRoll&, nPreQty&, nPrePrice As Single
    Dim nInRoll&, nInQty&, nInPrice As Single
    Dim nOutRoll&, nOutQty&, nOutPrice As Single
    Dim nNowQty&, nNowRoll&, nNowPrice As Single
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    sDate = MakeDate(DF_SHORT, dtpDate)
    nChkCustom = IIf(chkSearch(0).Value = vbChecked, 1, 0)
    sCustomID = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(2).Value = vbChecked, 1, 0)
    sArticleID = txtSearch(2).Tag
   
    Set oAccount = New PlusLib2.CSubul
    oAccount.Connection = g_adoCon

    Set rs = oAccount.GetAccountByCustom(sDate, nChkCustom, sCustomID, nChkArticle, sArticleID)
    
    Set oAccount = Nothing

    With grdData
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            nPreRoll = rs!PreRoll
            nPreQty = rs!PreQty
            nPrePrice = rs!PrePrice
            
            nInRoll = rs!InRoll + rs!InReversRoll
            nInQty = rs!InQty + rs!InReversQty
            nInPrice = rs!InPrice + rs!InReversPrice
            
            nOutRoll = rs!OutRoll + rs!OutReversRoll
            nOutQty = rs!OutQty + rs!OutReversQty
            nOutPrice = rs!OutPrice + rs!OutReversPrice
            
            nNowRoll = nPreRoll + nInRoll - nOutRoll
            nNowQty = nPreQty + nInQty - nOutQty
            nNowPrice = nPrePrice + nInPrice - nOutPrice
            
            .AddItem CStr(i) & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & _
                rs!KCustom & vbTab & rs!Article & vbTab & " " & vbTab & " " & vbTab & _
                SetCurrency(rs!PreRoll) & vbTab & SetCurrency(rs!PreQty) & vbTab & SetCurrency(rs!PrePrice) & vbTab & " " & vbTab & _
                SetCurrency(rs!InRoll) & vbTab & SetCurrency(rs!InQty) & vbTab & SetCurrency(rs!InPrice) & vbTab & " " & vbTab & _
                SetCurrency(rs!InReversRoll) & vbTab & SetCurrency(rs!InReversQty) & vbTab & SetCurrency(rs!InReversPrice) & vbTab & " " & vbTab & _
                SetCurrency(rs!OutRoll) & vbTab & SetCurrency(rs!OutQty) & vbTab & SetCurrency(rs!OutPrice) & vbTab & " " & vbTab & _
                SetCurrency(rs!OutReversRoll) & vbTab & SetCurrency(rs!OutReversQty) & vbTab & SetCurrency(rs!OutReversPrice) & vbTab & " " & vbTab & _
                SetCurrency(nNowRoll) & vbTab & SetCurrency(nNowQty) & vbTab & SetCurrency(nNowPrice)

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

        .Redraw = flexRDDirect

        .SetFocus
    End With

    Screen.MousePointer = vbArrow

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oAccount = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub





Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Cols = 31
        Call SetVSFlexGrid(grdData)
        
        .ScrollBars = flexScrollBarBoth
        '.FixedRows = 2
        .Rows = 2
        .FixedRows = 2
        .FixedCols = 6
        
        .RowHeight(0) = 350
        .RowHeight(1) = 350

        .Redraw = flexRDNone

        .TextMatrix(0, 0) = " "
        .TextMatrix(0, 1) = " ":            .ColWidth(1) = 0
        .TextMatrix(0, 2) = " ":            .ColWidth(2) = 0:            .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = " ":            .ColWidth(3) = 0:            .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "거래처":       .ColWidth(4) = 1300:            .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(0, 5) = "품명":         .ColWidth(5) = 2000:            .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(0, 6) = " ":            .ColWidth(6) = 0
        .TextMatrix(0, 7) = " ":            .ColWidth(7) = 0
        .TextMatrix(0, 8) = "전월재고":     .ColWidth(8) = 850:             .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(0, 9) = "전월재고":     .ColWidth(9) = 1000:            .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(0, 10) = "전월재고":    .ColWidth(10) = 1000:           .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(0, 11) = " ":           .ColWidth(11) = 0
        .TextMatrix(0, 12) = "입고량":      .ColWidth(12) = 850:            .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(0, 13) = "입고량":      .ColWidth(13) = 1000:           .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(0, 14) = "입고량":      .ColWidth(14) = 1000:           .ColAlignment(14) = flexAlignRightCenter
        .TextMatrix(0, 15) = " ":           .ColWidth(15) = 0
        .TextMatrix(0, 16) = "반출량":      .ColWidth(16) = 850:            .ColAlignment(16) = flexAlignRightCenter
        .TextMatrix(0, 17) = "반출량":      .ColWidth(17) = 1000:           .ColAlignment(17) = flexAlignRightCenter
        .TextMatrix(0, 18) = "반출량":      .ColWidth(18) = 1000:           .ColAlignment(18) = flexAlignRightCenter
        .TextMatrix(0, 19) = " ":           .ColWidth(19) = 0
        .TextMatrix(0, 20) = "출고량":      .ColWidth(20) = 850:            .ColAlignment(20) = flexAlignRightCenter
        .TextMatrix(0, 21) = "출고량":      .ColWidth(21) = 1000:           .ColAlignment(17) = flexAlignRightCenter
        .TextMatrix(0, 22) = "출고량":      .ColWidth(22) = 1000:           .ColAlignment(18) = flexAlignRightCenter
        .TextMatrix(0, 23) = " ":           .ColWidth(23) = 0
        .TextMatrix(0, 24) = "반입량":      .ColWidth(24) = 850:            .ColAlignment(16) = flexAlignRightCenter
        .TextMatrix(0, 25) = "반입량":      .ColWidth(25) = 1000:           .ColAlignment(17) = flexAlignRightCenter
        .TextMatrix(0, 26) = "반입량":      .ColWidth(26) = 1000:           .ColAlignment(18) = flexAlignRightCenter
        .TextMatrix(0, 27) = " ":           .ColWidth(27) = 0
        .TextMatrix(0, 28) = "당월재고":    .ColWidth(28) = 850:            .ColAlignment(24) = flexAlignRightCenter
        .TextMatrix(0, 29) = "당월재고":    .ColWidth(29) = 850:            .ColAlignment(25) = flexAlignRightCenter
        .TextMatrix(0, 30) = "당월재고":    .ColWidth(30) = 850:            .ColAlignment(26) = flexAlignRightCenter
        
        .TextMatrix(1, 0) = " "
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = " "
        .TextMatrix(1, 3) = " "
        .TextMatrix(1, 4) = "거래처"
        .TextMatrix(1, 5) = "품명"
        .TextMatrix(1, 6) = " "
        .TextMatrix(1, 7) = " "
        .TextMatrix(1, 8) = "절수"
        .TextMatrix(1, 9) = "수량"
        .TextMatrix(1, 10) = "금액"
        .TextMatrix(1, 11) = " "
        .TextMatrix(1, 12) = "절수"
        .TextMatrix(1, 13) = "수량"
        .TextMatrix(1, 14) = "금액"
        .TextMatrix(1, 15) = " "
        .TextMatrix(1, 16) = "절수"
        .TextMatrix(1, 17) = "수량"
        .TextMatrix(1, 18) = "금액"
        .TextMatrix(1, 19) = " "
        .TextMatrix(1, 20) = "절수"
        .TextMatrix(1, 21) = "수량"
        .TextMatrix(1, 22) = "금액"
        .TextMatrix(1, 23) = " "
        .TextMatrix(1, 24) = "절수"
        .TextMatrix(1, 25) = "수량"
        .TextMatrix(1, 26) = "금액"
        .TextMatrix(1, 27) = " "
        .TextMatrix(1, 28) = "절수"
        .TextMatrix(1, 29) = "수량"
        .TextMatrix(1, 30) = "금액"
        
        .MergeCells = flexMergeFixedOnly
        '.MergeCells = flexMergeRestrictColumns
        
        .MergeRow(0) = True

        For i = 0 To 7
            .MergeCol(i) = True
        Next i

    
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub ReportPrint()
    Dim oAccount As PlusLib2.CSubul
    Dim rs       As Recordset
    Dim i%
    Dim sParam() As String
    Dim sDate$
    Dim nChkCustom%, sCustomID$
    Dim nChkArticle%, sArticleID$
            
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    sDate = MakeDate(DF_SHORT, dtpDate)
    nChkCustom = IIf(chkSearch(0).Value = vbChecked, 1, 0)
    sCustomID = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(2).Value = vbChecked, 1, 0)
    sArticleID = txtSearch(2).Tag
    
    Set oAccount = New PlusLib2.CSubul
    oAccount.Connection = g_adoCon

    Set rs = oAccount.GetAccountByCustom(sDate, nChkCustom, sCustomID, nChkArticle, sArticleID)
    
    Set oAccount = Nothing
    
    ReDim sParam(0)

    sParam(0) = "조양염직"
  
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oAccount = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub



