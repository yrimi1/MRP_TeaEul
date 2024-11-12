VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectDate 
   ClientHeight    =   10710
   ClientLeft      =   90
   ClientTop       =   2655
   ClientWidth     =   15180
   Icon            =   "frmInspectDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   15180
   Begin VB.Frame fraSearch 
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   15255
      Begin VB.TextBox txtSearch 
         Height          =   285
         Index           =   3
         Left            =   6960
         TabIndex        =   15
         Top             =   135
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Index           =   4
         Left            =   3600
         TabIndex        =   13
         Top             =   120
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Index           =   5
         Left            =   3600
         TabIndex        =   14
         Top             =   435
         Width           =   1545
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "전일"
         Height          =   315
         Index           =   0
         Left            =   30
         MousePointer    =   99  '사용자 정의
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   420
         Width           =   585
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   630
         MousePointer    =   99  '사용자 정의
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   420
         Width           =   585
      End
      Begin VB.OptionButton optClss 
         Caption         =   "가공불량"
         Height          =   180
         Index           =   2
         Left            =   9105
         TabIndex        =   10
         Top             =   330
         Width           =   1110
      End
      Begin VB.OptionButton optClss 
         Caption         =   "제직불량"
         Height          =   180
         Index           =   1
         Left            =   10320
         TabIndex        =   9
         Top             =   330
         Width           =   1110
      End
      Begin VB.OptionButton optClss 
         Caption         =   "전체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   8220
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   750
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   600
         Left            =   11460
         TabIndex        =   5
         Top             =   135
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1058
         _Version        =   196609
         Caption         =   "      검색(&F)"
         PictureAlignment=   1
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   0
         Top             =   120
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116785153
         CurrentDate     =   36271
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "검사 일자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   600
         Left            =   13950
         TabIndex        =   3
         Top             =   135
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1058
         _Version        =   196609
         Caption         =   "      닫기(&X)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   600
         Left            =   12705
         TabIndex        =   4
         Top             =   135
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1058
         _Version        =   196609
         Caption         =   "      인쇄(&P)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1230
         TabIndex        =   7
         Top             =   420
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116785153
         CurrentDate     =   36271
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   5580
         TabIndex        =   16
         Top             =   135
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "기    계"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   60
            Width           =   1215
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   4
         Left            =   5160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
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
         Index           =   10
         Left            =   2640
         TabIndex        =   19
         Top             =   435
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품  명"
            Height          =   180
            Index           =   5
            Left            =   45
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   60
            Width           =   780
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   5
         Left            =   5160
         TabIndex        =   21
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
         Index           =   6
         Left            =   2640
         TabIndex        =   25
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거래처"
            Height          =   180
            Index           =   4
            Left            =   45
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   60
            Width           =   840
         End
      End
      Begin VB.Frame fraOrder 
         Height          =   420
         Left            =   5580
         TabIndex        =   22
         Top             =   345
         Width           =   2595
         Begin VB.OptionButton optOrder 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   1365
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   165
            Width           =   1155
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdInspect 
      Height          =   8595
      Left            =   0
      TabIndex        =   6
      Top             =   660
      Width           =   15255
      _cx             =   26908
      _cy             =   15161
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   14737632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
      ExtendLastCol   =   -1  'True
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
Attribute VB_Name = "frmInspectDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bloading As Boolean

Private Sub chkSearch_Click(Index As Integer)
    If Index = 4 Or Index = 5 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            cmdFind(Index).Enabled = True
            txtSearch(Index).SetFocus
        Else
            txtSearch(Index).Enabled = False
            cmdFind(Index).Enabled = False
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 4 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 5 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    End If

End Sub

Private Sub cmdPrint_Click()
Dim irow As Integer

    On Error GoTo Err_Handler

    Screen.MousePointer = vbHourglass

    With grdInspect
        .Redraw = flexRDNone
        .RowHeight(0) = 600
        .RowHeight(1) = 350
        .Cell(flexcpFontSize, 2, 0, .Rows - 1, .Cols - 1) = 7
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = vbWhite
        For irow = 3 To .Rows - 1
            If .RowHeight(irow) <> 0 Then
                .RowHeight(irow) = 300
            End If
        Next irow
        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite
        
        .PrintGrid "태을염직", True, 2, 100, 500
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .Cell(flexcpFontSize, 2, 0, .Rows - 1, .Cols - 1) = 8
        For irow = 3 To .Rows - 1
            If .RowHeight(irow) <> 0 Then
                .RowHeight(irow) = 500
            End If
        Next irow
        .GridLinesFixed = flexGridInset
        .GridColorFixed = 0
        
        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbDefault
    MsgBox "인쇄 되었습니다", vbInformation, "인쇄"

    Exit Sub

Err_Handler:
    MsgBox "발행 오류 : " & Err.Description, vbExclamation, "발행 Error"
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpDate(0) = DateSerial(Year(Date), Month(Date), Day(Date) - 1)
            dtpDate(1) = DateSerial(Year(Date), Month(Date), Day(Date) - 1)
        Case 1
            dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
            dtpDate(1) = Date
    End Select
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
    
    cmdPrint.Picture = LoadResPicture("PRINT", vbResIcon)
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
End Sub


Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 15300, 9660
    
    dtpDate(0) = Now
    dtpDate(1) = Now

    For i = 4 To 5
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
        txtSearch(i).Enabled = False
    Next i

    Call InitGrid

    cmdSearch.Picture = LoadResPicture("FIND", vbResIcon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub InitGrid()
    Dim i%
    Dim irow As Integer

    With grdInspect
        .FixedRows = 3:     .FixedCols = 0
        .Rows = 3:          .Cols = 20
        .WordWrap = True
        .Redraw = flexRDNone
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical

        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 400
        .FontSize = 8
        .Cell(flexcpFontBold, 0, 0, 2, .Cols - 1) = True
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .TextMatrix(2, 0) = "일자":                     .ColWidth(0) = 550:     .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(2, 1) = "카드":                     .ColWidth(1) = 750:    .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(2, 2) = "거래선":                   .ColWidth(2) = 800:     .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(2, 3) = "오더No":                   .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(2, 4) = "관리No":                   .ColWidth(4) = 700:     .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(2, 5) = "품명":                     .ColWidth(5) = 1300:    .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(2, 6) = "구분":                     .ColWidth(6) = 750:     .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(2, 7) = "규격":                     .ColWidth(7) = 550:     .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(2, 8) = "순위":                     .ColWidth(8) = 0:       .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(2, 9) = "색상":                     .ColWidth(9) = 1200:    .ColAlignment(9) = flexAlignLeftCenter
        .TextMatrix(2, 10) = "수주":                    .ColWidth(10) = 800:    .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(2, 11) = "Lot":                     .ColWidth(11) = 400:    .ColAlignment(11) = flexAlignCenterCenter
        .TextMatrix(2, 12) = "투입":                    .ColWidth(12) = 850:    .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(2, 13) = "검사":                    .ColWidth(13) = 0:      .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(2, 14) = "합격":                    .ColWidth(14) = 850:    .ColAlignment(14) = flexAlignRightCenter
        .TextMatrix(2, 15) = "가공 불량":               .ColWidth(15) = 2700:   .ColAlignment(15) = flexAlignLeftCenter
        .TextMatrix(2, 16) = "제직 불량":               .ColWidth(16) = 1450:   .ColAlignment(16) = flexAlignLeftCenter
        .TextMatrix(2, 17) = "난단":                    .ColWidth(17) = 450:    .ColAlignment(17) = flexAlignCenterCenter
        .TextMatrix(2, 18) = "견본":                    .ColWidth(18) = 450:    .ColAlignment(18) = flexAlignCenterCenter
        .TextMatrix(2, 19) = "호기":                    .ColWidth(19) = 450:    .ColAlignment(19) = flexAlignCenterCenter
        
        
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "[[ 검사실적 일보 ]]"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 14
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 0, 1, 15) = " "
        .Cell(flexcpText, 1, 16, 1, .Cols - 1) = "실적일자 : "
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = 10
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = True
        
        .MergeCells = flexMergeFixedOnly
        For irow = 0 To 2
            .MergeRow(irow) = True
        Next irow
'        .MergeCol(14) = True
'        .MergeCol(15) = True
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub grdInspect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdInspect
        If .MouseCol > -1 And .MouseRow > 1 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub grdInspect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    grdInspect.ToolTipText = grdInspect.TextMatrix(grdInspect.Row, grdInspect.Col)
End Sub

Private Sub cmdSearch_Click()
    If m_bloading Then Exit Sub
    Call FillGridOrder
End Sub

Private Sub FillGridOrder()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim tRs      As Recordset
    Dim i%, iNowRow%
    Dim iCnt As Integer
    Dim irow As Integer
    Dim iCol As Integer
    Dim nClss As Integer        '0: 전체, 1:제직불량만, 2:가공불량만
    Dim sDefect0 As String
    Dim sDefect1 As String
    Dim sCustom$, sExamDate$, sOrderID$, sOrderSeq$, sLotNo$, sCard$
    Dim nChkOrder%, nChkCustom%, nChkArticle%
    Dim sOrder$, sCustomID$, sArticle$

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    nClss = IIf(optClss(1).Value = True, 1, IIf(optClss(2).Value = True, 2, 0))
    
    nChkOrder = IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0)
    sOrder = txtSearch(3)
    nChkCustom = IIf(chkSearch(4).Value = vbChecked, 1, 0)
    sCustomID = txtSearch(4).Tag
    nChkArticle = IIf(chkSearch(5).Value = vbChecked, 1, 0)
    sArticle = txtSearch(5).Tag
    
    Set rs = oInspect.GetInspectByLot(MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), nClss, _
                                    nChkCustom, sCustomID, nChkArticle, sArticle, nChkOrder, sOrder)
        ' nClss     0: 전체, 1:제직불량만, 2:가공불량만

    If rs.RecordCount > 0 Then
        With grdInspect
            .Redraw = flexRDNone
            iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
            .Rows = .FixedRows
            For i = 1 To rs.RecordCount
                If i = 1 Then
                    sCustom = Trim(rs!kCustom)
                    sExamDate = rs!ExamDate
                    sOrderID = rs!OrderID
                    sOrderSeq = CStr(rs!OrderSeq)
                    sLotNo = rs!LotNo
                    sCard = Trim(rs!CardID) & Trim(rs!SplitID)
                End If
                DoEvents
                
                iCnt = 0:   sDefect0 = "":  sDefect1 = ""
                
                If rs!WeaveQty > 0 Or rs!DyeQty > 0 Then
                    Set tRs = oInspect.GetDefectArray(MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), rs!OrderID, rs!OrderSeq, rs!LotNo)
                    If tRs.RecordCount > 0 Then
                        If rs!DyeQty > 0 Then
                            sDefect0 = CStr(rs!DyeQty) & "="  ' 가공불량
                        End If
                        If rs!WeaveQty > 0 Then
                            sDefect1 = CStr(rs!WeaveQty) & "="    ' 제직불량
                        End If
                        
                        For iCnt = 1 To tRs.RecordCount
                            If tRs!DefectClss = "1" Then    ' 가공불량
                                sDefect0 = sDefect0 & tRs!KDefect & "(" & tRs!DyeRoll & "*" & tRs!DyeQty & ") "
                            Else
                                sDefect1 = sDefect1 & tRs!KDefect & "(" & tRs!WeaveRoll & "*" & tRs!WeaveQty & ") "
                            End If
                            tRs.MoveNext
                        Next iCnt
                    End If
                    Set tRs = Nothing
                End If
                
                If sCustom <> Trim(rs!kCustom) Then
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = 0
                    sCustom = rs!kCustom
                End If
                                 
                If i > 1 And sExamDate = rs!ExamDate And sOrderID = rs!OrderID And sOrderSeq = CStr(rs!OrderSeq) And sLotNo = rs!LotNo Then
                    If sCard <> Trim(rs!CardID) & Trim(rs!SplitID) Then
                        If Trim(.TextMatrix(.Rows - 1, 1)) = "" Then
                            If Trim(rs!CardID) = "" Then
                                .TextMatrix(.Rows - 1, 1) = ""
                                .TextMatrix(.Rows - 1, 12) = ""
                            Else
                                If Trim(rs!SplitID) = "" Then
                                    .TextMatrix(.Rows - 1, 1) = CStr(CLng(Right(rs!CardID, 4)))
                                Else
                                    .TextMatrix(.Rows - 1, 1) = CStr(CLng(Right(rs!CardID, 4))) & "(" & Trim(rs!SplitID) & ")"
                                End If
                                .TextMatrix(.Rows - 1, 12) = Format(rs!CardQty, "###,###")
                            End If
                        Else
                            If Trim(rs!CardID) = "" Then
                                .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
                                .TextMatrix(.Rows - 1, 12) = .TextMatrix(.Rows - 1, 12)
                            Else
                                If Trim(rs!SplitID) = "" Then
                                    .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1) & ", " & CStr(CLng(Right(rs!CardID, 4)))
                                Else
                                    .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1) & ", " & CStr(CLng(Right(rs!CardID, 4))) & "(" & Trim(rs!SplitID) & ")"
                                End If
                                .TextMatrix(.Rows - 1, 12) = Format(CLng(.TextMatrix(.Rows - 1, 12)) + rs!CardQty, "###,###")
                            End If
                        End If
                    End If
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = MakeDate(DF_MD, rs!ExamDate)
                    If Trim(rs!CardID) = "" Then
                        .TextMatrix(.Rows - 1, 1) = ""
                    Else
                        If Trim(rs!SplitID) = "" Then
                            .TextMatrix(.Rows - 1, 1) = CStr(CLng(Right(rs!CardID, 4)))
                        Else
                            .TextMatrix(.Rows - 1, 1) = CStr(CLng(Right(rs!CardID, 4))) & "(" & Trim(rs!SplitID) & ")"
                        End If
                    End If
                    .TextMatrix(.Rows - 1, 2) = Trim(rs!kCustom)
                    .TextMatrix(.Rows - 1, 3) = Trim(rs!OrderNo)
                    .TextMatrix(.Rows - 1, 4) = MakeOrderID(rs!OrderID, OM_COMPACT)
                    .TextMatrix(.Rows - 1, 5) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 6) = Trim(rs!WorkName)
                    .TextMatrix(.Rows - 1, 7) = rs!WorkWidth
                    .TextMatrix(.Rows - 1, 8) = rs!OrderSeq
                    .TextMatrix(.Rows - 1, 9) = Trim(rs!Color)
                    .TextMatrix(.Rows - 1, 10) = Format(rs!ColorQty, "###,###") & IIf(rs!UnitClss = "1", "M", "")
                    .TextMatrix(.Rows - 1, 11) = rs!LotNo
'                    .TextMatrix(.Rows - 1, 12) = Format(rs!StuffQty, "###,###")    ' 투입(Inspect)
                    .TextMatrix(.Rows - 1, 12) = Format(rs!CardQty, "###,###")      ' 투입(Card)
                    .TextMatrix(.Rows - 1, 13) = Format(rs!CtrlQty, "###,###")      ' 검사
                    .TextMatrix(.Rows - 1, 14) = Format(rs!PassQty, "###,###")      ' 합격
                    .TextMatrix(.Rows - 1, 15) = sDefect0                           ' 가공
                    .TextMatrix(.Rows - 1, 16) = sDefect1                           ' 제직
                    .TextMatrix(.Rows - 1, 17) = Format(rs!CutQty, "###,###")       ' 난단
                    .TextMatrix(.Rows - 1, 18) = Format(rs!SampleQty, "###,###")    ' 견본
                    .TextMatrix(.Rows - 1, 19) = Trim(rs!machID)
                    
                    .RowHeight(.Rows - 1) = 500
                End If
                sExamDate = rs!ExamDate
                sOrderID = rs!OrderID
                sOrderSeq = CStr(rs!OrderSeq)
                sLotNo = rs!LotNo
                sCard = Trim(rs!CardID) & Trim(rs!SplitID)
                
    
    
                rs.MoveNext
            Next i
            rs.Close
            Set rs = Nothing
            
            ' 일 합계를 표시 합니다.
            Set tRs = oInspect.GetInspectByLotPerMonth(Format(dtpDate(0), "YYYYMMDD"), Format(dtpDate(1), "YYYYMMDD"), nClss, _
                                    nChkCustom, sCustomID, nChkArticle, sArticle, nChkOrder, sOrder)
            If tRs.RecordCount > 0 Then
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 500
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 11) = "합    계"
                .MergeCells = flexMergeFree
                .MergeRow(.Rows - 1) = True

                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 11) = True
                .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 11) = flexAlignCenterCenter
'                .TextMatrix(.Rows - 1, 12) = Format(tRs!StuffQty, "###,###")   ' 투입량(Inspect 테이블)
                .TextMatrix(.Rows - 1, 12) = Format(tRs!CardQty, "###,###")     ' 투입량(Card 테이블)
                .TextMatrix(.Rows - 1, 14) = Format(tRs!PassQty, "###,###")
                .TextMatrix(.Rows - 1, 15) = Format(IIf(nClss = 1, 0, tRs!DyeQty), "###,###")
                .TextMatrix(.Rows - 1, 16) = Format(IIf(nClss = 2, 0, tRs!WeaveQty), "###,###")
                .Cell(flexcpAlignment, .Rows - 1, 15, .Rows - 1, 16) = flexAlignRightCenter
            End If
            tRs.Close
            Set tRs = Nothing
            ' 월 누계를 표시 합니다
            Set tRs = oInspect.GetInspectByLotPerMonth(Format(dtpDate(0), "YYYYMM") & "01", Format(dtpDate(1), "YYYYMM") & "31", nClss, _
                                    nChkCustom, sCustomID, nChkArticle, sArticle, nChkOrder, sOrder)
            If tRs.RecordCount > 0 Then
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 500
                .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 11) = Format(dtpDate(0), "MM") & "월 누계"
                .MergeCells = flexMergeFree
                .MergeRow(.Rows - 1) = True
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 11) = True
                .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 11) = flexAlignCenterCenter
'                .TextMatrix(.Rows - 1, 12) = Format(tRs!StuffQty, "###,###")    ' 투입량(Inspect 테이블)
                .TextMatrix(.Rows - 1, 12) = Format(tRs!CardQty, "###,###")     ' 투입량(Card 테이블)
                .TextMatrix(.Rows - 1, 14) = Format(tRs!PassQty, "###,###")
                .TextMatrix(.Rows - 1, 15) = Format(IIf(nClss = 1, 0, tRs!DyeQty), "###,###")
                .TextMatrix(.Rows - 1, 16) = Format(IIf(nClss = 2, 0, tRs!WeaveQty), "###,###")
                .Cell(flexcpAlignment, .Rows - 1, 15, .Rows - 1, 16) = flexAlignRightCenter

            End If
            tRs.Close
            Set tRs = Nothing
        '----------------------------------------------------------------------------------------------------
            .Cell(flexcpText, 1, 10, 1, .Cols - 1) = "실적일자 : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD")
            .Cell(flexcpAlignment, 1, 10, 1, .Cols - 1) = flexAlignRightCenter
'            .MergeCells = flexMergeFree
'            For iCol = 0 To 10
'                .MergeCol(iCol) = True
'            Next iCol
    
            .Redraw = flexRDDirect
            .SetFocus
        End With
    Else
        grdInspect.Rows = grdInspect.FixedRows
        MsgBox LoadResString(203), vbInformation
    End If
    Set oInspect = Nothing

    m_bloading = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    m_bloading = False
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub optClss_Click(Index As Integer)
    Select Case Index
        Case 0:
            optClss(0).FontBold = True
            optClss(1).FontBold = False
            optClss(2).FontBold = False
        Case 1:
            optClss(0).FontBold = False
            optClss(1).FontBold = True
            optClss(2).FontBold = False
        Case 2:
            optClss(0).FontBold = False
            optClss(1).FontBold = False
            optClss(2).FontBold = True
    End Select
End Sub

Private Sub optOrder_Click(Index As Integer)
    
    If Index = 0 Then
        chkSearch(3).Caption = "Order No."
    Else
        chkSearch(3).Caption = "관리번호"
    End If

End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 4 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 5 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    End If

End Sub
