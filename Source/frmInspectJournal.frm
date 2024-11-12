VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectJournal 
   ClientHeight    =   10710
   ClientLeft      =   -1290
   ClientTop       =   1680
   ClientWidth     =   15420
   Icon            =   "frmInspectJournal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   15420
   WindowState     =   2  '최대화
   Begin VSFlex7LCtl.VSFlexGrid grdInspect 
      Height          =   8925
      Left            =   15
      TabIndex        =   2
      Top             =   720
      Width           =   15255
      _cx             =   26908
      _cy             =   15743
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16051421
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483639
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
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
   Begin VB.Frame fraSearch 
      Height          =   795
      Left            =   0
      TabIndex        =   3
      Top             =   -90
      Width           =   15285
      Begin Threed.SSCommand cmdSearch 
         Height          =   570
         Left            =   3660
         TabIndex        =   10
         Top             =   165
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1005
         _Version        =   196609
         Caption         =   "      검색(&F)"
         PictureAlignment=   1
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   450
         Index           =   0
         Left            =   1380
         TabIndex        =   1
         Top             =   240
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   23789569
         CurrentDate     =   36271
      End
      Begin Threed.SSPanel pnlName 
         Height          =   450
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   794
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "검사일자"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   0
            Top             =   135
            Width           =   1050
         End
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   570
         Left            =   13575
         TabIndex        =   8
         Top             =   165
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1005
         _Version        =   196609
         Caption         =   "      닫기(&X)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   570
         Left            =   11925
         TabIndex        =   9
         Top             =   165
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1005
         _Version        =   196609
         Caption         =   "      인쇄(&P)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Frame fraOrder 
         BorderStyle     =   0  '없음
         Height          =   645
         Left            =   5250
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   3165
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   405
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   105
            Width           =   1200
         End
      End
   End
End
Attribute VB_Name = "frmInspectJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Const REPORTFILE1 = "\Report\Inspect.rpt"
'Private Const REPORTFILE2 = "\Report\InspectByLot.rpt"
'Private Const REPORTFILE3 = "\Report\InspectRollDetail.rpt"
'
'Private Const LIMIT_ROW2 = 6
'Private Const LIMIT_ROW4 = 18
'Private Const LIMIT_WIDTH2 = 1540
'
'Private m_bSortForward As Boolean
'
'Private m_sOperate As String * 1
Private m_bLoading As Boolean

Private Sub cmdPrint_Click()
Dim iRow As Integer

    On Error GoTo Err_Handler

    Screen.MousePointer = vbHourglass

    With grdInspect
        .Redraw = flexRDNone
        .RowHeight(0) = 600
        .RowHeight(1) = 350
        .Cell(flexcpFontSize, 2, 0, .Rows - 1, .Cols - 1) = 7
        For iRow = 2 To .Rows - 1
            If .RowHeight(iRow) <> 0 Then
                .RowHeight(iRow) = 300
            End If
        Next iRow
        .PrintGrid "[ 발행일 : " & Format(Now, "YYYY/MM/DD HH:NN") & " ] ", True, 2, 600, 500
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .Cell(flexcpFontSize, 2, 0, .Rows - 1, .Cols - 1) = 8
        For iRow = 2 To .Rows - 1
            If .RowHeight(iRow) <> 0 Then
                .RowHeight(iRow) = 400
            End If
        Next iRow
        
        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbDefault
    MsgBox "인쇄 되었습니다", vbInformation, "인쇄"

    Exit Sub

Err_Handler:
    MsgBox "발행 오류 : " & Err.Description, vbExclamation, "발행 Error"
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrUnpressed
    
    frmInspectJournal.WindowState = 2
    
    cmdPrint.Picture = LoadResPicture("PRINT", vbResIcon)
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
    chkSearch(0).Value = 1
End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub

Private Sub Form_Load()
    Dim i%

    dtpDate(0) = Now
    Me.Show

    Call InitGrid

    cmdSearch.Picture = LoadResPicture("FIND", vbResIcon)
    chkSearch(0).Value = vbChecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
    PlusMDI.tbrMain.Buttons("Menu").Value = tbrPressed
End Sub

Private Sub InitGrid()
    Dim i%
    Dim iRow As Integer

    With grdInspect
        .FixedRows = 3:     .FixedCols = 0
        .Rows = 3:          .Cols = 19
        .WordWrap = True
        .Redraw = flexRDNone
        .ScrollBars = flexScrollBarVertical

        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 400
        .FontSize = 8
        .Cell(flexcpFontBold, 0, 0, 2, .Cols - 1) = True
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .TextMatrix(2, 0) = "밧쟈" & vbCrLf & "No":     .ColWidth(0) = 500
        .TextMatrix(2, 1) = "거래선":                   .ColWidth(1) = 800:     .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(2, 2) = "오다No":                   .ColWidth(2) = 0:       .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(2, 3) = "관리No":                   .ColWidth(3) = 700:     .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(2, 4) = "품명":                     .ColWidth(4) = 1500:    .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(2, 5) = "구분":                     .ColWidth(5) = 750:     .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(2, 6) = "규격":                     .ColWidth(6) = 550:     .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(2, 7) = "순위":                     .ColWidth(7) = 0:       .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(2, 8) = "색상":                     .ColWidth(8) = 1300:    .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(2, 9) = "수주":                     .ColWidth(9) = 650:     .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(2, 10) = "Lot":                     .ColWidth(10) = 400:    .ColAlignment(10) = flexAlignCenterCenter
        .TextMatrix(2, 11) = "투입":                    .ColWidth(11) = 750:    .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(2, 12) = "검사":                    .ColWidth(12) = 0:      .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(2, 13) = "합격":                    .ColWidth(13) = 750:    .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(2, 14) = "가공 불량":               .ColWidth(14) = 3300:   .ColAlignment(14) = flexAlignLeftCenter
        .TextMatrix(2, 15) = "제직 불량":               .ColWidth(15) = 1650:   .ColAlignment(15) = flexAlignLeftCenter
        .TextMatrix(2, 16) = "난단":                    .ColWidth(16) = 450:    .ColAlignment(16) = flexAlignCenterCenter
        .TextMatrix(2, 17) = "견본":                    .ColWidth(17) = 450:    .ColAlignment(17) = flexAlignCenterCenter
        .TextMatrix(2, 18) = "호기":                    .ColWidth(18) = 450:    .ColAlignment(18) = flexAlignCenterCenter
        
        
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "[[ 검사실적 일보 ]]"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 14
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 0, 1, 14) = " "
        .Cell(flexcpText, 1, 15, 1, .Cols - 1) = "실적일자 : "
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = 10
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = True
        
        .MergeCells = flexMergeFixedOnly
        For iRow = 0 To 1
            .MergeRow(iRow) = True
        Next iRow
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub grdInspect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdInspect.ToolTipText = grdInspect.TextMatrix(grdInspect.Row, grdInspect.Col)
End Sub

Private Sub optOrder_Click(Index As Integer)
'    If Index = 0 Then
'        grdInspect.ColWidth(2) = 700
'        grdInspect.ColWidth(3) = 0
'        chkSearch(2).Caption = "OrderNo"
'    Else
'        grdInspect.ColWidth(2) = 0
'        grdInspect.ColWidth(3) = 700
'        chkSearch(2).Caption = "관리No"
'    End If
End Sub

Private Sub cmdSearch_Click()
'    If Len(txtSearch(1)) = 0 Then chkSearch(1) = vbUnchecked
'    If Len(txtSearch(2)) = 0 Then chkSearch(2) = vbUnchecked
'    If Len(txtSearch(3)) = 0 Then chkSearch(3) = vbUnchecked
    
    Call FillGridOrder
End Sub

Private Sub FillGridOrder()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim tRs      As Recordset
    Dim i%, iNowRow%
    Dim iCnt As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sDefect0 As String
    Dim sDefect1 As String
    Dim sCustom As String

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bLoading = True

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetInspectByLot(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)))

    If rs.RecordCount > 0 Then
        With grdInspect
            .Redraw = flexRDNone
            iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
            .Rows = .FixedRows
            For i = 1 To rs.RecordCount
                If i = 1 Then
                    sCustom = rs!KCustom
                End If
                DoEvents
                
                iCnt = 0:   sDefect0 = "":  sDefect1 = ""
                
                If rs!DefectQty > 0 Or rs!defectqty1 > 0 Then
                    
                    Set tRs = oInspect.GetDefectArray(MakeDate(DF_SHORT, dtpDate(0)), rs!OrderID, rs!LotNo)
                    
                    If tRs.RecordCount > 0 Then
                        If rs!DefectQty > 0 Then
                            sDefect0 = CStr(rs!DefectQty) & "="
                        End If
                        If rs!defectqty1 > 0 Then
                            sDefect1 = CStr(rs!defectqty1) & "="
                        End If
                        
                        For iCnt = 1 To tRs.RecordCount
                            If tRs!KindID = "01" Then
                                sDefect0 = sDefect0 & tRs!KDefect & "(" & tRs!defect0 & ") "
                            Else
                                sDefect1 = sDefect1 & tRs!KDefect & "(" & tRs!defect1 & ") "
                            End If
                            tRs.MoveNext
                        Next iCnt
                    End If
                    Set tRs = Nothing
                End If
                
                If sCustom <> rs!KCustom Then
                    .AddItem " "
                    .RowHeight(.Rows - 1) = 0
                    sCustom = rs!KCustom
                End If
                                 
                .AddItem "" & vbTab & rs!KCustom & vbTab & rs!OrderNo & vbTab & _
                        CStr(CInt(Mid(rs!OrderID, 5, 2))) & "-" & CStr(CInt(Mid(rs!OrderID, 7))) & vbTab & _
                        rs!Article & vbTab & rs!Work & vbTab & rs!Width & vbTab & rs!ColorID & vbTab & _
                        rs!ColorID & "." & rs!Color & vbTab & Format(rs!ColorQty, "###,###") & vbTab & rs!LotNo & vbTab & _
                        Format(rs!StuffQty, "###,###") & vbTab & Format(rs!CtrlQty, "###,###") & vbTab & Format(rs!PassQty, "###,###") & vbTab & _
                        sDefect0 & vbTab & sDefect1 & vbTab & _
                        Format(rs!CutQty, "###,###") & vbTab & Format(rs!SampleQty, "###,###")
    
                .RowHeight(.Rows - 1) = 400
    
                rs.MoveNext
            Next i
            ' 일 합계를 표시 합니다.
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 400
            .TextMatrix(.Rows - 1, 8) = "합    계"
            .Cell(flexcpFontBold, .Rows - 1, 8, .Rows - 1, 8) = True
            .Cell(flexcpAlignment, .Rows - 1, 8, .Rows - 1, 8) = flexAlignCenterCenter
            .TextMatrix(.Rows - 1, 11) = Format(.Aggregate(flexSTSum, .FixedRows, 11, .Rows - 2, 11), "###,###")
            .TextMatrix(.Rows - 1, 13) = Format(.Aggregate(flexSTSum, .FixedRows, 13, .Rows - 2, 13), "###,###")
            ' 월 누계를 표시 합니다
            Set tRs = oInspect.GetInspectByLotPerMonth(MakeDate(DF_SHORT, Format(dtpDate(0), "YYYYMM") & "01"), MakeDate(DF_SHORT, dtpDate(0)))
            If tRs.RecordCount > 0 Then
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 400
                .TextMatrix(.Rows - 1, 8) = "월 누 계"
                .Cell(flexcpFontBold, .Rows - 1, 8, .Rows - 1, 8) = True
                .Cell(flexcpAlignment, .Rows - 1, 8, .Rows - 1, 8) = flexAlignCenterCenter
                .TextMatrix(.Rows - 1, 11) = Format(tRs!StuffQty, "###,###")
                .TextMatrix(.Rows - 1, 13) = Format(tRs!PassQty, "###,###")
            End If
            tRs.Close
            Set tRs = Nothing
        '----------------------------------------------------------------------------------------------------
            rs.Close
            Set rs = Nothing
            .Cell(flexcpText, 1, 15, 1, .Cols - 1) = "실적일자 : " & Format(dtpDate(0), "YYYY/MM/DD")
    
            .Cell(flexcpBackColor, 2, 0, .Rows - 1, 8) = &HF4ECDD
            .MergeCells = flexMergeFree
            For iCol = 0 To 9
                .MergeCol(iCol) = True
            Next iCol
    
            .Redraw = flexRDDirect
            .SetFocus
        End With
    Else
        MsgBox "해당 일자의 데이터가 없읍니다" & vbCr & vbCr & "확인하시고 다시 작업 하세요.", vbInformation
    End If
    Set oInspect = Nothing

    m_bLoading = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    m_bLoading = False
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub chkSearch_Click(Index As Integer)
'    If Index = 0 Then
'        If chkSearch(Index) Then
'            dtpDate(0).Enabled = True
'            dtpDate(0).SetFocus
'        Else
'            dtpDate(0).Enabled = False
'            cmdSearch.SetFocus
'        End If
'    Else
'        If chkSearch(Index) Then
'            If Index = 1 Then cmdFind(0).Enabled = True
'            If Index = 3 Then cmdFind(1).Enabled = True
'            txtSearch(Index).Enabled = True
'
'            txtSearch(Index).SetFocus
'        Else
'            If Index = 1 Then cmdFind(0).Enabled = False
'            If Index = 3 Then cmdFind(1).Enabled = False
'            txtSearch(Index).Enabled = False
'
'            cmdSearch.SetFocus
'        End If
'    End If
End Sub

'Private Sub txtSearch_GotFocus(Index As Integer)
''    Call GotFocusText(txtSearch(Index))
'End Sub
'
'Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    Call MoveFocus(KeyCode)
'End Sub
'
'Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        If Index = 1 Then
'            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
'        ElseIf Index = 3 Then
'            Call ReturnCode(LG_ARTICLE, 0, False, txtSearch(3))
'        End If
'    Else
'        KeyAscii = KeyPress(txtSearch(Index), KeyAscii)
'    End If
'End Sub
'
'Private Sub cmdFind_Click(Index As Integer)
'    If Index = 0 Then
'        Call ReturnCode(LG_CUSTOM, 0, True, txtSearch(1))
'    ElseIf Index = 1 Then
'        Call ReturnCode(LG_ARTICLE, 0, True, txtSearch(3))
'    End If
'End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

