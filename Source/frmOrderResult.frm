VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOrderResult 
   ClientHeight    =   9255
   ClientLeft      =   2565
   ClientTop       =   1635
   ClientWidth     =   11850
   Icon            =   "frmOrderResult.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   780
      Left            =   10980
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   15
      ToolTipText     =   "자료 저장"
      Top             =   45
      Width           =   780
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2145
      MousePointer    =   99  '사용자 정의
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   435
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1500
      MousePointer    =   99  '사용자 정의
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   435
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2145
      MousePointer    =   99  '사용자 정의
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   75
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1500
      MousePointer    =   99  '사용자 정의
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   7230
      TabIndex        =   10
      Top             =   75
      Width           =   1980
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   2
      Left            =   7230
      TabIndex        =   9
      Top             =   465
      Width           =   1980
   End
   Begin VB.Frame fraOrder 
      Height          =   765
      Left            =   60
      TabIndex        =   6
      Top             =   -15
      Width           =   1305
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   210
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   1155
      End
   End
   Begin Threed.SSCommand cmdHTML 
      Height          =   690
      Left            =   8445
      TabIndex        =   5
      Top             =   8535
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      HTML(&H)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   6735
      TabIndex        =   4
      Top             =   8535
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   360
      TabIndex        =   1
      Top             =   3645
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
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10155
      TabIndex        =   0
      Top             =   8535
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   4065
      TabIndex        =   16
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   54460417
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   4065
      TabIndex        =   17
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   54460417
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2820
      TabIndex        =   18
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "수주일자"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   19
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   5970
      TabIndex        =   20
      Top             =   75
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   21
         Top             =   45
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   5970
      TabIndex        =   22
      Top             =   465
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품  명"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   23
         Top             =   45
         Width           =   1095
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   9255
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   75
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   2
      Left            =   9255
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   450
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin MRPPlus2.WizFlexGroup grdData 
      Height          =   7620
      Left            =   15
      TabIndex        =   28
      Top             =   855
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   13441
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   4
      Left            =   5385
      TabIndex        =   27
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   5
      Left            =   5385
      TabIndex        =   26
      Top             =   525
      Width           =   360
   End
End
Attribute VB_Name = "frmOrderResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExcel_Click()
    If grdData.FlexGrid.Rows = 1 Then
        MsgBox LoadResString(111), vbInformation
        cmdSearch.SetFocus

        Exit Sub
    End If
    
    Call MakeExcelGrid(grdData.FlexGrid)
End Sub

Private Sub cmdHTML_Click()
    If grdData.FlexGrid.Rows = 1 Then
        MsgBox LoadResString(111), vbInformation
        Exit Sub
    End If
    If MakeHtmlGrid(grdData.FlexGrid, "C:\" & Me.Caption & ".html") Then
        Call RelateOpen(Me.hwnd, "C:\" & Me.Caption & ".html")
    End If

End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11975, 9660
    
    Call InitGrid
    Call SetOperate(Me)
    Call SetDtpDate(1, dtpDate(0), dtpDate(1))
    
    chkSearch(0).Value = vbChecked
    
    Show

    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
    
    cmdFind(1).Enabled = False
    cmdFind(2).Enabled = False
   
    With cmdFind(1)
        .MousePointer = ssCustom
        .MouseIcon = LoadResPicture("POINTER", vbResCursor)
        .Picture = LoadResPicture("FIND", vbResIcon)
    End With
    
    With cmdFind(2)
        .MousePointer = ssCustom
        .MouseIcon = LoadResPicture("POINTER", vbResCursor)
        .Picture = LoadResPicture("FIND", vbResIcon)
    End With
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then '[1] 수주일자
        If chkSearch(Index).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else '[2,3] 거래처, 품명
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            cmdFind(Index).Enabled = True
        Else
            txtSearch(Index).Enabled = False
            cmdFind(Index).Enabled = False
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, True, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , True, txtSearch(2))
    End If
End Sub

Private Sub cmdSearch_Click()
    Call InitGrid
    Call FillGrid
    

End Sub

Private Sub FillGrid()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    
    Dim i%, iRowCount%, iNowRow%

    On Error GoTo ErrHandler

    proProgress.Value = 0
    lblCount = LoadResString(304)

    pnlProgress.Visible = True

    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon

    Set rs = oOrder.GetbasisOrder(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                IIf(chkSearch(1), 1, 0), txtSearch(1).Tag, IIf(chkSearch(2), 1, 0), txtSearch(2).Tag)
    Set oOrder = Nothing
    
    With grdData.FlexGrid
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        iRowCount = rs.RecordCount
        For i = 1 To iRowCount
            DoEvents
            lblCount = CStr(i) & " / " & CStr(iRowCount)
            proProgress.Value = CInt((i / iRowCount) * 100)
        
            .AddItem CStr(i) & vbTab & IIf(IsNull(rs!CloseDate), "", "■") & vbTab & rs!KCustom & vbTab & _
                MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & Format(MakeDate(DF_LONG, rs!OrderDate), "YYYY") & vbTab & _
                Format(MakeDate(DF_LONG, rs!OrderDate), "MM") & vbTab & Format(MakeDate(DF_LONG, rs!OrderDate), "DD") & vbTab & _
                rs!Article & vbTab & rs!Work & vbTab & MakeRating(rs!FlexRate, rs!LossRate) & vbTab & rs!Width2 & vbTab & _
                CStr(rs!OrderQty) & vbTab & IIf(rs!OrderUnit = "0", "YDS", "MTS") & vbTab & CStr(rs!UnitCost) & vbTab & IIf(rs!Unit = 0, "\", "$")

            rs.MoveNext
        Next i
        rs.Close

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = iNowRow
            
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If

        .Redraw = flexRDDirect
    End With
    Set rs = Nothing
    
    pnlProgress.Visible = False
    Exit Sub
ErrHandler:
    Set rs = Nothing
    Set oOrder = Nothing
    
    Call ErrorBox(Err.Number, "OrderTotal.FillGrid", Err.Description)
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))
    
End Sub

Private Sub InitGrid()
    Dim i%

    grdData.InitGroup
    With grdData
        With .FlexGrid
            .Redraw = flexRDNone
            .Rows = 1
            .Cols = 16
            
            .ScrollTrack = True
            
            .TextArray(0) = "":                 .ColWidth(0) = 450:     .ColAlignment(0) = flexAlignCenterCenter
            .TextArray(1) = "완료":             .ColWidth(1) = 450:     .ColAlignment(1) = flexAlignCenterCenter
            .TextArray(2) = "거래처":           .ColWidth(2) = 1650:    .ColAlignment(2) = flexAlignLeftCenter
            .TextArray(3) = "관리 번호":        .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignCenterCenter
            .TextArray(4) = "Order No.":        .ColWidth(4) = 1300:    .ColAlignment(4) = flexAlignLeftCenter
            .TextArray(5) = "년":               .ColWidth(5) = 630:     .ColAlignment(5) = flexAlignCenterCenter
            .TextArray(6) = "월":               .ColWidth(6) = 310:     .ColAlignment(6) = flexAlignCenterCenter
            .TextArray(7) = "일":               .ColWidth(7) = 310:     .ColAlignment(7) = flexAlignCenterCenter
            .TextArray(8) = "품명":             .ColWidth(8) = 1450:    .ColAlignment(8) = flexAlignLeftCenter
            .TextArray(9) = "가공구분":         .ColWidth(9) = 800:     .ColAlignment(9) = flexAlignLeftCenter
            .TextArray(10) = "축+Loss":         .ColWidth(10) = 800:    .ColAlignment(10) = flexAlignCenterCenter
            .TextArray(11) = "원단 폭":         .ColWidth(11) = 750:    .ColAlignment(11) = flexAlignCenterCenter
            .TextArray(12) = "수주량":          .ColWidth(12) = 1000:   .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "단위":            .ColWidth(13) = 550:    .ColAlignment(13) = flexAlignCenterCenter
            .TextArray(14) = "단가":            .ColWidth(14) = 700:    .ColAlignment(14) = flexAlignRightCenter
            .TextArray(15) = "단위":            .ColWidth(15) = 0
                                    
            .ColFormat(12) = "#,##0"
            .ColFormat(14) = "#,##0"
                
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
            Next i
            .Redraw = flexRDDirect
        End With

        .ColLock(12) = True
        .ColTotal(12) = True
        .Update
    End With

End Sub

Private Sub optOrder_Click(Index As Integer)
    Dim i%
    
    With grdData.FlexGrid
        For i = .FixedCols To .Cols - 1
            If .TextArray(i) = optOrder(Index).Caption Then
                .ColWidth(i) = 1300
            End If
            
            If .TextArray(i) = optOrder(Abs(Index - 1)).Caption Then
                .ColWidth(i) = 0
            End If
        Next i
    End With
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
        End If
        cmdSearch.SetFocus
    End If
End Sub
