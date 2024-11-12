VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDyeAuxSubul 
   Caption         =   "염조제 수불현황"
   ClientHeight    =   9270
   ClientLeft      =   2055
   ClientTop       =   2820
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.ComboBox cboDyeAux 
      Height          =   300
      Left            =   5340
      Style           =   2  '드롭다운 목록
      TabIndex        =   14
      Top             =   30
      Width           =   1485
   End
   Begin VB.TextBox txtDyeAux 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8280
      TabIndex        =   10
      Top             =   30
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   10980
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   5
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   90
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   90
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   2100
      TabIndex        =   2
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116785153
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   2100
      TabIndex        =   3
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116785153
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   750
      TabIndex        =   4
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입출일자"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Value           =   1  '확인
         Width           =   1095
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   6
      Top             =   8490
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
      Left            =   30
      TabIndex        =   7
      Top             =   690
      Width           =   11730
      _cx             =   20690
      _cy             =   13705
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8460
      TabIndex        =   8
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   11
      Left            =   7050
      TabIndex        =   11
      Top             =   30
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "염조제"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   0
      Left            =   10260
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   30
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   315
      Left            =   3990
      TabIndex        =   15
      Top             =   30
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   196609
      Caption         =   "염조제구분"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   0
      Left            =   3390
      TabIndex        =   17
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   1
      Left            =   3390
      TabIndex        =   16
      Top             =   450
      Width           =   360
   End
End
Attribute VB_Name = "frmDyeAuxSubul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index
        Case 0     '입고일자 Term
            If chkSearch(Index) = vbChecked Then
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            Else
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            End If
        Case 2    '품명
            If chkSearch(Index) = vbChecked Then
                txtDyeAux.Enabled = True
                txtDyeAux.SetFocus
                cmdFind(0).Enabled = True
            Else
                txtDyeAux.Enabled = False
                txtDyeAux.Tag = ""
                cmdSearch.SetFocus
                cmdFind(0).Enabled = False
            End If
    End Select
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
            Call ReturnCode(IIf(cboDyeAux.ListIndex = 0, LG_DYE, LG_AUX), , False, txtDyeAux)
    End Select
End Sub

Private Sub cmdPrint_Click()
    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
        Call ColResize("-")
        Call FillGrdPrint
        Call ColResize("+")
    End If
End Sub
Sub ColResize(ByVal pType As String)
    Dim II%, JJ As Integer
    
    If pType = "-" Then
        JJ = -1
    Else
        JJ = 1
    End If
    
    With grdData
        For II = 0 To .Cols - 1
            If pType = "-" Then
                .ColWidth(II) = .ColWidth(II) * 0.8
                
            Else
                .ColWidth(II) = .ColWidth(II) / 0.8
            End If
        Next II
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
        
        
'        .TextMatrix(3, 1) = "일자":             .ColWidth(1) = 600:                 .ColAlignment(1) = flexAlignCenterCenter
'        .TextMatrix(3, 2) = "거래처명":         .ColWidth(2) = 1800:                .ColAlignment(2) = flexAlignLeftCenter
'        .TextMatrix(3, 3) = "실 입고처":        .ColWidth(3) = 1200:                .ColAlignment(3) = flexAlignLeftCenter
'        .TextMatrix(3, 4) = "품명":             .ColWidth(4) = 2400:                .ColAlignment(4) = flexAlignLeftCenter
'        .TextMatrix(3, 5) = "관리번호":         .ColWidth(5) = 1300:                .ColAlignment(5) = flexAlignCenterCenter
'        .TextMatrix(3, 6) = "OrderNO":          .ColWidth(6) = 1300:                .ColAlignment(6) = flexAlignLeftCenter
'        .TextMatrix(3, 7) = "가공":             .ColWidth(7) = 1000:                 .ColAlignment(7) = flexAlignCenterCenter
'        .TextMatrix(3, 8) = "절 수":            .ColWidth(8) = 800:                 .ColAlignment(8) = flexAlignRightCenter
'        .TextMatrix(3, 9) = "수   량":          .ColWidth(9) = 900:                 .ColAlignment(9) = flexAlignRightCenter
        
        .Redraw = flexRDDirect
    End With

End Sub
Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub


Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String
    
    If chkSearch(0).Value Then
        sDate = Format(dtpDate(0), "YYYY/MM/DD")
        eDate = Format(dtpDate(1), "YYYY/MM/DD")
    Else
        sDate = ""
        eDate = ""
    End If
    
    With grdData
        .Redraw = flexRDBuffered

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
'        For i = 0 To 3
'           .MergeRow(i) = True
'        Next i

        .FontSize = 7
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "염조제 수불현황"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, 2) = "▶ 입출일자 : " & sDate & " ~ " & eDate
'        .Cell(flexcpText, 1, .Cols - 4, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        
        .Cell(flexcpText, 1, 3, 1, 4) = "▶ 염조제구분 : " & cboDyeAux.Text
        .Cell(flexcpText, 1, 6, 1, .Cols - 2) = "▶ 염조제명   : " & IIf(chkSearch(2).Value, txtDyeAux.Text, "(전체)")
        .Cell(flexcpText, 2, 7, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .PrintGrid "태을염직", True, 1, 100, 500

'        .GridLinesFixed = flexGridNone
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True

        .FontSize = 9
        
        For i = 0 To 3
           .MergeRow(i) = False
        Next i
        
        .GridLinesFixed = flexGridInset

        .Redraw = flexRDDirect
    End With
    
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
    Me.Move 0, 0, 11970, 9660

    Call InitGrid
    Call SetOperate(Me)
    '----- 날짜설정
    dtpDate(0) = Now
    dtpDate(1) = Now
    
'    CboStuffClss2.ListIndex = 0
    
    '----- 검색용 입고구분 설정
    With cboDyeAux
        .Clear
        .AddItem "염료"
        .AddItem "조제"
    End With
    cboDyeAux.ListIndex = 0
    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdFind(0).Enabled = False

End Sub

Private Sub InitGrid()
    Dim i%
    
    Call SetVSFlexGrid(grdData)
    With grdData
        .Cols = 11
        .Rows = 4
        .FixedRows = 4
        .FixedCols = 1
        
        .RowHeightMin = 300
        .RowHeight(3) = 400
        
        .TextMatrix(3, 0) = ""
        .TextMatrix(3, 1) = "일자":             .ColWidth(1) = 1000:                .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "염조제명":         .ColWidth(2) = 3000:                .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(3, 3) = "입고수량":         .ColWidth(3) = 1100:                .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(3, 4) = "입고금액":         .ColWidth(4) = 1200:                .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(3, 5) = "":                 .ColWidth(5) = 10:                  .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "사용수량":         .ColWidth(6) = 1100:                .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(3, 7) = "사용금액":         .ColWidth(7) = 1200:                .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(3, 8) = "":                 .ColWidth(8) = 10:                  .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(3, 9) = "재고수량":         .ColWidth(9) = 1100:                .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "재고금액":        .ColWidth(10) = 1590:               .ColAlignment(10) = flexAlignRightCenter
        
        
'        .ColHidden(8) = True
'        .ColHidden(9) = True
'        .ColHidden(10) = True

        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .MergeCells = flexMergeFree
        
        
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With

End Sub


Sub FillgrdData()
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%, II%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim nStockCnt As Integer, nStockPrice As Long
    

  '  On Error GoTo ErrHandler

    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon
    oDyeAux.UserName = g_sUserName
    
    If chkSearch(0).Value Then
        sDate = MakeDate(DF_SHORT, dtpDate(0))
        eDate = MakeDate(DF_SHORT, dtpDate(1))
    Else
        sDate = ""
        eDate = ""
    End If
    
    Set rs = oDyeAux.GetDyeAuxSubulsDraft(sDate, eDate, IIf(chkSearch(2) = vbChecked, 1, 0), txtDyeAux.Tag)

    Set oDyeAux = Nothing
    
    With grdData
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        If rs.RecordCount = 1 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
        
            Do Until rs.EOF
                If Left(Trim(rs!Depth), 1) = "Z" Then
                    dDate_str = ""
                Else
                    dDate_str = MakeDate(DF_MD, rs!DyeAuxDate)
                End If
                
                .AddItem "" & vbTab & dDate_str & vbTab & Trim(rs!DyeAux) & vbTab & SetCurrency(rs!InQty, 2) & vbTab & SetCurrency(rs!InPrice, 0) & vbTab & "" & vbTab & _
                            SetCurrency(rs!OutQty, 2) & vbTab & SetCurrency(rs!OutPrice, 0) & "" & vbTab & 0 & vbTab & 0
                rs.MoveNext
            Loop
        End If
        
        nStockCnt = 0: nStockPrice = 0
        For II = .FixedRows To .Rows - 1
            If II = .FixedRows Then
                nStockCnt = .ValueMatrix(II, 3) - .ValueMatrix(II, 6)
                nStockPrice = .ValueMatrix(II, 4) - .ValueMatrix(II, 7)
            Else
                nStockCnt = .ValueMatrix(II - 1, 9) + .ValueMatrix(II, 3) - .ValueMatrix(II, 6)
                nStockPrice = .ValueMatrix(II - 1, 10) + .ValueMatrix(II, 4) - .ValueMatrix(II, 7)
            End If
            
            .TextMatrix(II, 9) = SetCurrency(nStockCnt, 0)
            .TextMatrix(II, 10) = SetCurrency(nStockPrice, 0)
        Next II
        .Redraw = flexRDDirect
    End With
    
    
    
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffINList.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oDyeAux = Nothing
    
End Sub

