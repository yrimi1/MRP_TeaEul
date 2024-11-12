VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTaxList 
   ClientHeight    =   9255
   ClientLeft      =   765
   ClientTop       =   855
   ClientWidth     =   11850
   Icon            =   "frmTaxList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7485
      Left            =   30
      TabIndex        =   2
      Top             =   930
      Width           =   11775
      _cx             =   20770
      _cy             =   13203
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
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
   Begin Threed.SSFrame fraSearch 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   1614
      _Version        =   196609
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   4800
         TabIndex        =   5
         Top             =   90
         Width           =   1485
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10920
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   0
         Left            =   1380
         TabIndex        =   6
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Format          =   116785152
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   1
         Left            =   1365
         TabIndex        =   7
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Format          =   116785152
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   90
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
            Caption         =   "청구 년월"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   9
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   3270
         TabIndex        =   10
         Top             =   90
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
            TabIndex        =   11
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   6330
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   90
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
         Left            =   2775
         TabIndex        =   14
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   2775
         TabIndex        =   13
         Top             =   150
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8400
      TabIndex        =   3
      Top             =   8460
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
      Left            =   10140
      TabIndex        =   4
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   120
      Top             =   8610
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSCommand cmdTaxPrint 
      Height          =   690
      Left            =   90
      TabIndex        =   15
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "계산서 발행"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmTaxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'
'변경이력
'
'2013.12.12   자체    오승욱   S_201312_태을염직_99   지번주소에서 도로명 주소로 입력가능하게,거래처 주소 도로명 주소 Select
'**************************************************************************************************

Option Explicit

Private m_sPrinter As String

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(Index) Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True

            dtpDate(0).SetFocus
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False

            cmdSearch.SetFocus
        End If
    Else
        If chkSearch(Index) Then
            cmdFind(Index).Enabled = True
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
        Else
            cmdFind(Index).Enabled = False
            txtSearch(Index).Enabled = False
            cmdSearch.SetFocus
        End If
    End If
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
End Sub

Private Sub cmdPrint_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    Call FillGridPrint
End Sub

Private Sub PrintTax(nRow As Integer)
    Dim oCustom As PlusLib2.CCustom
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nFormulas%, nCnt%
    Dim nTQty&, nTAmount&, nTTax&
    Dim sTaxSeq$
    Dim sOrderFlag$, sTaxClss$, sDealClss$
    
    On Error GoTo ErrHandler
    

    '***********************************************************************
    '공급받는자 정보
    '-----------------------------------------------------------------------
        
    Set oCustom = New PlusLib2.CCustom
    oCustom.Connection = g_adoCon
    Set rs = oCustom.GetCustomOne(grdData.TextMatrix(nRow, 11))
    Set oCustom = Nothing
    
    sTaxSeq = Right(grdData.TextMatrix(nRow, 13), 7)
    
    With cryReport
        .Reset
        .PrintFileType = crptText
        .ReportFileName = App.Path & "\Report\Tax.Rpt"
    
        '***************************************************************************
         '공급 받는자 정보 출력
        '---------------------------------------------------------------------------
        .Formulas(0) = "TaxSeq='" & Left(sTaxSeq, 3) & "-" & Right(sTaxSeq, 4) & "'"
        '사업자번호
        .Formulas(1) = "CustomNo= '" & Left(CheckNull(rs!CustomNo), 3) & " - " & Mid(CheckNull(rs!CustomNo), 4, 2) & " - " & Right(CheckNull(rs!CustomNo), 5) & "'"
        .Formulas(2) = "Custom='" & CheckNull(rs!kCustom) & "'"                 '회사명
        .Formulas(3) = "Chief='" & CheckNull(rs!Chief) & "'"                    '대표자
        
''        'S_201312_태을염직_99 에 의한 수정-OLD소스
''        .Formulas(4) = "Address='" & CheckNull(rs!Address1) & " " & CheckNull(rs!Address2) & "'"
        'S_201312_태을염직_99 에 의한 수정-NEW 소스
        If CheckNull(rs!Address1) <> "" Then             '도로명 주소 있으면
            .Formulas(4) = "Address='" & CheckNull(rs!Address1) & " " & CheckNull(rs!Address2) & "'"
        Else                            '도로명 주소 없으면-지번주소
            .Formulas(4) = "Address='" & CheckNull(rs!AddressJiBun1) & " " & CheckNull(rs!AddressJiBun2) & "'"
        End If
        
        .Formulas(5) = "Condition='" & CheckNull(rs!Condition) & "'"            '업태
        .Formulas(6) = "Category='" & CheckNull(rs!Category) & "'"              '종목
        '***************************************************************************
        
        'S_201312_태을염직_99 에 의한 추가-엑셀 하드 코딩 대신 DB에서 가져옴
        '***************************************************************************
        '공급자 정보 출력
        '---------------------------------------------------------------------------
        .ParameterFields(0) = "CustomNo1" & ";" & Format(g_companyInfo.Company_No, "###-##-#####") & ";True"                 '사업자번호
        .ParameterFields(1) = "Custom1" & ";" & g_companyInfo.Company_Name & ";True"                   '상호
        .ParameterFields(2) = "Chief1" & ";" & g_companyInfo.Chief & ";True"                    '대표자
        If CheckNull(g_companyInfo.Address1) <> "" Then              '도로명 주소 있으면
            .ParameterFields(3) = "Address1" & ";" & g_companyInfo.Address1 & " " & g_companyInfo.Address2 & ";True"                  '주소
        Else                            '도로명 주소 없으면-지번주소
            .ParameterFields(3) = "Address1" & ";" & g_companyInfo.AddressJiBun1 & " " & g_companyInfo.AddressJiBun2 & ";True"                  '주소
        End If

        .ParameterFields(4) = "Condition1" & ";" & g_companyInfo.Company_type & ";True"                '업태
        .ParameterFields(5) = "Category1" & ";" & g_companyInfo.Category & ";True"              '종목
        '***************************************************************************
        
    End With
    rs.Close
    Set rs = Nothing
    '***********************************************************************
        
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetTax(grdData.TextMatrix(nRow, 13), grdData.TextMatrix(nRow, 12))
    Set oSubul = Nothing
    
    cryReport.Formulas(7) = "PrnDate='" & rs!PrnDate & "'"
    sOrderFlag = rs!OrderFlag
    sTaxClss = rs!TaxClss
    sDealClss = rs!DealClss
    
    With cryReport
        For i = 0 To rs.RecordCount - 1
            nTQty = nTQty + rs!SumQty
            nTAmount = nTAmount + rs!Amount
            nTTax = nTTax + rs!Tax
            
            rs.MoveNext
        Next i
    End With
    rs.MoveFirst
    
    With cryReport
        nCnt = 0
        nFormulas = 7
        For i = 0 To rs.RecordCount - 1
            If nCnt < 3 Then
                .Formulas(nFormulas + (i * 5) + 1) = "Article" & (i + 1) & "='" & rs!Article & "'"
                .Formulas(nFormulas + (i * 5) + 2) = "WorkName" & (i + 1) & "='" & rs!WorkName & "'"
                .Formulas(nFormulas + (i * 5) + 3) = "SumQty" & (i + 1) & "='" & rs!SumQty & "'"
                .Formulas(nFormulas + (i * 5) + 4) = "Amount" & (i + 1) & "='" & rs!Amount & "'"
                .Formulas(nFormulas + (i * 5) + 5) = "Tax" & (i + 1) & "='" & IIf(rs!Tax = 0, "", rs!Tax) & "'"
            Else
                .Formulas(nFormulas + (i * 5) + 1) = "Article" & (i + 1) & "='외 " & rs.RecordCount - nCnt & "건'"
                Exit For
            End If
            nCnt = nCnt + 1
            rs.MoveNext
        Next i
    End With
    
    With cryReport
        .Formulas(28) = "TAmount='" & nTAmount & "'"
        .Formulas(29) = "TTax='" & nTTax & "'"
        .Formulas(30) = "Space='" & 10 - Len(CStr(nTAmount)) & "'"
        .Formulas(31) = "Total='" & nTAmount + nTTax & "'"
        
        If sOrderFlag = "0" And sTaxClss = "불포함" Then
            If sDealClss = "1" Then
                .Formulas(32) = "Remark='LC/OPEN'"
            ElseIf sDealClss = "2" Then
                .Formulas(32) = "Remark='구매승인서'"
            ElseIf sDealClss = "3" Then
                .Formulas(32) = "Remark='임가공계약서'"
            End If
        Else
            .Formulas(32) = "Remark=''"
        End If
        
        .SelectionFormula = ""
        .PrinterDriver = m_sPrinter
        .PrinterName = m_sPrinter
       .PrinterPort = "LPT1:"
        .WindowState = crptMaximized
'        If bPreview Then
'           .Destination = crptToWindow
'        Else
            .Destination = crptToPrinter
'        End If
            .CopiesToPrinter = 2
        .Action = 1
    End With
    Exit Sub
    
ErrHandler:
    Set oCustom = Nothing
    Set oSubul = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmTaxLst.PrintTax", Err.Description)
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub cmdTaxPrint_Click()
    Dim i%
    Dim sPrinter As String
        
    sPrinter = Printer.DeviceName
        
    If frmPrinter.SelectPrinter(sPrinter, m_sPrinter) Then
        With grdData
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then
                    Call PrintTax(i)
                End If
            Next i
        End With
        Call ReturnPrinter(sPrinter)
    End If
End Sub

Private Sub Form_Load()

    Me.Move 0, 0, 11970, 9660

    Call SetOperate(Me)
    
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).Enabled = False
    txtSearch(1).Enabled = False
   
    dtpDate(0) = Now
    dtpDate(1) = Now
    
    Call InitGrid
End Sub

Private Sub InitGrid()
    Dim i%

    With grdData
        .Redraw = flexRDNone
        .Cols = 14
        Call SetVSFlexGrid(grdData)

        .Rows = 4
        .FixedCols = 0
        .FixedRows = 4

        .RowHeightMin = 350
        .RowHeight(3) = 400

        .TextMatrix(3, 0) = " ":          .ColWidth(0) = 0
        .TextMatrix(3, 1) = "":           .ColWidth(1) = 300:        .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "거래처":     .ColWidth(2) = 2400:       .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(3, 3) = "청구년월":   .ColWidth(3) = 1200:       .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "계산서번호": .ColWidth(4) = 1300:       .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "청구량":     .ColWidth(5) = 1200:       .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(3, 6) = "청구금액":   .ColWidth(6) = 1500:       .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(3, 7) = "부가세":     .ColWidth(7) = 1500:       .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(3, 8) = "합계":       .ColWidth(8) = 1500:       .ColAlignment(8) = flexAlignRightCenter
        
        .TextMatrix(3, 11) = "CustomID":
        .TextMatrix(3, 12) = "청구년월":
        .TextMatrix(3, 13) = "TaxSeq"
        
        .ColFormat(5) = "#,###"
        .ColFormat(6) = "#,###"
        .ColFormat(7) = "#,###"
        .ColFormat(8) = "#,###"
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
                
        For i = 9 To 13
            .ColWidth(i) = 0
        Next i
        
        .MergeCells = flexMergeFixedOnly
        For i = 0 To 13
            .MergeCol(i) = True
        Next i
        
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub FillGridData()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset
    Dim i%, JJ%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetTaxList(IIf(chkSearch(0), 1, 0), Left(MakeDate(DF_SHORT, dtpDate(0)), 6), Left(MakeDate(DF_SHORT, dtpDate(1)), 6), _
                        IIf(chkSearch(1), 1, 0), txtSearch(1).Tag)
    Set oSubul = Nothing

    With grdData
        .Redraw = flexRDNone

        .Rows = .FixedRows
        JJ = 1
        For i = 1 To rs.RecordCount
            If rs!Depth = "1" Then
                .AddItem CStr(JJ)
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = rs!kCustom
                .TextMatrix(.Rows - 1, 3) = Left(rs!BasisYearMon, 4) & "년 " & Right(rs!BasisYearMon, 2) & "월"
                .TextMatrix(.Rows - 1, 4) = Left(rs!TaxSeq, 2) & "-" & Mid(rs!TaxSeq, 3, 2) & "-" & Right(rs!TaxSeq, 4)
                .TextMatrix(.Rows - 1, 5) = rs!SumQty
                .TextMatrix(.Rows - 1, 6) = rs!Amount
                .TextMatrix(.Rows - 1, 7) = rs!Tax
                .TextMatrix(.Rows - 1, 8) = rs!TotalPrice
                .TextMatrix(.Rows - 1, 10) = rs!Depth
                .TextMatrix(.Rows - 1, 11) = rs!CustomID
                .TextMatrix(.Rows - 1, 12) = rs!PrnDate
                .TextMatrix(.Rows - 1, 13) = rs!TaxSeq
                
                .Cell(flexcpChecked, .Rows - 1, 1, .Rows - 1, 1) = flexUnchecked
                JJ = JJ + 1
            Else
                .AddItem ""
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = rs!kCustom
                .TextMatrix(.Rows - 1, 5) = rs!SumQty
                .TextMatrix(.Rows - 1, 6) = rs!Amount
                .TextMatrix(.Rows - 1, 7) = rs!Tax
                .TextMatrix(.Rows - 1, 8) = rs!TotalPrice
                .TextMatrix(.Rows - 1, 10) = rs!Depth
                
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE9E9E9
            End If
            
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        .Redraw = flexRDDirect
        .SetFocus
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oSubul = Nothing

    Call ErrorBox(Err.Number, "frmTaxList.FillGridData", Err.Description)
End Sub


Private Sub grdData_Click()
    With grdData
        If .Row < .FixedRows Or .Col <> 1 Then Exit Sub
        
        If .TextMatrix(.Row, 10) = 2 Then Exit Sub
        
        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, 1) = flexChecked
        End If
    End With
End Sub


Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    End If
End Sub

Sub FillGridPrint()
    Dim i%
    Dim sDate As String, eDate As String
    
    If chkSearch(0).Value Then
        sDate = Format(dtpDate(0), "YYYY년 MM월")
        eDate = Format(dtpDate(1), "YYYY년 MM월")
    Else
        sDate = ""
        eDate = ""
    End If
    
    With grdData
        .Redraw = flexRDNone

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHeight(2) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        .RowHeight(2) = 350
        .ExtendLastCol = False
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i

        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "세금계산서 현황"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 1, 4, 1, 8) = "▶ 청구년월 : " & sDate & " ~ " & eDate
'        .Cell(flexcpText, 1, .Cols - 4, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD")
        
        .ColHidden(1) = True
'        .ColWidth(2) = 1900
'        .ColWidth(5) = 1100
'        .ColWidth(6) = 1400
'        .ColWidth(7) = 1400
'        .ColWidth(8) = 1400
        
        Call SetPrintMode(grdData, 1, True)
        
        .PrintGrid "태을염직", True, 1, 0, 500
        
        Call SetPrintMode(grdData, 1, False)
        
        .ExtendLastCol = True
        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .ColHidden(1) = False

'        .FontSize = 9
'        .ColWidth(2) = 2000
'        .ColWidth(5) = 1200
'        .ColWidth(6) = 1500
'        .ColWidth(7) = 1500
'        .ColWidth(8) = 1500


        .Redraw = flexRDDirect
    End With
End Sub

