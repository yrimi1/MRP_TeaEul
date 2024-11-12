VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSubulAdjust 
   Caption         =   "수불마감처리"
   ClientHeight    =   9270
   ClientLeft      =   1185
   ClientTop       =   2910
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15180
   Begin VB.TextBox txtOrderID 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   11070
      TabIndex        =   13
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "전체 선택"
      Height          =   315
      Index           =   0
      Left            =   30
      TabIndex        =   12
      Top             =   8880
      Width           =   1140
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "선택 해제"
      Height          =   315
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   8520
      Width           =   1140
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Left            =   7350
      TabIndex        =   8
      Top             =   30
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   12450
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
      Left            =   1740
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
      Left            =   1740
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   3810
      TabIndex        =   2
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   68616193
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   3810
      TabIndex        =   3
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   68616193
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   2460
      TabIndex        =   4
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "수불일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13530
      TabIndex        =   6
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
      Height          =   6420
      Left            =   0
      TabIndex        =   7
      Top             =   690
      Width           =   15150
      _cx             =   26723
      _cy             =   11324
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
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   0
      Left            =   8820
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
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
      Left            =   6060
      TabIndex        =   10
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "거래처"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   9780
      TabIndex        =   14
      Top             =   30
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
         Caption         =   "관리번호"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   30
      TabIndex        =   16
      Top             =   30
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1138
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   1140
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   11790
      TabIndex        =   19
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "정산처리(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTotal 
      Height          =   1410
      Left            =   0
      TabIndex        =   20
      Top             =   7110
      Width           =   15150
      _cx             =   26723
      _cy             =   2487
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
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   1
      Left            =   5100
      TabIndex        =   22
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   0
      Left            =   5100
      TabIndex        =   21
      Top             =   60
      Width           =   360
   End
End
Attribute VB_Name = "frmSubulAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index
       Case 1    '관리번호
            If chkSearch(1) Then
                txtOrderID.Enabled = True
                txtOrderID.SetFocus
            Else
                txtOrderID.Enabled = False
                txtOrderID.Text = ""
            End If
    End Select
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Call SetGridToggleChecked(grdData, Index, 1)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0                '[1] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustom)
        Case 2                '[3] 품명 코드
            Call ReturnCode(LG_ORDER, , False, txtOrderID)
    End Select
End Sub




Private Sub cmdSave_Click()
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim TSubulClose() As PlusLib2.TSubulClose
    Dim nPkey As Variant, sDate As String, eDate As String
    Dim II%, nCheckNon%

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    
    If grdData.Rows = grdData.FixedRows Then
        Exit Sub
    End If
    
    ReDim TSubulClose(grdData.Rows - grdData.FixedRows - 1)
    
    With grdData
        For II = .FixedRows To .Rows - 1
            nPkey = Split(.TextMatrix(II, .Cols - 1), "-")
            TSubulClose(II - .FixedRows).sSubulDate = nPkey(0)
            TSubulClose(II - .FixedRows).sSubulClss = nPkey(1)
            TSubulClose(II - .FixedRows).sIOClss = nPkey(2)
            TSubulClose(II - .FixedRows).nSeq = val(nPkey(3))
            TSubulClose(II - .FixedRows).sCloseClss = IIf(.Cell(flexcpChecked, II, .Col) = flexChecked, "*", "")
        Next II
    End With
    
    If oStuffIn.DoSubulAdjust(TSubulClose) Then
        Call InitgrdTotal
        Call FillgrdData
    Else
        MsgBox ("정산처리중 오류가 발생했습니다.")
    End If
    
End Sub

Private Sub cmdSearch_Click()
    Call InitgrdTotal
    Call FillgrdData
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
    PlusMDI.pnlMenu.Visible = False
    
    Me.Move 0, 0, 15300, 9660

    Call InitGrid
    Call SetOperate(Me)
    
    '----- 날짜설정
    dtpDate(0) = Now
    dtpDate(1) = Now
    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    txtOrderID.Enabled = False

End Sub

Private Sub InitGrid()
    Dim i%
    
    Call SetVSFlexGrid(grdData)
    With grdData
        .Cols = 15
        .Rows = 4
        .FixedRows = 4
        .FixedCols = 1
        
        .RowHeightMin = 300
        .RowHeight(3) = 400
        
        .TextMatrix(3, 0) = ""
        .TextMatrix(3, 1) = "선택":             .ColWidth(1) = 600:                 .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "수불일자":         .ColWidth(2) = 1000:                 .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "OrderNO":          .ColWidth(3) = 1500:                .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "OrderID":          .ColWidth(4) = 1200:                .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(3, 5) = "입출" & vbCrLf & "구분":         .ColWidth(5) = 700:                .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "입출고처":         .ColWidth(6) = 1200:                .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "품명":             .ColWidth(7) = 2800:                .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(3, 8) = "색상":             .ColWidth(8) = 2400:                .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "절수":             .ColWidth(9) = 800:                .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "수량":            .ColWidth(10) = 900:               .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(3, 11) = "Loss":            .ColWidth(11) = 800:                .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(3, 12) = "축율":            .ColWidth(12) = 800:               .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(3, 13) = "출고량":          .ColWidth(13) = 1500:               .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(3, 14) = "Pkey":            .ColWidth(14) = 0
        
        .MergeCells = flexMergeFree
        
        .ColDataType(1) = flexDTBoolean
        .SelectionMode = flexSelectionByRow
        
        .ExtendLastCol = False
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
    End With
    
    Call SetToggle
    
    
    Call SetVSFlexGrid(grdTotal)
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Rows = 5
        .FixedRows = 2
        .Cols = 7
        .ExtendLastCol = True
        
        .RowHeight(0) = 275
        
        .TextArray(0) = "":                       .ColWidth(0) = 710
        .TextArray(1) = "입고":                   .ColWidth(1) = 2000
        .TextArray(2) = "입고":                   .ColWidth(2) = 2000
        .TextArray(3) = "구          분":         .ColWidth(3) = 4000
        .TextArray(4) = "출고":                   .ColWidth(4) = 2000
        .TextArray(5) = "출고":                   .ColWidth(5) = 2000
        .TextArray(6) = "출고":                   .ColWidth(6) = 2000
        
        .TextArray(.Cols + 0) = "":               .ColWidth(0) = 710
        .TextArray(.Cols + 1) = "절수":           .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignRightCenter
        .TextArray(.Cols + 2) = "수량":           .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignRightCenter
        .TextArray(.Cols + 3) = "구          분": .ColWidth(3) = 4000: .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(.Cols + 4) = "절수":           .ColWidth(4) = 2000: .ColAlignment(4) = flexAlignRightCenter
        .TextArray(.Cols + 5) = "수량":           .ColWidth(5) = 2000: .ColAlignment(5) = flexAlignRightCenter
        .TextArray(.Cols + 6) = "실출고량":       .ColWidth(6) = 2000: .ColAlignment(6) = flexAlignRightCenter
        
        .TextMatrix(2, 3) = "정          산"
        .TextMatrix(3, 3) = "미    정    산"
        .TextMatrix(4, 3) = "합          계"
        
        Dim II%
        
        For II = 0 To .Rows - 1
            .MergeRow(II) = True
        Next II
        
        For II = 0 To .Cols - 1
            .MergeCol(II) = True
        Next II
        
        .MergeCells = flexMergeFixedOnly
        .Redraw = flexRDDirect
    End With

End Sub


Sub FillgrdData()
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    
    If txtCustom.Tag = "" Or Trim(txtCustom) = "" Then
        MsgBox ("거래처를 반드시 입력 하십시오")
        Exit Sub
    End If
    
    Set rs = oStuffIn.GetSubulAdjust(sDate, eDate, 1, txtCustom.Tag _
                                , IIf(chkSearch(1) = vbChecked, 1, 0), txtOrderID.Tag)

    Set oStuffIn = Nothing
    
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
                .AddItem "" & vbTab & "" & vbTab & MakeDate(DF_LONG, rs!SubulDate) & vbTab & Trim(rs!OrderNo) & vbTab & _
                         MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                         Trim(rs!IOType) & vbTab & Trim(rs!Custom) & vbTab & Trim(rs!Article) & vbTab & Trim(rs!Color) & vbTab & _
                         SetCurrency(rs!cnt, 0) & vbTab & SetCurrency(rs!Qty, 0) & vbTab & SetCurrency(rs!LossRate, 2) & vbTab & _
                         SetCurrency(rs!ChunkRate, 2) & vbTab & SetCurrency(rs!OutRealQty, 2) & vbTab & Trim(rs!Pkey)
                
                .Cell(flexcpChecked, .Rows - 1, 1) = IIf(rs!CloseClss = "*", flexChecked, flexUnchecked)
                
                ' Total
                Select Case rs!IOclss
                    Case "1", "3"
                        If rs!CloseClss = "*" Then
                            grdTotal.TextMatrix(2, 1) = SetCurrency(grdTotal.ValueMatrix(2, 1) + rs!cnt, 0)
                            grdTotal.TextMatrix(2, 2) = SetCurrency(grdTotal.ValueMatrix(2, 2) + rs!Qty, 0)
                        Else
                            grdTotal.TextMatrix(3, 1) = SetCurrency(grdTotal.ValueMatrix(3, 1) + rs!cnt, 0)
                            grdTotal.TextMatrix(3, 2) = SetCurrency(grdTotal.ValueMatrix(3, 2) + rs!Qty, 0)
                        End If
                        grdTotal.TextMatrix(4, 1) = SetCurrency(grdTotal.ValueMatrix(4, 1) + rs!cnt, 0)
                        grdTotal.TextMatrix(4, 2) = SetCurrency(grdTotal.ValueMatrix(4, 2) + rs!Qty, 0)
                    Case "5", "7"
                        If rs!CloseClss = "*" Then
                            grdTotal.TextMatrix(2, 4) = SetCurrency(grdTotal.TextMatrix(2, 4) + rs!cnt, 0)
                            grdTotal.TextMatrix(2, 5) = SetCurrency(grdTotal.TextMatrix(2, 5) + rs!Qty, 0)
                            grdTotal.TextMatrix(2, 6) = SetCurrency(grdTotal.TextMatrix(2, 6) + rs!OutRealQty, 0)
                        Else
                            grdTotal.TextMatrix(3, 4) = SetCurrency(grdTotal.TextMatrix(3, 4) + rs!cnt, 0)
                            grdTotal.TextMatrix(3, 5) = SetCurrency(grdTotal.TextMatrix(3, 5) + rs!Qty, 0)
                            grdTotal.TextMatrix(3, 6) = SetCurrency(grdTotal.TextMatrix(3, 6) + rs!OutRealQty, 0)
                        End If
                        
                        grdTotal.TextMatrix(4, 4) = SetCurrency(grdTotal.TextMatrix(4, 4) + rs!cnt, 0)
                        grdTotal.TextMatrix(4, 5) = SetCurrency(grdTotal.TextMatrix(4, 5) + rs!Qty, 0)
                        grdTotal.TextMatrix(4, 6) = SetCurrency(grdTotal.TextMatrix(4, 6) + rs!OutRealQty, 0)
                End Select
                
                rs.MoveNext
            Loop
        End If
        Call SetToggle
        
        .Redraw = flexRDDirect
    End With
    
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmSubulAdjust.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
End Sub
Sub InitgrdTotal()
    Dim II%
    With grdTotal
        For II = 2 To 4
            .TextMatrix(II, 1) = 0
            .TextMatrix(II, 2) = 0
            .TextMatrix(II, 4) = 0
            .TextMatrix(II, 5) = 0
            .TextMatrix(II, 6) = 0
        Next II
        .Redraw = flexRDDirect
    End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub grdData_Click()
    Dim Checked As Boolean
    
    With grdData
        If .MouseRow < .FixedRows Then Exit Sub
        
        If .Col = 1 Then
            Checked = IIf(.Cell(flexcpChecked, .Row, .Col) = flexChecked, False, True)  '체크되면 true, 체크해제는 false
            .Cell(flexcpChecked, .Row, .Col) = Checked
        End If
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    chkSearch(1).Caption = optOrder(Index).Caption
    Call SetToggle
End Sub
Sub SetToggle()
    Dim Index As Integer
    
    If optOrder(0).Value Then
        Index = 0
    Else
        Index = 1
    End If
    
    Select Case Index
        Case 0: grdData.ColWidth(3) = 1500
                grdData.ColWidth(4) = 0
                
        Case 1: grdData.ColWidth(3) = 0
                grdData.ColWidth(4) = 1200
                
    End Select
End Sub

Private Sub txtCustom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_CUSTOM, , False, txtCustom)
        Call MoveFocus(KeyAscii)
    End If

End Sub
