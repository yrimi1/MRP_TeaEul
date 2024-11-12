VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderAcptView 
   Caption         =   "일자별 Order접수 명세서"
   ClientHeight    =   9270
   ClientLeft      =   1740
   ClientTop       =   1155
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.ComboBox CboOrderFlag 
      Height          =   300
      Left            =   9480
      Style           =   2  '드롭다운 목록
      TabIndex        =   22
      Top             =   360
      Width           =   1395
   End
   Begin VB.TextBox txtExchRate 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   7380
      TabIndex        =   21
      Top             =   360
      Width           =   1635
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   285
      Left            =   6090
      TabIndex        =   20
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   196609
      Caption         =   "환율"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   11040
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   19
      ToolTipText     =   "자료 저장"
      Top             =   0
      Width           =   780
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   30
      MousePointer    =   99  '사용자 정의
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   690
      MousePointer    =   99  '사용자 정의
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   4200
      TabIndex        =   5
      Top             =   30
      Width           =   1455
   End
   Begin VB.TextBox txtArticle 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7380
      TabIndex        =   4
      Top             =   30
      Width           =   1635
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   0
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
      Height          =   7800
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   11820
      _cx             =   20849
      _cy             =   13758
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
      Left            =   8490
      TabIndex        =   2
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   8
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
      Left            =   1350
      TabIndex        =   9
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
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "수주일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   0
      Left            =   5670
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   360
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
      Left            =   2910
      TabIndex        =   12
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   11
      Left            =   6090
      TabIndex        =   14
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   2
      Left            =   9090
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   2910
      TabIndex        =   17
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
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
         Index           =   3
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Width           =   1185
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   9465
      TabIndex        =   23
      Top             =   30
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "사용구분"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "frmOrderAcptView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub chkSearch_Click(Index As Integer)
    Dim dChk_bol As Boolean
    dChk_bol = chkSearch(Index).Value
    
    Select Case Index
'        Case 0
'            dtpDate(0).Enabled = dChk_bol
'            dtpDate(1).Enabled = dChk_bol
'
'            If dChk_bol = True Then
'                dtpDate(0).SetFocus
'            End If
            
        Case 1
            txtCustom(1).Enabled = dChk_bol
            cmdFind(0).Enabled = dChk_bol
            
            If dChk_bol = True Then
                txtCustom(1).SetFocus
            Else
                txtCustom(1).Text = ""
                txtCustom(1).Tag = ""
            End If
            
            
        Case 2
            txtArticle.Enabled = dChk_bol
            cmdFind(2).Enabled = dChk_bol
            
            If dChk_bol = True Then
                txtArticle.SetFocus
            Else
                txtArticle.Text = ""
                txtArticle.Tag = ""
            End If
            
        Case 3
            txtSearch(3).Enabled = dChk_bol
            
            If dChk_bol = True Then
                txtSearch(3).SetFocus
            Else
                txtSearch(3).Text = ""
                txtSearch(3).Tag = ""
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
            Call ReturnCode(LG_CUSTOM, , False, txtCustom(1))
        Case 2                '[3] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
    End Select

End Sub

Private Sub cmdPrint_Click()
    Dim vColWidth()
    
    ReDim vColWidth(grdData.Cols - 1)
    
    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
        Call ColResize_ColWidth(grdData, ES_REDUCE, 10, vColWidth)
        Call FillGrdPrint
        Call ColResize_ColWidth(grdData, ES_EXPAND, 10, vColWidth)
    End If
End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub


Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String
    
    With grdData
        .Redraw = flexRDBuffered
   '     .ExtendLastCol = False

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .FontSize = 7
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "일자별 ORDER접수 명세서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, 3) = "▶ 접수일자 : " & MakeDate(DF_FULL, dtpDate(0)) & "~" & MakeDate(DF_FULL, dtpDate(1))
        .Cell(flexcpText, 1, 4, 1, 6) = "▶ 관리번호 : " & IIf(chkSearch(3).Value = vbChecked, Trim(txtSearch(3)), AllStr)
        .Cell(flexcpText, 1, 7, 1, .Cols - 1) = "▶ 거 래 처 : " & IIf(chkSearch(1).Value = vbChecked, Trim(txtCustom(1)), AllStr)
        
        .Cell(flexcpText, 2, 1, 2, 6) = "▶ 품    명 : " & IIf(chkSearch(2).Value = vbChecked, Trim(txtArticle), AllStr)
        .Cell(flexcpText, 2, 8, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpBackColor, 0, 0, .FixedRows - 2, .Cols - 1) = vbWhite
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbWhite
        
        
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignLeftCenter
        
        .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignLeftCenter
        
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
        .PrintGrid "태을염직", True, 1, 600, 500

        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True

        For i = .FixedRows To .Rows - 1
            Call SetGrdColor(grdData, .TextMatrix(i, 10), i, 0, i, .Cols - 1)
        Next i
        .FontSize = 9
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    
End Sub




Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[3] 금일
        dtpDate(0) = Date - 1
        dtpDate(1) = Date - 1
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If

End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11970, 9660

    Call InitGrid
    Call SetOperate(Me)
    
    With CboOrderFlag
        .AddItem "1.전체"
        .AddItem "2.LOCAL"
        .AddItem "3.내수"
        .AddItem "4.시가공"
        .ListIndex = 0
    End With
    
    '----- 날짜설정
    dtpDate(0) = Now
    dtpDate(1) = Now
    
    txtCustom(1).Enabled = chkSearch(1).Value
    cmdFind(0).Enabled = chkSearch(1).Value
    
    txtArticle.Enabled = chkSearch(2).Value
    cmdFind(2).Enabled = chkSearch(2).Value
    
    txtSearch(3).Enabled = chkSearch(3).Value
'    dtpDate(0).Enabled = chkSearch(0).Value
'    dtpDate(1).Enabled = chkSearch(0).Value
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).MouseIcon = LoadResPicture("POINTER", vbResCursor)

    
    
    Call FillgrdData
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
        .TextMatrix(3, 1) = "거래처명":         .ColWidth(1) = 1800:                .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(3, 2) = "관리번호":         .ColWidth(2) = 1300:                .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "OrderNO":          .ColWidth(3) = 1500:                .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "품명":             .ColWidth(4) = 1600:                .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(3, 5) = "가공구분":         .ColWidth(5) = 1000:                .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(3, 6) = "C수":              .ColWidth(6) = 600:                 .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(3, 7) = "Order량":          .ColWidth(7) = 1200:                .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(3, 8) = "단가":             .ColWidth(8) = 900:                 .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(3, 9) = "금액":             .ColWidth(9) = 1400:                .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "depth":             .ColWidth(10) = 0:                 .ColAlignment(10) = flexAlignRightCenter
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColHidden(0) = True
        
        .MergeCells = flexMergeFree
        .ScrollBars = flexScrollBarBoth
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .Redraw = flexRDDirect
    End With

End Sub


Sub FillgrdData()
    Dim oCls As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String, dOrderQty_Str$
    Dim sDate As String, eDate As String, sFromDate As String, sToDate As String
    Dim nChkOrder As Integer, sOrderID As String, nChkCustom As Integer, sCustomID As String, nChkArticle As Integer, sArticleID As String
    Dim nColorVal As Long, ExchRate As Single
    
    On Error GoTo ErrHandler
  
    nChkOrder = 0: sOrderID = "": nChkCustom = 0: sCustomID = "": nChkArticle = 0: sArticleID = ""
    
    sDate = MakeDate(DF_SHORT, dtpDate(0))    '-- 수주일자 시작일자
    eDate = MakeDate(DF_SHORT, dtpDate(1))    '-- 수주일자 끝일자
    
    '-- 시작일이 1일이 아닌경우 해당월의 1일부터 - 시작일 전일 까지 이월로 처리 하기위해
    If Right(sDate, 2) <> "01" Then
        sFromDate = Left(sDate, 6) & "01"
        sToDate = MakeDate(DF_SHORT, DateAdd("D", -1, dtpDate(0)))
    Else
        sFromDate = sDate
        sToDate = ""
    End If
    
    If chkSearch(3).Value = vbChecked Then
        nChkOrder = 1
        sOrderID = Trim(txtSearch(3).Text)
    End If
    
    If chkSearch(1).Value = vbChecked Then
        nChkCustom = 1
        sCustomID = Trim(txtCustom(1).Tag)
    End If
    
    If chkSearch(2).Value = vbChecked Then
        nChkArticle = 1
        sArticleID = Trim(txtArticle.Tag)
    End If
    
    ExchRate = val(txtExchRate)
    
    Set oCls = New PlusLib2.COrder
    oCls.Connection = g_adoCon
    oCls.UserName = g_sUserName
    
    Set rs = oCls.GetOrderAcptView(sDate, eDate, sFromDate, sToDate _
                                , nChkOrder, sOrderID _
                                , nChkCustom, sCustomID _
                                , nChkArticle, sArticleID, ExchRate, CboOrderFlag.ListIndex)

    Set oCls = Nothing
    
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
                Select Case rs!UnitClss
                Case "Y"
                    dOrderQty_Str$ = SetCurrency(rs!OrderQty, 0) & Space(4)
                Case "M"
                    dOrderQty_Str$ = SetCurrency(rs!OrderQty, 0) & " M"
                
                Case Else
                    dOrderQty_Str$ = SetCurrency(rs!OrderQty, 0) & Space(4)
                End Select
                
                .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & IIf(Trim(rs!OrderID) <> "", MakeOrderID(rs!OrderID, OM_EXPAND), "") & vbTab & _
                            Trim(rs!OrderNo) & vbTab & Trim(rs!Article) & vbTab & _
                            Trim(rs!WorkName) & vbTab & rs!ColorCnt & vbTab & dOrderQty_Str$ & vbTab & _
                            SetCurrency(rs!UnitPrice, 2) & vbTab & SetCurrency(rs!Price, 0) & vbTab & rs!Depth
                            
                If (rs!Depth <> "2" And rs!Depth <> "1") Then
                    Call SetGrdColor(grdData, rs!Depth, .Rows - 1, 0, .Rows - 1, .Cols - 1)
                End If
                rs.MoveNext
            Loop
        End If
        .ScrollBars = flexScrollBarBoth
        
        .MergeCells = flexMergeFree
        .MergeCol(1) = True
        
'        For i = 1 To 3
'            .MergeCol(i) = True
'        Next i
        
        .Redraw = flexRDDirect
    End With
    
    
    
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmOrderAcptView.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oCls = Nothing
    
End Sub

Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(2)
    End If

End Sub



Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(0)
    
    End If

End Sub
