VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutWareView 
   Caption         =   "제품출고현황"
   ClientHeight    =   9270
   ClientLeft      =   1605
   ClientTop       =   4005
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15195
   Begin VB.ComboBox cboTaxClss 
      Height          =   300
      Left            =   7890
      Style           =   2  '드롭다운 목록
      TabIndex        =   17
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   11040
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   16
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   7890
      TabIndex        =   7
      Top             =   30
      Width           =   1455
   End
   Begin VB.TextBox txtArticle 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3930
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      Top             =   30
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116785153
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   6570
      TabIndex        =   1
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   196609
      Caption         =   "사용구분"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13530
      TabIndex        =   2
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
      TabIndex        =   3
      Top             =   690
      Width           =   15120
      _cx             =   26670
      _cy             =   13705
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
      Left            =   11850
      TabIndex        =   4
      Top             =   8490
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   0
      Left            =   9330
      TabIndex        =   8
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
      Left            =   6570
      TabIndex        =   9
      Top             =   30
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
         TabIndex        =   10
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   11
      Left            =   2670
      TabIndex        =   11
      Top             =   360
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
         TabIndex        =   12
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   13
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
      Index           =   0
      Left            =   2670
      TabIndex        =   14
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
         TabIndex        =   15
         Top             =   60
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   315
      Left            =   30
      TabIndex        =   18
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   196609
      Caption         =   "출고일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   19
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   116785153
      CurrentDate     =   36871
   End
End
Attribute VB_Name = "frmOutWareView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum eColor
    CL_DEPTH1 = &HFFFFC0
    CL_DEPTH2 = &HFFFF80
    CL_DEPTH3 = &HFFFF00
    CL_DEPTH4 = &HB4C729
End Enum

Private Sub chkSearch_Click(Index As Integer)
    Dim dChk_bol As Boolean
    dChk_bol = chkSearch(Index).Value
    
    Select Case Index
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

'
'Private Sub chkSearch_Click(Index As Integer)
'    Select Case Index
'        Case 0     '입고일자 Term
'            If chkSearch(Index) = vbChecked Then
'                dtpDate(0).Enabled = True
'                dtpDate(1).Enabled = True
'            Else
'                dtpDate(0).Enabled = False
'                dtpDate(1).Enabled = False
'            End If
'        Case 1    '거래처
'            If chkSearch(Index) = vbChecked Then
'                txtCustom(1).Enabled = True
'                txtCustom(1).SetFocus
'                cmdFind(0).Enabled = True
'            Else
'                txtCustom(1).Enabled = False
'                cmdFind(0).Enabled = False
'                txtCustom(1).Tag = ""
'            End If
'        Case 2    '품명
'            If chkSearch(Index) = vbChecked Then
'                txtArticle.Enabled = True
'                txtArticle.SetFocus
'                cmdFind(2).Enabled = True
'            Else
'                txtArticle.Enabled = False
'                txtArticle.Tag = ""
'                cmdSearch.SetFocus
'                cmdFind(2).Enabled = False
'            End If
'       Case 3    '관리번호
'            If chkSearch(Index) Then
'                txtSearch(3).Enabled = True
'                txtSearch(3).SetFocus
'            Else
'                txtSearch(3).Enabled = False
'                txtSearch(3).Text = ""
'            End If
'
'        Case 4     '입고구분
'            If chkSearch(Index) = vbChecked Then
'                CboStuffClss2.Enabled = True
'            Else
'                CboStuffClss2.Enabled = False
'            End If
'        Case 5     '확정구분
'            If chkSearch(Index) = vbChecked Then
'                cboOrderID.Enabled = True
'            Else
'                cboOrderID.Enabled = False
'            End If
'
'    End Select
'End Sub

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
    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
     '   Call ColResize(grdData, ES_REDUCE, 10)
        Call FillGrdPrint
     '   Call ColResize(grdData, ES_EXPAND, 10)
    End If
End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub


Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String
    Dim nRowHeight As Integer
    Dim nBackColor As Long
    Dim nPageHV As Integer

    
    With grdData
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "제품출고현황"
        .RowHeight(0) = 1000
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "▶ 출고일자 : " & MakeDate(DF_FULL, dtpDate(0)) & " ~ " & MakeDate(DF_FULL, dtpDate(1))
        .Cell(flexcpText, 2, 1, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpBackColor, 0, 0, .FixedRows - 2, .Cols - 1) = vbWhite
        .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignRightCenter
        
        Call SetPrintMode(grdData, 1, True, nPageHV)
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
        .ColHidden(4) = True
        
        For i = .FixedRows To .Rows - 1
            .RowHeight(i) = 400
            ' 일계, 총계의 금액은 BackColor을 설정 한다.
            If (.TextMatrix(i, 11) = "Z4" Or .TextMatrix(i, 11) = "Z5") And .ValueMatrix(i, 10) <> 0 Then
                .Cell(flexcpBackColor, i, 6, i, .Cols - 1) = PRNHeaderColor
            End If
        Next i
        
        .ExtendLastCol = True
        
        .PrintGrid "태을염직", True, 2, 100, 500
        
 '----  인쇄하기 이전으로 원상복귀
        Call SetPrintMode(grdData, 1, False, nPageHV)

        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColHidden(4) = False

        .ExtendLastCol = True
        
        For i = .FixedRows To .Rows - 1
             Call SetGrdColor(grdData, Mid(.TextMatrix(i, 12), 2), i, 0, i, .Cols - 1)
        Next i
        .Redraw = flexRDDirect
        
    End With
    
End Sub







Private Sub Form_Load()
    
    PlusMDI.pnlMenu.Visible = False
    Me.Move 0, 0, 15300, 9660

    Call InitGrid
    Call SetOperate(Me)
    
    '----- 날짜설정
    dtpDate(0) = Now
    dtpDate(1) = Now
    
    txtCustom(1).Enabled = chkSearch(1).Value
    cmdFind(0).Enabled = chkSearch(1).Value
    
    txtArticle.Enabled = chkSearch(2).Value
    cmdFind(2).Enabled = chkSearch(2).Value
    
    txtSearch(3).Enabled = chkSearch(3).Value
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    
    With cboTaxClss
        .AddItem "9.전체"
        .AddItem "0.비사용"
        .AddItem "1.사용"
        .ListIndex = 0
    End With
    
    Call FillgrdData
    
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
        .TextMatrix(3, 1) = "거래처명":         .ColWidth(1) = 2000:                .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(3, 2) = "출고일자":         .ColWidth(2) = 1200:                .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "품    명":         .ColWidth(3) = 2600:                .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "관리번호":         .ColWidth(4) = 900:                 .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "OrderNO":          .ColWidth(5) = 1800:                .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(3, 6) = "가공구분":         .ColWidth(6) = 1400:                .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "절 수":            .ColWidth(7) = 900:                 .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(3, 8) = "출고량":           .ColWidth(8) = 1300:                .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(3, 9) = " ":                .ColWidth(9) = 0:                   .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "단 가":           .ColWidth(10) = 900:                .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(3, 11) = "금  액":          .ColWidth(11) = 1400:               .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(3, 12) = "Depth":           .ColWidth(12) = 0
        .TextMatrix(3, 13) = "nRec":            .ColWidth(13) = 0
        .TextMatrix(3, 14) = "sDefine":         .ColWidth(14) = 0
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .ColHidden(0) = True
        
        .MergeCells = flexMergeRestrictColumns
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
    Dim oCls As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim II%, nCheckCnt%, nItemCnt%, sDepth As String
    Dim dDate_str As String, dOutQty_Str As String
    Dim sDate As String, eDate As String, sFromDate As String, sToDate As String
    Dim nChkOrder As Integer, sOrderID As String, nChkCustom As Integer, sCustomID As String, nChkArticle As Integer, sArticleID As String
    Dim dUnitPrice As String, dPrice As String
    Dim sDefine As String, nCol As Integer, nToCol As Integer
    
    On Error GoTo ErrHandler
'
'    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6) & "01"
'    eDate = MakeDate(DF_SHORT, dtpDate(0))
    
    nChkOrder = 0: sOrderID = "": nChkCustom = 0: sCustomID = "": nChkArticle = 0: sArticleID = ""
    
    sDate = MakeDate(DF_SHORT, dtpDate(0))    '-- 수주일자 시작일자
    eDate = MakeDate(DF_SHORT, dtpDate(1))    '-- 수주일자 끝일자
    
    sFromDate = Left(sDate, 6) & "01"
    sToDate = MakeDate(DF_SHORT, DateAdd("D", -1, dtpDate(0)))
    
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

    Set oCls = New PlusLib2.COutWare
    oCls.Connection = g_adoCon
    oCls.UserName = g_sUserName
    
    Set rs = oCls.GetOutWareView(sDate, eDate, nChkOrder, sOrderID, nChkCustom, sCustomID, nChkArticle, sArticleID, Left(cboTaxClss, 1))

    Set oCls = Nothing
    
    With grdData
        .Rows = .FixedRows
        .Redraw = flexRDDirect
        II = 0
        
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
                    dOutQty_Str = SetCurrency(rs!OutQty, 0) & Space(2)
                Case "M"
                    dOutQty_Str = SetCurrency(rs!OutQty, 0) & Space(1) & "M"
                Case Else
                    dOutQty_Str = SetCurrency(rs!OutQty, 0) & Space(2)
                End Select
                
'                If rs!UnitPrice = 0 Then
'                    Select Case rs!Depth
'                        Case "Z0"
'                            Select Case rs!OutClss
'                                Case "3", "4", "5": dUnitPrice = ""
'                                Case Else:          dUnitPrice = "미확정"
'                            End Select
'                        Case "Z4", "Z5"
'                            Select Case rs!UnitPriceClss
'                                Case 1, 3:        dUnitPrice = ""
'                                Case 2:           dUnitPrice = "미확정"
'                            End Select
'                        Case Else: dUnitPrice = ""
'
'                    End Select
'                Else
'                End If
                
                If rs!UnitPrice = 0 Then
                    dUnitPrice = ""
                
                Else
                
                    dUnitPrice = SetCurrency(rs!UnitPrice, 0)
                End If
                
                If rs!Price = 0 Then
                    dPrice = ""
                Else
'                    If rs!Depth = "Z0" Then
'                        dPrice = "\ " & Space(13 - Len(SetCurrency(rs!Price, 0, 1))) & SetCurrency(rs!Price, 0, 1)
'                    Else
                        dPrice = "\ " & SetCurrency(rs!Price, 0, 1)
'                    End If
                End If
                
                '--- 업체별 제품 개수에 의해 제품계, 일계 란 Hidden처리 루틴 시작 ------------
                If rs!nRec <> 1 Then
                
                    If Trim(rs!kCustom) <> Trim(.TextMatrix(.Rows - 1, 1)) Then
                        .AddItem ""
                        .RowHidden(.Rows - 1) = True
                    ElseIf rs!OutDate <> MakeDate(DF_SHORT, .TextMatrix(.Rows - 1, 2)) Then
                            .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & MakeDate(DF_LONG, rs!OutDate)
                            .RowHidden(.Rows - 1) = True
                    End If
                                    
                    .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & MakeDate(DF_LONG, rs!OutDate) & vbTab & Trim(rs!Article) & vbTab & _
                                IIf(rs!Depth = "Z0" Or rs!Depth = "Z1", MakeOrderID(rs!OrderID, OM_COMPACT), rs!OrderID) & vbTab & _
                                Trim(rs!OrderNo) & vbTab & _
                                Trim(rs!WorkName) & vbTab & SetCurrency(rs!OutRoll, 0) & vbTab & _
                                dOutQty_Str & vbTab & "" & vbTab & dUnitPrice & vbTab & dPrice & vbTab & Trim(rs!Depth) & vbTab & rs!nRec & vbTab & rs!sDefine
                
                    If Trim(rs!Depth) <> "Z0" Then
                        sDefine = rs!sDefine
                        Select Case Trim(rs!Depth)
                            '--- 제품계
                            Case "Z1":                                nCol = 5: nToCol = 5
                                .TextMatrix(.Rows - 1, 4) = ""
                            '--- 일자계
                            Case "Z2":                                nCol = 3: nToCol = 3
                            '--- 월계
                            Case "Z3":                                nCol = 3: nToCol = 3
                                .TextMatrix(.Rows - 1, 2) = ""
                            '--- 업체가 바뀌면 초기화
                            Case "Z4":                                nCol = 2: nToCol = 2
                            Case "Z5":                                nCol = 1: nToCol = 1
                                .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 5) = ""
                        End Select
                        
                        .Cell(flexcpText, .Rows - 1, nCol, .Rows - 1, nToCol) = sDefine
                        .Cell(flexcpFontBold, .Rows - 1, nCol, .Rows - 1, .Cols - 1) = True
                        .Cell(flexcpAlignment, .Rows - 1, nCol, .Rows - 1, nToCol) = flexAlignCenterCenter
                        
                        Call SetGrdColor(grdData, Mid(.TextMatrix(.Rows - 1, 12), 2), .Rows - 1, nCol, .Rows - 1, .Cols - 1)
                    
                    End If
                
                End If
                
                rs.MoveNext
            Loop
        End If
        
        
        '--- 업체별 제품 개수에 의해 제품계, 일계 란 Hidden처리 루틴 끝 ------------
        .ScrollBars = flexScrollBarBoth
        
        .MergeCells = flexMergeFree
        

        
        For II = 1 To 6
            .MergeCol(II) = True
        Next II
        
        For II = .FixedRows To .Rows - 1
            .MergeRow(II) = True
        Next II
        
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


Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

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
