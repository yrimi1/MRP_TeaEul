VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffINList 
   Caption         =   "생지입고명세서"
   ClientHeight    =   9270
   ClientLeft      =   1215
   ClientTop       =   1050
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.ComboBox CboOrderFlag 
      Height          =   300
      Left            =   10095
      Style           =   2  '드롭다운 목록
      TabIndex        =   27
      Top             =   330
      Width           =   915
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   6540
      TabIndex        =   15
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtArticle 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8400
      TabIndex        =   14
      Top             =   330
      Width           =   1275
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   360
      Width           =   1155
   End
   Begin VB.ComboBox cboOrderID 
      Height          =   300
      Left            =   6540
      Style           =   2  '드롭다운 목록
      TabIndex        =   12
      Top             =   30
      Width           =   1485
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   11040
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   6
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin VB.ComboBox CboStuffClss2 
      Height          =   300
      Left            =   3960
      Style           =   2  '드롭다운 목록
      TabIndex        =   5
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   720
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
      Left            =   30
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1350
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
      Left            =   1350
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
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입고 일자"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Value           =   1  '확인
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   2670
      TabIndex        =   7
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입고구분"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   26
         Top             =   90
         Width           =   1125
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   8
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
      Left            =   0
      TabIndex        =   9
      Top             =   690
      Width           =   11820
      _cx             =   20849
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
      Left            =   8490
      TabIndex        =   10
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
      Left            =   8010
      TabIndex        =   16
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
      Left            =   5220
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   11
      Left            =   8400
      TabIndex        =   19
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
         TabIndex        =   20
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   2
      Left            =   9750
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   330
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
      TabIndex        =   22
      Top             =   360
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
         Index           =   3
         Left            =   60
         TabIndex        =   23
         Top             =   60
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   315
      Left            =   5220
      TabIndex        =   24
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "확정구분"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   25
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   10080
      TabIndex        =   28
      Top             =   0
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "사용구분"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "frmStuffINList"
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
        Case 1    '거래처
            If chkSearch(Index) = vbChecked Then
                txtCustom(1).Enabled = True
                txtCustom(1).SetFocus
                cmdFind(0).Enabled = True
            Else
                txtCustom(1).Enabled = False
                cmdFind(0).Enabled = False
                txtCustom(1).Tag = ""
            End If
        Case 2    '품명
            If chkSearch(Index) = vbChecked Then
                txtArticle.Enabled = True
                txtArticle.SetFocus
                cmdFind(2).Enabled = True
            Else
                txtArticle.Enabled = False
                txtArticle.Tag = ""
                cmdSearch.SetFocus
                cmdFind(2).Enabled = False
            End If
       Case 3    '관리번호
            If chkSearch(Index) Then
                txtSearch(3).Enabled = True
                txtSearch(3).SetFocus
            Else
                txtSearch(3).Enabled = False
                txtSearch(3).Text = ""
            End If
            
        Case 4     '입고구분
            If chkSearch(Index) = vbChecked Then
                CboStuffClss2.Enabled = True
            Else
                CboStuffClss2.Enabled = False
            End If
        Case 5     '확정구분
            If chkSearch(Index) = vbChecked Then
                cboOrderID.Enabled = True
            Else
                cboOrderID.Enabled = False
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
    'If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
     '   Call ColResize(grdData, ES_REDUCE, 30)
        Call FillGrdPrint
     '   Call ColResize(grdData, ES_EXPAND, 30)
    'End If
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
        

        Call SetPrintMode(grdData, 1, True)
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "생지입고명세서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .RowHeight(0) = 1000
        
        .Cell(flexcpText, 1, 1, 1, 3) = "▶ 기간 : " & sDate & " ~ " & eDate
        
        
'        .Cell(flexcpText, 1, 4, 1, 4) = "▶ 입고 : " & IIf(chkSearch(4).Value, CboStuffClss2.Text, "(전체)")
'        .Cell(flexcpText, 1, 6, 1, 7) = "▶ 확정 : " & IIf(chkSearch(5).Value, cboOrderID.Text, "(전체)")
'        .Cell(flexcpText, 2, 8, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpBackColor, 0, 0, .FixedRows - 2, .Cols - 1) = vbWhite
        .Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1, .Cols - 1) = vbWhite
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
        .ColHidden(6) = True
        .ColHidden(0) = True
        
        .ExtendLastCol = False
        
'        For i = .FixedRows To .Rows - 1
'            ' 일계, 총계의 금액은 BackColor을 설정 한다.
'            If (.TextMatrix(i, 10) = "Z2" Or .TextMatrix(i, 10) = "Z3" Or .TextMatrix(i, 10) = "Z4") Then
'                .Cell(flexcpBackColor, i, 8, i, 9) = PRNHeaderColor
'            End If
'        Next i
        
        .PrintGrid "태을염직", True, 1, 500, 500
        
        
        Call SetPrintMode(grdData, 1, False)

        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, 10)
                Case "Z2"
                    .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                    .Cell(flexcpFontBold, i, 1, i, .Cols - 1) = True
                    
                Case "Z3"
                    .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                    .Cell(flexcpFontBold, i, 1, i, .Cols - 1) = True
                
                Case "Z4"
                    .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                    .Cell(flexcpFontBold, i, 1, i, .Cols - 1) = True
            End Select
            
'            Call SetGrdColor(grdData, Mid(.TextMatrix(i, 10), 2), i, 1, i, .Cols - 1)
            
        Next i
        
        .ColHidden(6) = False

        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    
End Sub
Sub FillGridReSET()
    Dim II As Integer
    
    With grdData
            
        For II = .FixedRows To .Rows - 1
            Select Case .TextMatrix(II, 10)
                Case "Z1"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE9E9E9
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "Z2"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE5E5E5
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "Z3"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "Z4"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0C0C0
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            End Select
        Next II
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
    With CboStuffClss2
        .Clear
        .AddItem "1.생지"
        .ItemData(0) = 1
        .AddItem "3.반품 생지"
        .ItemData(1) = 3
        .ListIndex = 0
    End With
    
    '----- 확정구분
    With cboOrderID
        .Clear
        .AddItem "수주확정"
        .AddItem "수주미확정"
        .ListIndex = 0
    End With
    
    With CboOrderFlag
        .AddItem "9.전체"
        .AddItem "0.비사용"
        .AddItem "1.사용"
        .ListIndex = 0
    End With
    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).MouseIcon = LoadResPicture("POINTER", vbResCursor)

    cmdFind(0).Enabled = False
    cmdFind(2).Enabled = False
    
    txtCustom(1).Enabled = False
    txtArticle.Enabled = False
    txtSearch(3).Enabled = False
    CboStuffClss2.Enabled = False
    cboOrderID.Enabled = False
    

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
        
        .TextMatrix(3, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(3, 1) = "일자":             .ColWidth(1) = 600:                 .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "거래처명":         .ColWidth(2) = 1600:                .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(3, 3) = "실 입고처":        .ColWidth(3) = 1200:                .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "품명":             .ColWidth(4) = 2400:                .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(3, 5) = "관리번호":         .ColWidth(5) = 600:                 .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "OrderNO":          .ColWidth(6) = 1300:                .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "가공":             .ColWidth(7) = 1000:                .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(3, 8) = "절 수":            .ColWidth(8) = 1100:                .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(3, 9) = "수   량":          .ColWidth(9) = 1450:                .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "depth":           .ColWidth(10) = 0:                  .ColAlignment(10) = flexAlignRightCenter
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .MergeCells = flexMergeFree
        
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With

End Sub


Sub FillgrdData()
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    If chkSearch(0).Value Then
        sDate = MakeDate(DF_SHORT, dtpDate(0))
        eDate = MakeDate(DF_SHORT, dtpDate(1))
    Else
        sDate = ""
        eDate = ""
    End If
    
    ' 확정구분
    If chkSearch(5).Value Then
        nCheckNon = cboOrderID.ListIndex + 1
    Else
        nCheckNon = 0  '전체
    End If
    
    If chkSearch(4).Value Then
        StuffClss = CStr(CboStuffClss2.ItemData(CboStuffClss2.ListIndex))
    Else
        StuffClss = ""
    End If
    
    Set rs = oStuffIn.GetStuffINList(IIf(chkSearch(0).Value = 1, 1, 0), sDate, eDate _
                                , IIf(chkSearch(1) = vbChecked, 1, 0) _
                                , txtCustom(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0) _
                                , txtArticle.Tag _
                                , IIf(chkSearch(4) = vbChecked, 1, 0) _
                                , StuffClss _
                                , IIf(chkSearch(3) = vbChecked, 1, 0) _
                                , txtSearch(3).Text _
                                , nCheckNon, Left(CboOrderFlag, 1))

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
                If Trim(rs!Depth) = "Z3" Or Trim(rs!Depth) = "Z4" Then
                    dDate_str = ""
                Else
                    dDate_str = MakeDate(DF_MD, rs!StuffDate)
                End If
                
                If .Rows > .FixedRows And Trim(rs!kCustom) <> .TextMatrix(.Rows - 1, 2) Then
                    .AddItem "" & vbTab & dDate_str
                    .RowHidden(.Rows - 1) = True
                End If
                
                .AddItem "" & vbTab & dDate_str & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Custom) & vbTab & Trim(rs!Article) & vbTab & _
                            IIf(Trim(rs!OrderID) = "", "", Right(MakeOrderID(rs!OrderID, OM_EXPAND), 4)) & vbTab & Trim(rs!OrderNo) & vbTab & _
                            Trim(rs!WorkName) & vbTab & Format$(CheckNum(rs!TotRoll), "###,##0") & vbTab & Format$(CheckNum(rs!TotQty), "##,###,##0") & vbTab & rs!Depth
                rs.MoveNext
            Loop
        End If
            
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, 10)
                Case "Z2"
                    .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                    .Cell(flexcpFontBold, i, 1, i, .Cols - 1) = True
                    
                Case "Z3"
                    .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                    .Cell(flexcpFontBold, i, 1, i, .Cols - 1) = True
                
                Case "Z4"
                    .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                    .Cell(flexcpFontBold, i, 1, i, .Cols - 1) = True
            End Select
            
'            Call SetGrdColor(grdData, Mid(.TextMatrix(i, 10), 2), i, 1, i, .Cols - 1)
            
        Next i
        
        
        .MergeCells = flexMergeFree
        For i = 1 To 3
            .MergeCol(i) = True
        Next i
        
        .Redraw = flexRDDirect
    End With
    
    
    
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffINList.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
End Sub

