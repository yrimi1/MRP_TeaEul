VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliveryReport 
   Caption         =   "납품사실증명원"
   ClientHeight    =   9270
   ClientLeft      =   1770
   ClientTop       =   3180
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin Threed.SSPanel pnlPrint 
      Height          =   2325
      Left            =   3810
      TabIndex        =   10
      Top             =   3150
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   4101
      _Version        =   196609
      Caption         =   "SSPanel1"
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboCustom 
         Height          =   300
         Left            =   1590
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   510
         Width           =   2115
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   405
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "납품사실 증명원"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   1590
         TabIndex        =   12
         Top             =   840
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1296
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton Option2 
            Caption         =   "A4"
            Height          =   225
            Left            =   180
            TabIndex        =   14
            Top             =   420
            Width           =   1305
         End
         Begin VB.OptionButton Option1 
            Caption         =   "80 컬럼"
            Height          =   225
            Left            =   180
            TabIndex        =   13
            Top             =   120
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   300
         TabIndex        =   16
         Top             =   510
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄범위"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Left            =   300
         TabIndex        =   17
         Top             =   870
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄용지"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   2310
         TabIndex        =   18
         Top             =   1710
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   1710
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "인쇄"
      End
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   4110
      TabIndex        =   2
      Top             =   30
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   5970
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   3
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   30
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월"
      Format          =   78249987
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
      Caption         =   "납품년월"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   5
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
      Height          =   7770
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   690
      Width           =   11790
      _cx             =   20796
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
      Left            =   8490
      TabIndex        =   7
      Top             =   8520
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
      Left            =   5580
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
      Left            =   2790
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
         TabIndex        =   1
         Top             =   60
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDeliveryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const REPORTFILE = "\Report\DeliveryReport.rpt"


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index

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
            
    End Select
End Sub

Private Sub chkSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
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
    End Select
End Sub

Private Sub cmdPrint_Click()
    pnlPrint.Visible = True
End Sub

Private Sub cmdPrnCancel_Click()
    pnlPrint.Visible = False

End Sub

Private Sub cmdPrnOK_Click()
    Dim II%
    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
        If cboCustom.Text = AllStr Then
           
            For II = 1 To cboCustom.ListCount - 1
                Call SetPrnData(Right(cboCustom.List(II), 4))
                
            Next II
        Else
            Call SetPrnData(Right(cboCustom.Text, 4))
        
        End If
    End If
End Sub

Sub SetPrnData(ByVal CustomID As String)
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim sDate As String

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    
    Set rs = oStuffIn.GetDeliveryReport(sDate, 1, CustomID)


    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    ' Printing
    Screen.MousePointer = vbHourglass
    
    Set oStuffIn = Nothing
    
    ReDim sParam(2)
    sParam(0) = ""
    sParam(1) = ""
    sParam(2) = ""
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
    
    rs.Close
    Set rs = Nothing
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "SetPrnData", Err.Description)

End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub



Sub FillGridFooter()
    Dim i%
    Dim sDate As String, eDate As String
    Dim JJ%, nRow%

    If chkSearch(0).Value Then
        sDate = Format(dtpDate(0), "YYYY/MM/DD")
        eDate = Format(dtpDate(1), "YYYY/MM/DD")
    Else
        sDate = ""
        eDate = ""
    End If
    
    '----------------
    ' 인쇄시 관리번호 제외하고 인쇄
    ' chkReport.value = vbchecked 일때 ( 결재용 ) : Design 부분 제외한 후 인쇄
    '-------------------
    
    With grdData(1)
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        
        For i = 0 To .FixedRows
            .RowHidden(0) = False
        Next i

        .FontSize = 7
        nRow = 0
        .Cell(flexcpText, nRow, 1, nRow, .Cols - 1) = "납품사실증명원"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = True
        
        nRow = nRow + 1
        .Cell(flexcpText, nRow, 1, nRow, .Cols - 1) = "사업장 주소"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        nRow = nRow + 1
        .Cell(flexcpText, nRow, 1, nRow, .Cols - 1) = "상       호"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        nRow = nRow + 1
        .Cell(flexcpText, nRow, 1, nRow, .Cols - 1) = "대    표  자"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False

        nRow = nRow + 1
        .Cell(flexcpText, nRow, 1, nRow, .Cols - 1) = "사업자등록번호"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        nRow = nRow + 1
        .Cell(flexcpText, nRow, 3, nRow, .Cols - 1) = "부가가치세법 제11조 1항 및 동법시행 및 제44조 B항에"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False

        nRow = nRow + 1
        .Cell(flexcpText, nRow, 1, nRow, .Cols - 1) = "의하여 귀사에 수출품을 다음과 같이 납품하였음을 증명하여 주시기 바랍니다."
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        .Redraw = flexRDDirect
    End With
    
End Sub




Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11970, 9660

    Call InitGrid(0)
    Call SetOperate(Me)
    
    '----- 날짜설정
    dtpDate(0) = Now
    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    

    cmdFind(0).Enabled = False
    
    txtCustom(1).Enabled = False
    
    pnlPrint.Visible = False

End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim II%, nRows As Integer
    
    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Cols = 9
        .Rows = 10
        .FixedRows = 10
        .FixedCols = 1
        
        .RowHeightMin = 300
        
        nRows = 9
        
        .RowHeight(nRows) = 400
        
        .TextMatrix(nRows, 0) = ""
        .TextMatrix(nRows, 1) = "거래처명":         .ColWidth(1) = 1300:                .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "일자":             .ColWidth(2) = 600:                   .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "L/C NO":           .ColWidth(3) = 1300:                   .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(nRows, 4) = "품명":             .ColWidth(4) = 2000:                .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(nRows, 5) = "수량":             .ColWidth(5) = 1200:                .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(nRows, 6) = "단가":             .ColWidth(6) = 1300:                .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "금액":             .ColWidth(7) = 1300:                .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(nRows, 8) = "OrderNO":          .ColWidth(8) = 1300:                .ColAlignment(8) = flexAlignLeftCenter
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .MergeCells = flexMergeFree
        
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        For II = 0 To .FixedRows - 1
            .RowHidden(II) = True
        Next II
        
        .Redraw = flexRDDirect
    End With

End Sub

Sub FillgrdData()
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String
    Dim dCustom_str As String

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    
    Set rs = oStuffIn.GetDeliveryReport(sDate, IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag)

    Set oStuffIn = Nothing
    cboCustom.Clear
    cboCustom.AddItem AllStr
    With grdData(0)
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        If rs.RecordCount < 1 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
            Do Until rs.EOF
                .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & "" & vbTab & "" & vbTab & rs!Article & vbTab & rs!SumQty & vbTab & rs!WorkUnitPrice & vbTab & rs!Price & rs!CustomID
                
                If Trim(dCustom_str) = "" Then
                    dCustom_str = rs!kCustom
                    cboCustom.AddItem rs!kCustom & "  |  " & rs!CustomID
                ElseIf Trim(dCustom_str) <> Trim(rs!kCustom) Then
                    dCustom_str = rs!kCustom
                    cboCustom.AddItem rs!kCustom & "  |  " & rs!CustomID
                End If
                
                rs.MoveNext
            Loop
        End If
        
        .MergeCells = flexMergeFree
        .MergeCol(1) = True
        
''        For i = 1 To 1
''            .MergeCol(i) = True
''        Next i
        
        .Redraw = flexRDDirect
    End With
    
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "FrmDeliveryReport.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub



Private Sub txtCustom_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 1
            Call MoveFocus(KeyCode)
    End Select

End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call cmdFind_Click(0)
            End If
    End Select
End Sub

