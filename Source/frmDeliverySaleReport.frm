VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliverySaleReport 
   Caption         =   "수출용 원자재 매도 확약서( Offer Sheet )"
   ClientHeight    =   9270
   ClientLeft      =   2625
   ClientTop       =   2415
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.ComboBox cboPriceClss 
      Height          =   300
      Left            =   4260
      Style           =   2  '드롭다운 목록
      TabIndex        =   15
      Top             =   360
      Width           =   1470
   End
   Begin Threed.SSPanel pnlPrint 
      Height          =   3405
      Left            =   3600
      TabIndex        =   9
      Top             =   2670
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   6006
      _Version        =   196609
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboCustom 
         Height          =   300
         Left            =   1410
         Style           =   2  '드롭다운 목록
         TabIndex        =   10
         Top             =   1290
         Width           =   3375
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   405
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   714
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "수출용 원자재 매도 확약서"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1290
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄범위"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   2790
         TabIndex        =   13
         Top             =   2700
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   2700
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "인쇄"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   1410
         TabIndex        =   17
         Top             =   570
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1191
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton opPrn 
            Caption         =   "전체 명세서 인쇄"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   90
            Width           =   2025
         End
         Begin VB.OptionButton opPrn 
            Caption         =   "업체별 명세서 인쇄"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   18
            Top             =   390
            Value           =   -1  'True
            Width           =   2025
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   570
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   1410
         TabIndex        =   21
         Top             =   1650
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1296
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optPRN 
            Caption         =   "수출용 원자재 매도 확약서"
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   23
            Top             =   120
            Value           =   -1  'True
            Width           =   2865
         End
         Begin VB.OptionButton optPRN 
            Caption         =   "구매 승인 신청서"
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   450
            Width           =   1995
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1650
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   4260
      TabIndex        =   1
      Top             =   30
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   6240
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   2
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   780
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   30
      TabIndex        =   3
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
      TabIndex        =   4
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
      Height          =   7830
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   690
      Width           =   11820
      _cx             =   20849
      _cy             =   13811
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
      TabIndex        =   6
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
      Left            =   5730
      TabIndex        =   7
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
      TabIndex        =   8
      Top             =   30
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   2790
      TabIndex        =   16
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "화폐단위"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   25
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월 dd일"
      Format          =   117309443
      CurrentDate     =   36871
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   720
      Index           =   1
      Left            =   90
      TabIndex        =   26
      Top             =   8490
      Visible         =   0   'False
      Width           =   4680
      _cx             =   8255
      _cy             =   1270
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
End
Attribute VB_Name = "frmDeliverySaleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const REPORTFILE = "\Report\DeliverySaleReport.rpt"


Private Sub PrnConfirm(ByVal CustomID As String, ByVal kCustom As String)
    
    Dim II%, JJ%
    
    With grdData(1)
        .Rows = .FixedRows
        For II = 0 To .Cols - 1
            .ColWidth(II) = grdData(0).ColWidth(II)
        Next II
        
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            If Trim(grdData(0).TextMatrix(II, 1)) = Trim(kCustom) Then
                .AddItem ""
                .RowHeight(.Rows - 1) = 350
                For JJ = 2 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
            End If
            .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
        Next II
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ExtendLastCol = False
        .TextMatrix(0, 2) = "DYEING CHAGE FOR " & vbCrLf & "  품       명  "
        'Call ColResize(grdData(1), ES_REDUCE, 10)
        .PrintGrid "태을염직", True, 1, 100, 500
    End With


End Sub


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index
        Case 0
            cboPriceClss.Enabled = chkSearch(0).Value

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
    Dim II%, vCustom As Variant
    
    If opPrn(0).Value = True Then
        Call FillGrdPrint
    Else
        If optPrn(0).Value Then
            If cboCustom.Text = AllStr Then
                For II = 1 To cboCustom.ListCount - 1
                    vCustom = Split(cboCustom.Text, "|")
                    Call SetPrnData(Trim(vCustom(1)), Trim(vCustom(0)))
                    
                Next II
            Else
                vCustom = Split(cboCustom.Text, "|")
                Call SetPrnData(Trim(vCustom(1)), Trim(vCustom(0)))
            
            End If
        Else
            vCustom = Split(cboCustom.Text, "|")
            Call PrnConfirm(Trim(vCustom(1)), Trim(vCustom(0)))
        End If
    End If
    
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
    
    With grdData(0)
'        .Redraw = flexRDNone
        
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = vbWhite
        .GridColor = vbBlack
        
        .RowHidden(0) = False
        .RowHidden(1) = False
        
        .RowHeight(0) = 600
        .RowHeight(1) = 450
        
        .FontSize = 10
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "Offer Sheet 현황"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, 3) = "▶ 납품년월 : " & Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
        .Cell(flexcpText, 1, .Cols - 4, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        
        .Cell(flexcpText, 1, 4, 1, 4) = "▶ 거래처   : " & IIf(chkSearch(1).Value, txtCustom(1).Text, "(전체)")
        .Cell(flexcpText, 2, 7, 2, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        .Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1, .Cols - 1) = vbWhite
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
        .PrintGrid "태을염직", True, 2, 100, 500

        .GridLinesFixed = flexGridInset
        .GridColor = &HE0E0E0
        
'        For i = .FixedRows To .Rows - 1
'            Call SetGrdColor(grdData(0), Mid(.TextMatrix(i, 10), 2), i, 1, i, .Cols - 1)
'        Next i
        
        .FontSize = 9
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .Redraw = flexRDDirect
    End With
    
End Sub


Sub SetPrnData(ByVal CustomID As String, ByVal kCustom As String)
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim sDate As String
    Dim nTotQtyYDS As Long, nTotPrice As Long
    Dim nPriceClss As Integer, sPriceClss As String

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    nTotQtyYDS = 0: nTotPrice = 0
    nPriceClss = cboPriceClss.ItemData(cboPriceClss.ListIndex)
    
    Set rs = oStuffIn.GetDeliverySaleReport(sDate, 1, CustomID, nPriceClss, nTotQtyYDS, nTotPrice)

    
    Me.PopupMenu PlusMDI.mnuPopup
    
    ' Printing
    Screen.MousePointer = vbHourglass
    
    Set oStuffIn = Nothing
    
    ReDim sParam(5)
    sParam(0) = kCustom & "  귀중"
    sParam(1) = ""
    sParam(2) = ""
    sParam(3) = SetCurrency(nTotQtyYDS, 0)
    sParam(4) = IIf(nPriceClss = "0", SetCurrency(nTotPrice, 0), SetCurrency(nTotPrice, 2))
    sParam(5) = IIf(nPriceClss = "0", "\", "$")
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
    
'    rs.Close
'    Set rs = Nothing
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, "SetPrnData", Err.Description)

End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
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
    Call InitGrid(1)
    Call SetOperate(Me)
    
    '----- 날짜설정
    
    dtpDate(0) = Now
    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)

    cmdFind(0).Enabled = False
    
    txtCustom(1).Enabled = False
    
    pnlPrint.Visible = False

    ' 화폐구분
    With cboPriceClss
        .AddItem "원화":        .ItemData(0) = 0
        .AddItem "달러":        .ItemData(1) = 1
        .AddItem AllStr:        .ItemData(2) = 9
    End With
    
    cboPriceClss.ListIndex = 0

End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim II%, nRows As Integer
    
    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Cols = 8
        .Rows = 10
        .FixedRows = 10
        .FixedCols = 1
        
        .RowHeightMin = 300
        
        nRows = 9
        
        .RowHeight(nRows) = 400
        
        .TextMatrix(nRows, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "거래처명":         .ColWidth(1) = 2800:         .ColAlignment(1) = flexAlignCenterCenter:     .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "품명":             .ColWidth(2) = 2400:         .ColAlignment(2) = flexAlignLeftCenter:       .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "가공방법":         .ColWidth(3) = 2000:         .ColAlignment(3) = flexAlignCenterCenter:     .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "단가":             .ColWidth(4) = 1200:         .ColAlignment(4) = flexAlignRightCenter:      .FixedAlignment(4) = flexAlignCenterCenter
        .TextMatrix(nRows, 5) = "출고수량":         .ColWidth(5) = 1300:         .ColAlignment(5) = flexAlignRightCenter:      .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "공급가액":         .ColWidth(6) = 1300:         .ColAlignment(6) = flexAlignRightCenter:      .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "Depth":            .ColWidth(7) = 0:            .ColAlignment(7) = flexAlignRightCenter:      .FixedAlignment(7) = flexAlignCenterCenter
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .MergeCells = flexMergeFree
        
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        For II = 0 To .FixedRows - 2
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
    Dim nTotQtyYDS As Long, nTotPrice As Long
    Dim nPriceClss As Integer, sPriceClss As String

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    nTotQtyYDS = 0: nTotPrice = 0
    nPriceClss = cboPriceClss.ItemData(cboPriceClss.ListIndex)
    
    Set rs = oStuffIn.GetDeliverySaleReport(sDate, IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag _
                                           , nPriceClss, nTotQtyYDS, nTotPrice)

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
            .Rows = .FixedRows
            Do Until rs.EOF
                If Trim(rs!kCustom) <> Trim(.TextMatrix(.Rows - 1, 1)) Then
                    .AddItem "" & vbTab & Trim(rs!kCustom)
                    .RowHidden(.Rows - 1) = True
                     cboCustom.AddItem rs!kCustom & "  |  " & rs!CustomID
                End If
                
                .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & IIf(Trim(rs!Article) = "ZZZZ", "소계", Trim(rs!Article)) & vbTab & Trim(rs!WorkName) & vbTab & _
                         rs!PriceClssSTR & SetCurrency(rs!WorkUnitPrice, 2) & vbTab & SetCurrency(rs!SumQty) & " " & rs!UnitClss & vbTab & _
                         rs!PriceClssSTR & IIf(rs!PriceClssSTR = "$", SetCurrency(rs!AmountPrice, 2), SetCurrency(rs!AmountPrice, 0)) & vbTab & rs!Depth
                If rs!Depth <> "Z0" Then
                    .TextMatrix(.Rows - 1, 4) = ""
                    Call SetGrdColor(grdData(0), Right(rs!Depth, 1), .Rows - 1, 2, .Rows - 1, .Cols - 1)
                End If
                rs.MoveNext
            Loop
            
        End If
        .MergeCol(1) = True
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



Private Sub opPrn_Click(Index As Integer)
    If Index = 0 Then
        cboCustom.Enabled = False
    Else
        cboCustom.Enabled = True
    End If
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

