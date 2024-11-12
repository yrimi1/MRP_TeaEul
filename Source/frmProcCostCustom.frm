VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcCostCustom 
   Caption         =   "가공료 집계표"
   ClientHeight    =   9270
   ClientLeft      =   15
   ClientTop       =   855
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.ComboBox cboTaxClss 
      Height          =   300
      Left            =   1350
      Style           =   2  '드롭다운 목록
      TabIndex        =   9
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   4080
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   1
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
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월 dd일"
      Format          =   116785155
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "청구년월"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   690
      Width           =   11790
      _cx             =   20796
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
      TabIndex        =   5
      Top             =   8520
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
      Index           =   1
      Left            =   2730
      TabIndex        =   6
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월 dd일"
      Format          =   116785155
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   285
      Left            =   9180
      TabIndex        =   7
      Top             =   150
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkTaxClss 
         Caption         =   "사용구분"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   1095
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   360
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "사용구분"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   6810
      TabIndex        =   11
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmProcCostCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
' 변경이력
' 요청ID : S_201111_태을염직_01
' 요청자 : 김대진
' 요청일자: 2011.11.03
' 요청내용 : 가공료집계표 엑셀 내보내기 기능 추가
' 변경일자 : 2011.11.03
' 변경내용 : 엑셀내보내기 버튼및 기능 추가
'*************************************************************************
Option Explicit

'S_201111_태을염직_01 에 의한 추가
Private Sub cmdExcel_Click()
    If grdData(0).Rows = grdData(0).FixedRows Then Exit Sub

    Call MakeExcelGrid(grdData(0))

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
'        Call ColResize(grdData, ES_REDUCE, 30)
 '       Call ColResize(grdData, ES_EXPAND, 30)
        Call FillGrdPrint
    End If


End Sub


Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub


Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String
    
    sDate = MakeDate(DF_FULL, dtpDate(0))
    eDate = MakeDate(DF_FULL, dtpDate(1))
    
    Call SetPrintMode(grdData(0), 1, True)
    With grdData(0)
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .FontName = "돋움체"
        .FontSize = 9
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "가공료 집계표"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "▶ 정산기간 : " & Left(sDate, 10) & " ~ " & Left(eDate, 10)
        .Cell(flexcpText, 2, 1, 2, .Cols - 1) = "▶ 사용구분 : " & IIf(chkTaxClss.Value = vbChecked, "사용", "비사용")
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter

        .ColWidth(8) = 0
        .ExtendLastCol = False
        
         grdData(0).PrintGrid "태을염직", True, 1, 100, 500
        .ColWidth(8) = 400
         
         Call SetPrintMode(grdData(0), 1, False)
         
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
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
    dtpDate(1) = Now
    
    With cboTaxClss
        .AddItem "0.비사용"
        .AddItem "1.사용"
        .AddItem "9.전체"
        .ListIndex = 0
    End With
    

End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim II%, nRows As Integer
    
    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Cols = 11
        .Rows = 4
        .FixedRows = 4
        .FixedCols = 1
        
        .RowHeightMin = 300
        
        
        For II = 0 To nRows - 1
            .RowHidden(II) = True
        
        Next II
        
        
        nRows = 3
        
        .RowHeight(nRows) = 400
        
        .TextMatrix(nRows, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "거래처":           .ColWidth(1) = 2000:                .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(nRows, 2) = "단위":             .ColWidth(2) = 500:                 .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "수량":             .ColWidth(3) = 1600:                .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(nRows, 4) = "단가":             .ColWidth(4) = 700:                 .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(nRows, 5) = "공급가액":         .ColWidth(5) = 1900:                .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(nRows, 6) = "V.A.T":            .ColWidth(6) = 1800:                .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(nRows, 7) = "공급가 총액":      .ColWidth(7) = 1900:                .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(nRows, 8) = "비고":             .ColWidth(8) = 400:                 .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(nRows, 9) = "Depth":            .ColWidth(9) = 0:                   .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "Unitclss":        .ColWidth(10) = 0:                  .ColAlignment(10) = flexAlignCenterCenter
        
        .MergeCells = flexMergeFree
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
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

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    eDate = Left(MakeDate(DF_SHORT, dtpDate(1)), 6)
    
    Screen.MousePointer = vbHourglass
    
    Set rs = oStuffIn.GetProcCostCustom(sDate, eDate, Left(cboTaxClss, 1))
    
    Set oStuffIn = Nothing
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
            i = 0
            Do Until rs.EOF
                i = i + 1
                .AddItem CStr(i) & vbTab & Trim(rs!kCustom) & vbTab & rs!UnitClssSTR & vbTab & IIf(rs!SumQty = 0, "", SetCurrency(rs!SumQty, 0)) & vbTab & _
                        IIf(rs!UnitPrice = 0, "", SetCurrency(rs!UnitPrice, 0)) & vbTab & _
                        IIf(rs!SumPrice = 0, "", SetCurrency(rs!SumPrice, 0)) & vbTab & _
                        IIf(rs!TaxPrice = 0, "", SetCurrency(rs!TaxPrice, 0)) & vbTab & _
                        IIf(rs!SumTot = 0, "", SetCurrency(rs!SumTot, 0)) & vbTab & "" & vbTab & rs!Depth & vbTab & rs!UnitClss
                
                
                If rs!Depth = "Z1" Or rs!Depth = "Z2" Then
                    .TextMatrix(.Rows - 1, 1) = rs!kCustom & " 건"
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                    .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, 1) = flexAlignCenterCenter
                    If rs!Depth = "Z2" Then
                        .Cell(flexcpBackColor, .Rows - 1, 5, .Rows - 1, 7) = PRNHeaderColor
                    End If
                Else
                    If rs!UnitClss = "2" Then
                        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, 1) = flexAlignCenterCenter
                    End If
                End If
                rs.MoveNext
            Loop
        End If
        
        .MergeCells = flexMergeFree
        .MergeCol(1) = True
        
''        For i = 1 To 3
''            .MergeCol(i) = True
''        Next i
        
        .Redraw = flexRDDirect
    End With
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "FrmProcCostReport.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub


