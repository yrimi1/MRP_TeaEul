VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetAccount 
   ClientHeight    =   9270
   ClientLeft      =   1665
   ClientTop       =   1530
   ClientWidth     =   11850
   Icon            =   "frmSetAccount.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11850
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   630
      Left            =   7020
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   17
      ToolTipText     =   "자료 저장"
      Top             =   0
      Width           =   780
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   4590
      TabIndex        =   16
      Top             =   0
      Width           =   2025
   End
   Begin Threed.SSPanel pnlPrint 
      Height          =   2955
      Left            =   3720
      TabIndex        =   0
      Top             =   2880
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5212
      _Version        =   196609
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboCustom 
         Height          =   300
         Left            =   1680
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   1260
         Width           =   2115
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   405
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   714
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "품명별 정산서 인쇄"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   1680
         TabIndex        =   2
         Top             =   3030
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1296
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton Option1 
            Caption         =   "80 컬럼"
            Height          =   225
            Left            =   180
            TabIndex        =   4
            Top             =   120
            Width           =   1305
         End
         Begin VB.OptionButton Option2 
            Caption         =   "A4"
            Height          =   225
            Left            =   180
            TabIndex        =   3
            Top             =   420
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   450
         TabIndex        =   6
         Top             =   1260
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄범위"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   7
         Top             =   3030
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄용지"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   2250
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   450
         TabIndex        =   9
         Top             =   2250
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "인쇄"
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   1650
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyy년 MM월 dd일"
         Format          =   54984707
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   1650
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄일자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   645
         Left            =   1680
         TabIndex        =   12
         Top             =   570
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1138
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton opPrn 
            Caption         =   "전체 명세서 인쇄"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   14
            Top             =   90
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.OptionButton opPrn 
            Caption         =   "업체별 명세서 인쇄"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   13
            Top             =   390
            Width           =   2145
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   450
         TabIndex        =   15
         Top             =   570
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "인쇄구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   18
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy년 MM월 dd일"
      Format          =   54984707
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "정산일자"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   20
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
      Index           =   0
      Left            =   0
      TabIndex        =   21
      Top             =   660
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
      Left            =   8460
      TabIndex        =   22
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
      Left            =   6630
      TabIndex        =   23
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
      Left            =   3300
      TabIndex        =   24
      Top             =   0
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
         TabIndex        =   25
         Top             =   60
         Width           =   975
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   600
      Index           =   1
      Left            =   1650
      TabIndex        =   26
      Top             =   8520
      Visible         =   0   'False
      Width           =   2520
      _cx             =   4445
      _cy             =   1058
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
   Begin VSFlex7LCtl.VSFlexGrid grdCoverTitle 
      Height          =   450
      Left            =   4560
      TabIndex        =   27
      Top             =   8610
      Visible         =   0   'False
      Width           =   3060
      _cx             =   5397
      _cy             =   794
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
Attribute VB_Name = "frmSetAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



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
    If MsgBox("인쇄하시겠습니까", vbYesNo) = vbYes Then
    '    pnlPrint.Visible = True
            Call ColResize(grdData(0), ES_REDUCE, 17)
            Call FillGrdList
            Call ColResize(grdData(0), ES_EXPAND, 17)
    End If
End Sub

Private Sub cmdPrnCancel_Click()
    pnlPrint.Visible = False
End Sub

Private Sub cmdPrnOK_Click()
    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
        If opPrn(0).Value = True Then
            Call ColResize(grdData(0), ES_REDUCE, 10)
            Call FillGrdList
            Call ColResize(grdData(0), ES_EXPAND, 10)
        Else
            Call FillGrdPrint
        End If
    End If

End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub

Sub FillGrdPrint()
    Dim II%
    
    If cboCustom.Text = AllStr Then
       
        For II = 1 To cboCustom.ListCount - 1
            Call SetDataToPrn(cboCustom.List(II))
            
        Next II
    Else
        Call SetDataToPrn(cboCustom.Text)
    
    End If
    
End Sub

Sub FillGrdPrintHeader(ByVal kCustom As String)
    Dim i%
    Dim sDate As String
    
    sDate = Format(dtpDate(0), "YYYY/MM/DD")
    
    With grdData(1)
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .FontSize = 7
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "품명별 정산서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "▶ 거 래 처 : " & kCustom
        .Cell(flexcpText, 2, 1, 2, .Cols - 1) = "▶ 정산일자 : " & sDate & " 현재 "
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter

        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    
End Sub
'표지출력
Sub FillPrintCoverTitle(ByVal SumPrice As Double, ByVal VatPrice As Double, ByVal kCustom As String)
    Dim nSumTotalPrice As Double, sTotPrice As String
    Dim sDate As String
    Dim nRow As Integer, II As Integer
    
    sDate = Format(dtpDate(0), "YYYY/MM/DD")
    
    
    
    nSumTotalPrice = SumPrice + VatPrice
    
    sTotPrice = ALP_TO_STR(nSumTotalPrice)
    
    With grdCoverTitle
        .Redraw = flexRDBuffered
        .ExtendLastCol = False
        
        .Rows = 30
        .Cols = 1
        .FixedCols = 0
        .FixedRows = 30
        .GridLinesFixed = flexGridNone
        
        
        nRow = 2
        .Cell(flexcpText, nRow, 0, nRow, 0) = "請     求     書"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 24
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = True
        
        
        nRow = 4
        .Cell(flexcpText, nRow, 0, nRow, 0) = "一金  " & sTotPrice & "(￦" & SetCurrency(nSumTotalPrice) & ")원整"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        
        
        nRow = 6
        .Cell(flexcpText, nRow, 0, nRow, 0) = "임가공료  : ￦" & SetCurrency(SumPrice)
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        
        nRow = 8
        .Cell(flexcpText, nRow, 0, nRow, 0) = "부  가 세  : ￦" & SetCurrency(VatPrice)
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        
        nRow = 10
        .Cell(flexcpText, nRow, 0, nRow, 0) = Mid(MakeDate(DF_SHORT, dtpDate(0)), 5, 2) & "월분 임가공료를 상기와 같이 청구합니다."
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        
        nRow = 12    '-----   2004.03.15
        .Cell(flexcpText, nRow, 0, nRow, 0) = MakeDate(DF_FULL, dtpDate(1))
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        
        nRow = 14    '-----
        .Cell(flexcpText, nRow, 0, nRow, 0) = "청구자 주소: 대구시 서구 비산동 2009-45번지 "
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        nRow = 16    '-----
        .Cell(flexcpText, nRow, 0, nRow, 0) = "       상호: 대영염공  주식회사"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        nRow = 18    '-----
        .Cell(flexcpText, nRow, 0, nRow, 0) = "      대표자: 이정화 (인) "
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        nRow = 20    '-----
        .Cell(flexcpText, nRow, 0, nRow, 0) = kCustom & " 貴中"
        .Cell(flexcpFontSize, nRow, 0, nRow, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRow, 0, nRow, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRow, 0, nRow, .Cols - 1) = False
        
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = vbWhite
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterCenter

        .ExtendLastCol = True
        
        
        For II = 0 To .Rows - 1
            .RowHeight(II) = 500
        Next II
        
        .Redraw = flexRDDirect
    End With
    
End Sub
Sub FillGrdList()
    Dim i%
    Dim sDate As String
    
    sDate = Format(dtpDate(0), "YYYY/MM/DD")
    
    With grdData(0)
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .FontSize = 8
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "품명별 정산서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 2, 1, 2, .Cols - 1) = "▶ 정산일자 : " & sDate & " 현재 "
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter

        .ExtendLastCol = True
        .Redraw = flexRDDirect
        .PrintGrid "대염영공(주)", False, 1, 500, 500
        
        .FontSize = 9
        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
    End With

End Sub
Sub SetDataToPrn(ByVal kCustom As String)
    Dim II%, JJ%
    Dim SumPrice As Double, VatPrice As Double
    
    Call FillGrdPrintHeader(kCustom)
    With grdData(1)
        .Rows = .FixedRows
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            If grdData(0).TextMatrix(II, 1) = kCustom Then
                .AddItem ""
                For JJ = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
            End If
            .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
        Next II
        
        SumPrice = .ValueMatrix(.Rows - 1, 8)
        VatPrice = .ValueMatrix(.Rows - 1, 9)
        
        .ColHidden(1) = True
        .ColHidden(0) = True
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, .Cols - 2) = "공급가액: " & SetCurrency(SumPrice, 0) & " 원" & _
                                                   Space(20) & "부가세: " & SetCurrency(VatPrice, 0) & " 원" & _
                                                   Space(20) & "총금액: " & SetCurrency(SumPrice + VatPrice, 0) & " 원"
        
        
        
        
    '    .Cell(flexcpText, .Rows - 1, 4, .Rows - 1, 7) = "  부가세: " & SetCurrency(VatPrice, 0) & " 원"
    '    .Cell(flexcpText, .Rows - 1, 8, .Rows - 1, .Cols - 2) = "  총금액: " & SetCurrency(SumPrice + VatPrice, 0) & " 원"
        .Cell(flexcpAlignment, .Rows - 1, 2, .Rows - 1, .Cols - 2) = flexAlignCenterCenter
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 2) = vbWhite
                                             
        .MergeRow(.Rows - 1) = True

        
        
        
    End With
    
    
    Call FillPrintCoverTitle(SumPrice, VatPrice, kCustom)
    
    '인쇄하기
    grdCoverTitle.PrintGrid "", False, 1, 200, 1000
    grdData(1).PrintGrid "품명별 정산서", False, 1, 200, 500

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
    dtpDate(1) = Now
    
'    CboStuffClss2.ListIndex = 0
    
    cboCustom.Enabled = False
    
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
        .TextMatrix(nRows, 1) = "거래처":           .ColWidth(1) = 1700:      .ColAlignment(1) = flexAlignLeftTop:      .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "품명":             .ColWidth(2) = 1900:      .ColAlignment(2) = flexAlignLeftTop:        .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "전월 이월" & vbCrLf & "당월 입고":         .ColWidth(3) = 1200:      .ColAlignment(3) = flexAlignRightTop:      .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "출고가공":             .ColWidth(4) = 1200:      .ColAlignment(4) = flexAlignLeftTop:       .FixedAlignment(4) = flexAlignCenterCenter
        .TextMatrix(nRows, 5) = "ORDER NO" & vbCrLf & "ORDER 량":           .ColWidth(5) = 1600:      .ColAlignment(5) = flexAlignLeftTop:      .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "소요량":         .ColWidth(6) = 600:      .ColAlignment(6) = flexAlignRightCenter:       .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "소요량":         .ColWidth(7) = 1000:      .ColAlignment(7) = flexAlignRightCenter:       .FixedAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = "출고 누계" & "당월 출고":         .ColWidth(8) = 1000:      .ColAlignment(8) = flexAlignRightCenter:       .FixedAlignment(8) = flexAlignCenterCenter
        .TextMatrix(nRows, 9) = "ORDER 잔량" & vbCrLf & "당월 재고":           .ColWidth(9) = 1200:      .ColAlignment(9) = flexAlignRightCenter:       .FixedAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "Depth":    .ColWidth(10) = 0
        
        .ColFormat(6) = "#,###"
        .ColFormat(7) = "#,###"
        .ColFormat(8) = "#,###"
        .ColFormat(9) = "#,###"
        
        .MergeCells = flexMergeFree
        .MergeRow(3) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        
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
    Dim oSetAccount As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nFlag%
    Dim sCustom$, sArticle$, sWork$, sOrderNo$
    Dim sCustom1$, sArticle1$, sWork1$, sOrderNo1$
    Dim nChkCnt%, nWorkCnt%, nItemCnt%
    Dim sUnitClss$
    
    On Error GoTo ErrHandler

    Set oSetAccount = New PlusLib2.CSubul
    oSetAccount.Connection = g_adoCon
    oSetAccount.UserName = g_sUserName
        
    
    Set rs = oSetAccount.GetSetAccountByArticle(MakeDate(DF_SHORT, dtpDate(0)), IIf(chkSearch(1).Value, 1, 0), txtCustom(1).Tag)

    Set oSetAccount = Nothing
    
    cboCustom.Clear
    cboCustom.AddItem AllStr
    
    With grdData(0)
        .Rows = .FixedRows
        .Redraw = flexRDNone

        If rs.RecordCount < 1 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
            Do Until rs.EOF
                If rs!UnitClss = "0" Then
                    sUnitClss = ""
                ElseIf rs!UnitClss = "1" Then
                    sUnitClss = "M"
                End If
                
                If sCustom <> rs!kCustom Then
                    .AddItem ""
                    .RowHidden(.Rows - 1) = True
                    cboCustom.AddItem rs!kCustom
                End If
                
                If nFlag = 0 Or (rs!Depth = 0 And sOrderNo <> rs!OrderNo) Then
'                If nFlag = 0 Then
                    .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & IIf(rs!StockQty = 0, "", Format(rs!StockQty, "#,###")) & vbCrLf & IIf(rs!StuffinQty = 0, "", Format(rs!StuffinQty, "#,###")) & vbTab & Trim(rs!WorkName) & vbTab & _
                    Trim(rs!OrderNo) & vbCrLf & MakeStrBySpace(Format(rs!OrderQty, "#,###"), 15, 0) & sUnitClss & vbTab & _
                    "" & vbTab & IIf(rs!TOutRealQty = 0, "", rs!TOutRealQty) & vbTab & IIf(rs!TOutQtyYDS = 0, "", rs!TOutQtyYDS) & vbTab & "" & vbTab & rs!Depth
                    nFlag = 1
                End If
                
                If rs!Depth = 0 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Trim(rs!kCustom)
                    .TextMatrix(.Rows - 1, 2) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!StockQty = 0, "", Format(rs!StockQty, "#,###")) & vbCrLf & IIf(rs!StuffinQty = 0, "", Format(rs!StuffinQty, "#,###"))
                    .TextMatrix(.Rows - 1, 4) = Trim(rs!WorkName)
                    .TextMatrix(.Rows - 1, 5) = Trim(rs!OrderNo) & vbCrLf & MakeStrBySpace(Format(rs!OrderQty, "#,###"), 15, 0) & sUnitClss
                    .TextMatrix(.Rows - 1, 6) = Right(rs!ResultDate, 2)
                    .TextMatrix(.Rows - 1, 7) = rs!OutRealQty
                    .TextMatrix(.Rows - 1, 8) = rs!OutQtyYDS
                    .TextMatrix(.Rows - 1, 9) = IIf(rs!UnitClss = "0", "", "[" & Format(rs!OutQty, "#,###") & "M]")
                    .TextMatrix(.Rows - 1, 10) = rs!Depth
                    .Cell(flexcpAlignment, .Rows - 1, 9, .Rows - 1, 9) = flexAlignLeftCenter
                ElseIf rs!Depth = 1 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Trim(rs!kCustom)
                    .TextMatrix(.Rows - 1, 2) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!StockQty = 0, "", Format(rs!StockQty, "#,###")) & vbCrLf & IIf(rs!StuffinQty = 0, "", Format(rs!StuffinQty, "#,###"))
                    .TextMatrix(.Rows - 1, 4) = Trim(rs!WorkName)
                    .TextMatrix(.Rows - 1, 5) = .TextMatrix(.Rows - 2, 5)
                    .TextMatrix(.Rows - 1, 6) = "계"
                    .TextMatrix(.Rows - 1, 7) = rs!OutRealQty
                    .TextMatrix(.Rows - 1, 8) = rs!OutQtyYDS
                    .TextMatrix(.Rows - 1, 9) = rs!OrderQty - rs!TOutQty - rs!OutQty
                    .TextMatrix(.Rows - 1, 10) = rs!Depth
                
                    .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = ED1_DEPTH
                ElseIf rs!Depth = 2 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Trim(rs!kCustom)
                    .TextMatrix(.Rows - 1, 2) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!StockQty = 0, "", Format(rs!StockQty, "#,###")) & vbCrLf & IIf(rs!StuffinQty = 0, "", Format(rs!StuffinQty, "#,###"))
                    .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 2, 4)
                    .TextMatrix(.Rows - 1, 5) = "출고가공계"
                    .TextMatrix(.Rows - 1, 6) = "" 'Right(rs!ResultDate, 2)
                    .TextMatrix(.Rows - 1, 7) = rs!OutRealQty
                    .TextMatrix(.Rows - 1, 8) = rs!OutQtyYDS
                    .TextMatrix(.Rows - 1, 9) = ""
                    .TextMatrix(.Rows - 1, 10) = rs!Depth
                
                    .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = ED2_DEPTH
                    .Cell(flexcpFontBold, .Rows - 1, 3, .Rows - 1, .Cols - 1) = True
                ElseIf rs!Depth = 3 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Trim(rs!kCustom)
                    .TextMatrix(.Rows - 1, 2) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!StockQty = 0, "", Format(rs!StockQty, "#,###")) & vbCrLf & IIf(rs!StuffinQty = 0, "", Format(rs!StuffinQty, "#,###"))
                    .TextMatrix(.Rows - 1, 4) = "소계"
                    .TextMatrix(.Rows - 1, 5) = ""
                    .TextMatrix(.Rows - 1, 6) = "" 'Right(rs!ResultDate, 2)
                    .TextMatrix(.Rows - 1, 7) = rs!OutRealQty
                    .TextMatrix(.Rows - 1, 8) = rs!OutQtyYDS
                    .TextMatrix(.Rows - 1, 9) = IIf(rs!StockQty + rs!StuffinQty - rs!OutRealQty = 0, "", rs!StockQty + rs!StuffinQty - rs!OutRealQty)
                    .TextMatrix(.Rows - 1, 10) = rs!Depth

                    .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = ED3_DEPTH
                    .Cell(flexcpFontBold, .Rows - 1, 3, .Rows - 1, .Cols - 1) = True
                ElseIf rs!Depth = 4 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Trim(rs!kCustom)
                    .TextMatrix(.Rows - 1, 2) = "합 계  "
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!StockQty + rs!StuffinQty = 0, "", Format(rs!StockQty + rs!StuffinQty, "#,###"))
                    .TextMatrix(.Rows - 1, 4) = ""
                    .TextMatrix(.Rows - 1, 5) = ""
                    .TextMatrix(.Rows - 1, 6) = ""
                    .TextMatrix(.Rows - 1, 7) = rs!OutRealQty
                    .TextMatrix(.Rows - 1, 8) = rs!OutQtyYDS
                    .TextMatrix(.Rows - 1, 9) = IIf(rs!StockQty + rs!StuffinQty - rs!OutRealQty = 0, "", rs!StockQty + rs!StuffinQty - rs!OutRealQty)
                    .TextMatrix(.Rows - 1, 10) = rs!Depth
                
                    .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = ED4_DEPTH
                    .Cell(flexcpFontBold, .Rows - 1, 2, .Rows - 1, .Cols - 1) = True
                End If
                
                sCustom = rs!kCustom
                sArticle = rs!Article
                sWork = rs!WorkName
                sOrderNo = rs!OrderNo
                rs.MoveNext
            Loop
        End If
                        
        .Redraw = flexRDDirect
    End With
    cboCustom.ListIndex = 0
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "FrmSetAccount.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oSetAccount = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub



Private Sub opPrn_Click(Index As Integer)
    Select Case Index
    Case 0: cboCustom.Enabled = False
    Case 1: cboCustom.Enabled = True
    End Select
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



