VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutWareReport 
   Caption         =   "Order별 출고명세서"
   ClientHeight    =   9345
   ClientLeft      =   1260
   ClientTop       =   2475
   ClientWidth     =   15255
   Icon            =   "frmOutWareReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   15255
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   6945
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   15195
      _cx             =   26802
      _cy             =   12250
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
   Begin Threed.SSFrame frmSearch 
      Height          =   585
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1032
      _Version        =   196609
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   480
         Left            =   9030
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   2
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   840
      End
      Begin VB.TextBox txtOrderID 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6960
         TabIndex        =   1
         Top             =   150
         Width           =   1485
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   2
         Left            =   2850
         TabIndex        =   5
         Top             =   150
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
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
         Caption         =   "수불년월"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   1
         Left            =   4380
         TabIndex        =   0
         Top             =   150
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   582
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   0
         Left            =   5430
         TabIndex        =   10
         Top             =   150
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
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
         Caption         =   "관리번호"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   0
         Left            =   8460
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   405
         Left            =   90
         TabIndex        =   15
         Top             =   90
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   714
         _Version        =   196609
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11790
      TabIndex        =   7
      Tag             =   "PERM_ADDNEW"
      Top             =   8640
      Visible         =   0   'False
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
      Left            =   13530
      TabIndex        =   8
      Top             =   8640
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   345
      Index           =   1
      Left            =   7380
      TabIndex        =   9
      Top             =   8850
      Visible         =   0   'False
      Width           =   2565
      _cx             =   4524
      _cy             =   609
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
   Begin VSFlex7LCtl.VSFlexGrid grdHeader 
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   15195
      _cx             =   26802
      _cy             =   1931
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
      Begin VB.CommandButton Command1 
         Caption         =   "검색"
         Height          =   345
         Left            =   5820
         TabIndex        =   13
         Top             =   690
         Width           =   585
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdStuffIN 
      Height          =   2895
      Left            =   7500
      TabIndex        =   12
      Top             =   1950
      Visible         =   0   'False
      Width           =   6495
      _cx             =   11456
      _cy             =   5106
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
      Begin VB.CommandButton Command2 
         Caption         =   "닫기"
         Height          =   375
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmOutWareReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LIMIT_ROW5 = 10
Private m_bloading As Boolean
Private sDate As String, eDate As String


Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0
            Call ReturnCode(LG_ORDER, , False, txtOrderID)
            If Trim(txtOrderID.Tag) = "" Then
'                txtOrderNo.Text = ""
'                txtOrderID.Text = ""
'                txtCustomID.Text = ""
'                TxtArticleID2.Text = ""
            Else
                txtOrderID.Text = txtOrderID.Tag
'                If FillStuffOrderData(txtOrderID) Then
'                    txtCustomID.Enabled = False
'                    TxtArticleID2.Enabled = False
'                Else
'                    txtCustomID.Enabled = True
'                    TxtArticleID2.Enabled = True
'                End If
            End If
    End Select

End Sub

Private Sub cmdPrint_Click()
    Dim II%, nCount As Integer, vCol() As Integer, JJ%
    Dim nRows As Integer
    Dim sPrinter As String
    
    sPrinter = Printer.DeviceName
    If Not frmPrinter.SelectPrinter(sPrinter) Then
        Exit Sub
    End If
    
    With grdData(0)
        nCount = 0
        For II = 0 To .Cols - 1
            If .ColWidth(II) > 0 Then
                ReDim Preserve vCol(nCount)
                vCol(nCount) = II
                nCount = nCount + 1
            End If
        Next II
    End With
    
    Call SetVSFlexGrid(grdData(1))
    With grdData(1)
        .Rows = 11
        .FixedRows = 11
        .FixedCols = 0
        .Cols = UBound(vCol) + 1
                 
        .RowHidden(7) = True
        '---- 데이터 Header Title과 데이터 옮기기
        For II = 0 To UBound(vCol)
            .TextMatrix(8, II) = grdData(0).TextMatrix(2, vCol(II))
            .TextMatrix(9, II) = grdData(0).TextMatrix(3, vCol(II))
            .TextMatrix(10, II) = grdData(0).TextMatrix(4, vCol(II))
            .ColWidth(II) = grdData(0).ColWidth(vCol(II))
            .FixedAlignment(II) = grdData(0).FixedAlignment(vCol(II))
        Next II
        .RowHeight(7) = 350
        .RowHeight(8) = 350
        .RowHeight(9) = 350
        
        '---- 데이터 옮기기
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            .AddItem ""
            For JJ = 0 To UBound(vCol)
                .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, vCol(JJ))
            Next JJ
        Next II
        
        .MergeCells = flexMergeFixedOnly
        For II = 0 To 9
            .MergeRow(II) = True
        Next II
        
        For II = 0 To .Cols - 1
            .MergeCol(II) = True
        Next II
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With
    
    
    '---- Print Title을 옮김
    
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
    Dim i%, sOrderID$, bFlag As Boolean, nDay%
    Dim nBaseCol As Integer
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    With grdData(1)
    
        Set rs = oSubul.GetOutWareReportOrder(sDate, eDate, IIf(optOrder(0).Value = True, "0", "1"), txtOrderID)
        If rs.RecordCount > 0 Then
            
            .RowHidden(0) = True
            .RowHidden(1) = True
            
            nRows = 2
            .RowHeight(nRows) = 1000
            .Cell(flexcpText, nRows, 0, nRows, .Cols - 1) = "ORDER별 출고명세서"
            .Cell(flexcpBackColor, nRows, 0, nRows, .Cols - 1) = vbWhite
            
            .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
            .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
'            .Cell(flexcpFontUnderline, nRows, 0, nRows, .Cols - 1) = True
            
            nRows = 3
            .RowHeight(nRows) = 350
            .Cell(flexcpText, nRows, 0, nRows, 0) = "기간"
            .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = MakeDate(DF_LONG, sDate) & " ~ " & MakeDate(DF_LONG, eDate)
            .Cell(flexcpAlignment, nRows, 1, nRows, .Cols - 1) = flexAlignLeftCenter
        
            nRows = 4
            .RowHeight(nRows) = 350
            .Cell(flexcpText, nRows, 0, nRows, 0) = "거래처"
            .Cell(flexcpText, nRows, 1, nRows, 3) = Trim(rs!kCustom)
            .Cell(flexcpFontBold, nRows, 1, nRows, 3) = True
            
            .Cell(flexcpText, nRows, 4, nRows, 4) = "Order NO"
            .Cell(flexcpText, nRows, 5, nRows, 6) = Trim(rs!OrderNo)
            .Cell(flexcpFontBold, nRows, 5, nRows, 6) = True
            
            .Cell(flexcpText, nRows, 7, nRows, 7) = "관리번호"
            .Cell(flexcpText, nRows, 8, nRows, 9) = Trim(rs!OrderID)
            .Cell(flexcpFontBold, nRows, 8, nRows, 9) = True
            
            .Cell(flexcpText, nRows, 10, nRows, 10) = "단가"
            .Cell(flexcpText, nRows, 11, nRows, 12) = SetCurrency(rs!UnitPrice, 2)
            .Cell(flexcpFontBold, nRows, 11, nRows, 12) = True
            
            .Cell(flexcpAlignment, nRows, 0, nRows, .Cols - 1) = flexAlignCenterCenter
            
            nRows = 5
            .RowHeight(nRows) = 350
            .Cell(flexcpText, nRows, 0, nRows, 0) = "ITEM"
            .Cell(flexcpText, nRows, 1, nRows, 3) = Trim(rs!Article)
            .Cell(flexcpFontBold, nRows, 1, nRows, 3) = True
            
            .Cell(flexcpText, nRows, 4, nRows, 4) = "가공방법"
            .Cell(flexcpText, nRows, 5, nRows, 6) = Trim(rs!WorkName)
            .Cell(flexcpFontBold, nRows, 5, nRows, 6) = True
            
            .Cell(flexcpText, nRows, 7, nRows, 7) = "ORDER량"
            .Cell(flexcpText, nRows, 8, nRows, 9) = SetCurrency(rs!OrderQty) & IIf(rs!UnitClss = "0", "  Y", "  M")
            .Cell(flexcpFontBold, nRows, 8, nRows, 9) = True
            
            .Cell(flexcpText, nRows, 10, nRows, 10) = "축율(%)"
            .Cell(flexcpText, nRows, 11, nRows, 12) = SetCurrency(rs!ChunkRate, 2) & " + " & SetCurrency(rs!LossRate, 2)
            .Cell(flexcpFontBold, nRows, 11, nRows, 12) = True
            
            .Cell(flexcpAlignment, nRows, 0, nRows, .Cols - 1) = flexAlignCenterCenter
            
            nRows = 6
            .RowHeight(nRows) = 350
            .Cell(flexcpText, nRows, 0, nRows, 0) = "I/G환산수량"
            .Cell(flexcpText, nRows, 1, nRows, 1) = SetCurrency(rs!IGQty)
            .Cell(flexcpFontBold, nRows, 1, nRows, 1) = True
            
            .Cell(flexcpText, nRows, 2, nRows, 2) = "가공구분"
            .Cell(flexcpText, nRows, 3, nRows, 4) = Trim(rs!WorkName)
            .Cell(flexcpFontBold, nRows, 3, nRows, 4) = True
            
            .Cell(flexcpText, nRows, 5, nRows, 5) = "입고내역"
            .Cell(flexcpText, nRows, 6, nRows, 11) = "[전월누계:" & SetCurrency(rs!PreQty, 0) & "]    [당월입고:" & SetCurrency(rs!CurQty, 0) & _
                                            "]    [합계:" & SetCurrency(rs!TotQty, 0) & "]    [과부족:" & SetCurrency(rs!OvrQty, 0) & "]"
            .Cell(flexcpFontBold, nRows, 6, nRows, 11) = True
                                            
            .Cell(flexcpText, nRows, 12, nRows, 12) = IIf(rs!CloseClss = "*", "완결", "미완결")
            .Cell(flexcpFontBold, nRows, 12, nRows, 12) = True
            
            .Cell(flexcpAlignment, nRows, 0, nRows, .Cols - 1) = flexAlignCenterCenter
            
            .MergeRow(nRows) = True
            
            .Cell(flexcpBackColor, 3, 0, .FixedRows, .Cols - 1) = vbWhite
            
            '--- grid의 사이즈를 90% 대로 줄임
            For II = 0 To .Cols - 1
                .ColWidth(II) = Int(.ColWidth(II) * 0.95)
            Next II
            .Redraw = flexRDDirect
        End If
        
        Set oSubul = Nothing
        rs.Close
        Set rs = Nothing
        
        For II = 0 To .Rows - 1
            .RowHeight(II) = 500
        Next II
        
        ' 소요량, 계산량 row는 보이지 않게
        .RowHidden(.Rows - 1) = True
        .RowHidden(.Rows - 2) = True
        .Cell(flexcpFontBold, .Rows - 3, 0, .Rows - 3, .Cols - 1) = True
        
        .PrintGrid "태을염직", True, 2, 100, 500
    End With
    
    Call ReturnPrinter(sPrinter)
    
    
    Exit Sub
ErrHandler:
End Sub





Private Sub Command1_Click()
    Call FillGrdStuffIN
End Sub

Private Sub Command2_Click()
    grdStuffIN.Visible = False
End Sub

Private Sub dtpDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 15360, 9840
    
    Call SetOperate(Me)
    Call ChangeMode(Me, True)
    
    dtpDate(1) = Now
    Call InitGridHeader
    Call InitGrid(0)
    
    Call InitGridSub
    
    grdStuffIN.ZOrder 0
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdPrint.Visible = False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub


' 입고 자료
Sub FillGrdStuffIN()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset

    Screen.MousePointer = vbHourglass

   ' On Error GoTo ErrHandler

    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon

    Set rs = oSubul.GetOutWareReportStuffIN(Trim(txtOrderID.Text))
    Set oSubul = Nothing
    
    With grdStuffIN
        .Rows = .FixedRows
    End With

    With grdStuffIN
        .Rows = .FixedRows
        Do Until rs.EOF
            .AddItem MakeDate(DF_LONG, rs!StuffDate) & vbTab & SetCurrency(rs!Roll) & vbTab & SetCurrency(rs!Qty) & vbTab & rs!Custom
            rs.MoveNext
        Loop
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        
            If .Rows < LIMIT_ROW5 Then
                .Height = (.RowHeight(.FixedRows) + 40) * .Rows + 350
                .ScrollBars = flexScrollBarNone
            Else
                .Height = 2700
                .ScrollBars = flexScrollBarVertical
            End If
            grdStuffIN.Visible = True
        Else
            MsgBox (" 검색할 내용이 없습니다.")
            grdStuffIN.Visible = False
        End If
    End With
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub cmdSearch_Click()
    Call InitGrid(0)
    Call FillGrdOrder
    Call FillGridData
End Sub

Private Sub FillGrdOrder()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
    Dim i%, sOrderID$, bFlag As Boolean, II%, JJ%, nDay%
    Dim nBaseCol As Integer
    
    
    Dim nRow As Integer

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon

    sDate = Left(MakeDate(DF_SHORT, dtpDate(1)), 6) + "01"
    
    ' 검색월과 시스템월이 같은지 비고
    If MakeDate(DF_SHORT, dtpDate(1)) > MakeDate(DF_SHORT, Now) Then
        MsgBox ("현재일 보다 뒤의 일자를 선택 했습니다.")
        Exit Sub
    End If
    
    If Left(MakeDate(DF_SHORT, dtpDate(1)), 6) = Left(MakeDate(DF_SHORT, Now), 6) Then
        eDate = MakeDate(DF_SHORT, Now)
    Else
        eDate = GetLastDayMonth(sDate, ED_CUR)
    End If
    
    With grdHeader
    
        Set rs = oSubul.GetOutWareReportOrder(sDate, eDate, IIf(optOrder(0).Value = True, "0", "1"), txtOrderID)
        If rs.RecordCount > 0 Then
            nRow = 1
            .TextMatrix(nRow, 2) = rs!kCustom
            
            .TextMatrix(nRow, 4) = rs!OrderNo
            .TextMatrix(nRow, 6) = rs!OrderID
            
            .TextMatrix(nRow, 8) = SetCurrency(rs!UnitPrice, 2)
            
            nRow = 2
            .TextMatrix(nRow, 2) = rs!Article
            .TextMatrix(nRow, 4) = rs!WorkName
            .TextMatrix(nRow, 6) = SetCurrency(rs!OrderQty) & IIf(rs!UnitClss = "0", "  Y", "  M")
            .TextMatrix(nRow, 8) = SetCurrency(rs!ChunkRate, 2) & " + " & SetCurrency(rs!LossRate, 2)
            
            nRow = 3
            .TextMatrix(nRow, 2) = SetCurrency(rs!IGQty)
            .Cell(flexcpText, 3, 4, 3, 7) = "[전월누계:" & SetCurrency(rs!PreQty, 0) & "]   [당월입고:" & SetCurrency(rs!CurQty, 0) & _
                                            "]    [합계:" & SetCurrency(rs!TotQty, 0) & "]    [과부족:" & SetCurrency(rs!OvrQty, 0) & "]"
            .TextMatrix(nRow, 8) = IIf(rs!CloseClss = "*", "완결", "미완결")
        End If
    End With
    
    Exit Sub
ErrHandler:


End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGridSub()
    Dim nRows As Integer, II As Integer
    
    Call SetVSFlexGrid(grdStuffIN)
    With grdStuffIN
        .Rows = 2
        .Cols = 4
        
        .FixedRows = 2
        .FixedCols = 0

        .RowHeight(0) = 250
        .RowHeight(1) = 250

        nRows = 0
        .TextMatrix(nRows, 0) = "날짜"
        .TextMatrix(nRows, 1) = "입고수량"
        .TextMatrix(nRows, 2) = "입고수량"
        .TextMatrix(nRows, 3) = "입고처"

        nRows = 1
        .TextMatrix(nRows, 0) = "날짜":                 .ColWidth(0) = 1000:        .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(nRows, 1) = "절수":                 .ColWidth(1) = 1800:       .ColAlignment(1) = flexAlignRightCenter
        .TextMatrix(nRows, 2) = "수량":                 .ColWidth(2) = 1800:       .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(nRows, 3) = "입고처":               .ColWidth(3) = 2200:       .ColAlignment(3) = flexAlignCenterCenter

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True

        For II = 0 To .Cols - 1
            .MergeCol(II) = True
            .FixedAlignment(II) = flexAlignCenterCenter
        Next II

        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse

        .ScrollBars = flexScrollBarNone
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub InitGridHeader()
    Dim i%, nRows%, nCol%

    Call SetVSFlexGrid(grdHeader)
    With grdHeader
        .Rows = 4
        .Cols = 9
        
        .FixedRows = 4
        .FixedCols = 0
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        nRows = 1
        .RowHeight(nRows) = 350
        .TextMatrix(nRows, 0) = "":                     .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "거래처":               .ColWidth(1) = 1800:       .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = " ":                    .ColWidth(2) = 2200:       .ColAlignment(2) = flexAlignCenterCenter
        .Cell(flexcpBackColor, nRows, 2, nRows, 2) = vbWhite
        
        
        .TextMatrix(nRows, 3) = "OrderNO":              .ColWidth(3) = 1800:       .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = " ":                    .ColWidth(4) = 2200:       .ColAlignment(4) = flexAlignCenterCenter
        .Cell(flexcpBackColor, nRows, 4, nRows, 4) = vbWhite
        
        .TextMatrix(nRows, 5) = "관리번호":             .ColWidth(5) = 1800:       .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = " ":                    .ColWidth(6) = 2200:       .ColAlignment(6) = flexAlignCenterCenter
        .Cell(flexcpBackColor, nRows, 6, nRows, 6) = vbWhite
        
        .TextMatrix(nRows, 7) = "단가":                 .ColWidth(7) = 1800:       .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = " ":                    .ColWidth(8) = 2200:       .ColAlignment(8) = flexAlignCenterCenter
        .Cell(flexcpBackColor, nRows, 8, nRows, 8) = vbWhite


        nRows = 2
        .RowHeight(nRows) = 350
        .TextMatrix(nRows, 1) = "ITEM"
        .TextMatrix(nRows, 3) = "가공방법"
        .TextMatrix(nRows, 5) = "Order량"
        .TextMatrix(nRows, 7) = "축율(%)"
        
        .Cell(flexcpBackColor, nRows, 2, nRows, 2) = vbWhite
        .Cell(flexcpBackColor, nRows, 4, nRows, 4) = vbWhite
        .Cell(flexcpBackColor, nRows, 6, nRows, 6) = vbWhite
        .Cell(flexcpBackColor, nRows, 8, nRows, 8) = vbWhite
        
        
        nRows = 3
        .RowHeight(nRows) = 350
        .TextMatrix(nRows, 1) = "I/G 환산수량"
        .TextMatrix(nRows, 3) = "입고내역"
        .TextMatrix(nRows, 4) = " "
        .TextMatrix(nRows, 5) = " "
        .TextMatrix(nRows, 6) = " "
        .TextMatrix(nRows, 7) = "완결확인"
        .Cell(flexcpBackColor, nRows, 2, nRows, 2) = vbWhite
        .Cell(flexcpBackColor, nRows, 4, nRows, 6) = vbWhite
        .Cell(flexcpBackColor, nRows, 8, nRows, 8) = vbWhite

'        .Cell(flexcpPicture, 3, 3) = LoadResPicture("B_FIND", vbResBitmap)
'        .Cell(flexcpPictureAlignment, 3, 3) = flexAlignCenterCenter

        .MergeCells = flexMergeFree
        .MergeRow(1) = True
        .MergeRow(2) = True
        .MergeRow(3) = True
        
        .RowHidden(0) = True
        
'
'        For i = 0 To .FixedRows - 3
'            .RowHidden(i) = True
'        Next i
        
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse
        
        .ScrollBars = flexScrollBarNone
        '.GridColorFixed = vbWhite
'        .BackColorFixed = vbWhite
'        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite
        .ColDataType(3) = flexDTBoolean
        .Redraw = flexRDDirect

    End With
End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim i%, nRows%, nCol%
    Dim nDays As Integer
    
    nDays = 40

    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Rows = 5
        .Cols = 87
        
        .FixedRows = 5
        .FixedCols = 0
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250


        nRows = 2
        .TextMatrix(nRows, 0) = ""
        .TextMatrix(nRows, 1) = "Color"
        .TextMatrix(nRows, 2) = "수량"
        
        '  40 * 2
        nCol = 2
        For i = 1 To nDays
            nCol = nCol + 1
            .TextMatrix(nRows, nCol) = "출고내역"
            
            nCol = nCol + 1
            .TextMatrix(nRows, nCol) = "출고내역"
        Next i
        
        nCol = 83:        .TextMatrix(nRows, nCol) = "출고내역"
        nCol = 84:        .TextMatrix(nRows, nCol) = "출고내역"
        nCol = 85:        .TextMatrix(nRows, nCol) = "출고내역"
        nCol = 86:        .TextMatrix(nRows, nCol) = "과부족"


        nRows = 3
        .TextMatrix(nRows, 0) = ""
        .TextMatrix(nRows, 1) = "Color"
        .TextMatrix(nRows, 2) = "수량"
        
        nCol = 2
        For i = 1 To nDays
            nCol = nCol + 1
            If i <= 31 Then
                .TextMatrix(nRows, nCol) = CStr(i) & "일"
                
                nCol = nCol + 1
                .TextMatrix(nRows, nCol) = CStr(i) & "일"
            Else
                .TextMatrix(nRows, nCol) = " "

                nCol = nCol + 1
                .TextMatrix(nRows, nCol) = " "
            
            End If
        Next i
        
        nCol = 83:        .TextMatrix(nRows, nCol) = "출고수량"
        nCol = 84:        .TextMatrix(nRows, nCol) = "출고수량"
        nCol = 85:        .TextMatrix(nRows, nCol) = "출고수량"
        nCol = 86:        .TextMatrix(nRows, nCol) = "과부족"
        
        
        nRows = 4
        
        .TextMatrix(nRows, 0) = "":                  .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "Color":             .ColWidth(1) = 1600:     .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(nRows, 2) = "수량":              .ColWidth(2) = 1100:     .ColAlignment(2) = flexAlignRightCenter
        
        nCol = 2
        For i = 1 To nDays
            nCol = nCol + 1
            If i > 31 Then
                .TextMatrix(nRows, nCol) = " ":            .ColWidth(nCol) = 0:      .ColAlignment(nCol) = flexAlignRightCenter
    '            .TextMatrix(nRows, nCol) = "절수":            .ColWidth(nCol) = 800:      .ColAlignment(nCol) = flexAlignCenterCenter
                
                nCol = nCol + 1
                .TextMatrix(nRows, nCol) = " ":            .ColWidth(nCol) = 1100:     .ColAlignment(nCol) = flexAlignRightCenter
                
            Else
                .TextMatrix(nRows, nCol) = "절수":            .ColWidth(nCol) = 0:      .ColAlignment(nCol) = flexAlignRightCenter
    '            .TextMatrix(nRows, nCol) = "절수":            .ColWidth(nCol) = 800:      .ColAlignment(nCol) = flexAlignCenterCenter
                
                nCol = nCol + 1
                .TextMatrix(nRows, nCol) = "수량":            .ColWidth(nCol) = 1100:     .ColAlignment(nCol) = flexAlignRightCenter
            End If
        Next i
        
        nCol = 83
        .TextMatrix(nRows, nCol) = "전월누계":             .ColWidth(nCol) = 1200:       .ColAlignment(nCol) = flexAlignRightCenter

        nCol = 84
        .TextMatrix(nRows, nCol) = "당월출고":             .ColWidth(nCol) = 1200:       .ColAlignment(nCol) = flexAlignRightCenter

        nCol = 85
        .TextMatrix(nRows, nCol) = "당월누계":             .ColWidth(nCol) = 1200:       .ColAlignment(nCol) = flexAlignRightCenter

        nCol = 86
        .TextMatrix(nRows, nCol) = "과부족":               .ColWidth(nCol) = 1200:       .ColAlignment(nCol) = flexAlignRightCenter

        .MergeCells = flexMergeFixedOnly
        .MergeRow(2) = True
        .MergeRow(3) = True
        
        .MergeCol(1) = True
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        
        For i = 2 To .Cols - 1
            .MergeCol(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
'        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse
        
        .ScrollBars = flexScrollBarBoth
        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite
        .WordWrap = False
        .Redraw = flexRDDirect
        
    End With
    
End Sub

Private Sub FillGridData()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
    Dim i%, sOrderID$, bFlag As Boolean, II%, JJ%, nDay%
    Dim nBaseCol As Integer
    
    nBaseCol = 7    '기본으로 나타낼 날짜 컬럼 수( 절수 나타나지 않을 경우 )

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    

    Set rs = oSubul.GetOutWareReport(sDate, eDate, IIf(optOrder(0).Value = True, "0", "1"), txtOrderID)
    Set oSubul = Nothing
    With grdData(0)
        .Redraw = flexRDNone

        .Rows = .FixedRows
        
        If rs.RecordCount < 1 Then
            Screen.MousePointer = vbDefault
             .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
            cmdPrint.Visible = False

            Exit Sub
        End If
        
        cmdPrint.Visible = True
        Do Until rs.EOF
            .AddItem ""
            .RowHeight(.Rows - 1) = 400
            .TextMatrix(.Rows - 1, 1) = Trim(rs!Color)
            .TextMatrix(.Rows - 1, 2) = SetCurrency(rs!ColorQty, 0)
            
            .TextMatrix(.Rows - 1, 3) = SetCurrency(rs!T01_RollQty, 0)
            .TextMatrix(.Rows - 1, 4) = SetCurrency(rs!T01_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 5) = SetCurrency(rs!T02_RollQty, 0)
            .TextMatrix(.Rows - 1, 6) = SetCurrency(rs!T02_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 7) = SetCurrency(rs!T03_RollQty, 0)
            .TextMatrix(.Rows - 1, 8) = SetCurrency(rs!T03_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 9) = SetCurrency(rs!T04_RollQty, 0)
            .TextMatrix(.Rows - 1, 10) = SetCurrency(rs!T04_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 11) = SetCurrency(rs!T05_RollQty, 0)
            .TextMatrix(.Rows - 1, 12) = SetCurrency(rs!T05_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 13) = SetCurrency(rs!T06_RollQty, 0)
            .TextMatrix(.Rows - 1, 14) = SetCurrency(rs!T06_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 15) = SetCurrency(rs!T07_RollQty, 0)
            .TextMatrix(.Rows - 1, 16) = SetCurrency(rs!T07_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 17) = SetCurrency(rs!T08_RollQty, 0)
            .TextMatrix(.Rows - 1, 18) = SetCurrency(rs!T08_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 19) = SetCurrency(rs!T09_RollQty, 0)
            .TextMatrix(.Rows - 1, 20) = SetCurrency(rs!T09_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 21) = SetCurrency(rs!T10_RollQty, 0)
            .TextMatrix(.Rows - 1, 22) = SetCurrency(rs!T10_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 23) = SetCurrency(rs!T11_RollQty, 0)
            .TextMatrix(.Rows - 1, 24) = SetCurrency(rs!T11_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 25) = SetCurrency(rs!T12_RollQty, 0)
            .TextMatrix(.Rows - 1, 26) = SetCurrency(rs!T12_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 27) = SetCurrency(rs!T13_RollQty, 0)
            .TextMatrix(.Rows - 1, 28) = SetCurrency(rs!T13_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 29) = SetCurrency(rs!T14_RollQty, 0)
            .TextMatrix(.Rows - 1, 30) = SetCurrency(rs!T14_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 31) = SetCurrency(rs!T15_RollQty, 0)
            .TextMatrix(.Rows - 1, 32) = SetCurrency(rs!T15_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 33) = SetCurrency(rs!T16_RollQty, 0)
            .TextMatrix(.Rows - 1, 34) = SetCurrency(rs!T16_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 35) = SetCurrency(rs!T17_RollQty, 0)
            .TextMatrix(.Rows - 1, 36) = SetCurrency(rs!T17_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 37) = SetCurrency(rs!T18_RollQty, 0)
            .TextMatrix(.Rows - 1, 38) = SetCurrency(rs!T18_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 39) = SetCurrency(rs!T19_RollQty, 0)
            .TextMatrix(.Rows - 1, 40) = SetCurrency(rs!T19_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 41) = SetCurrency(rs!T20_RollQty, 0)
            .TextMatrix(.Rows - 1, 42) = SetCurrency(rs!T20_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 43) = SetCurrency(rs!T21_RollQty, 0)
            .TextMatrix(.Rows - 1, 44) = SetCurrency(rs!T21_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 45) = SetCurrency(rs!T22_RollQty, 0)
            .TextMatrix(.Rows - 1, 46) = SetCurrency(rs!T22_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 47) = SetCurrency(rs!T23_RollQty, 0)
            .TextMatrix(.Rows - 1, 48) = SetCurrency(rs!T23_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 49) = SetCurrency(rs!T24_RollQty, 0)
            .TextMatrix(.Rows - 1, 50) = SetCurrency(rs!T24_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 51) = SetCurrency(rs!T25_RollQty, 0)
            .TextMatrix(.Rows - 1, 52) = SetCurrency(rs!T25_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 53) = SetCurrency(rs!T26_RollQty, 0)
            .TextMatrix(.Rows - 1, 54) = SetCurrency(rs!T26_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 55) = SetCurrency(rs!T27_RollQty, 0)
            .TextMatrix(.Rows - 1, 56) = SetCurrency(rs!T27_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 57) = SetCurrency(rs!T28_RollQty, 0)
            .TextMatrix(.Rows - 1, 58) = SetCurrency(rs!T28_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 59) = SetCurrency(rs!T29_RollQty, 0)
            .TextMatrix(.Rows - 1, 60) = SetCurrency(rs!T29_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 61) = SetCurrency(rs!T30_RollQty, 0)
            .TextMatrix(.Rows - 1, 62) = SetCurrency(rs!T30_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 63) = SetCurrency(rs!T31_RollQty, 0)
            .TextMatrix(.Rows - 1, 64) = SetCurrency(rs!T31_OutQty, 0)
            
            .TextMatrix(.Rows - 1, 83) = SetCurrency(rs!Pre_OutQty, 0)
            .TextMatrix(.Rows - 1, 84) = SetCurrency(rs!Now_OutQty, 0)
            .TextMatrix(.Rows - 1, 85) = SetCurrency(rs!Cur_OutQty, 0)
            .TextMatrix(.Rows - 1, 86) = SetCurrency(rs!Ovr_OutQty, 0)
            
            If Trim(rs!Depth) <> "Z0" Then
                Call SetGrdColor(grdData(0), Right(rs!Depth, 1), .Rows - 1, 0, .Rows - 1, .Cols - 1)
            End If
            
            
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If
        
'        ' total의 값이 0이면 colHidden처리  1-> 31까지
'        JJ = 0: II = 0: nDay = 0
        For i = 4 To 65 Step 2
            If .ValueMatrix(.Rows - 1, i) = 0 And .ValueMatrix(.Rows - 2, i) = 0 And .ValueMatrix(.Rows - 3, i) = 0 Then
                .ColWidth(i - 1) = 0
                .ColWidth(i) = 0
            Else
                ' 값이 있는 날짜 수
                II = II + 1
            End If
        Next i
        
        ' 공란개수
        JJ = II
        For i = 66 To 82 Step 2
            JJ = JJ + 1
            If JJ > nBaseCol Then
                .ColWidth(i) = 0
                .ColWidth(i - 1) = 0
            End If
        Next i
        
'        ' 기본공란으로 보여질 컬럼 수 확인
'        If JJ < 9 Then
'            If 31 - nDay < 9 - JJ Then
'            Else
'
'            End If
'        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmSubulReport.FillGridData", Err.Description)
End Sub



Private Sub optOrder_Click(Index As Integer)
    If optOrder(0).Value = True Then
        cmdFind(0).Enabled = False
    Else
        cmdFind(0).Enabled = True
    End If
    pnlCaption(0).Caption = optOrder(Index).Caption
End Sub

Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call MoveFocus(KeyAscii)
    End If

End Sub
