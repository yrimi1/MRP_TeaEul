VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutwareDetail 
   ClientHeight    =   9255
   ClientLeft      =   180
   ClientTop       =   630
   ClientWidth     =   11850
   Icon            =   "frmOutwareDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   450
      TabIndex        =   27
      Top             =   3810
      Visible         =   0   'False
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1535
      _Version        =   196609
      Alignment       =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin MSComctlLib.ProgressBar proProgress 
         Height          =   390
         Left            =   90
         TabIndex        =   28
         Top             =   375
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "180"
         Height          =   180
         Left            =   195
         TabIndex        =   29
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   25
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7365
      Left            =   0
      TabIndex        =   24
      Top             =   1140
      Width           =   11835
      _cx             =   20876
      _cy             =   12991
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
   Begin Threed.SSFrame frmSearch 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1931
      _Version        =   196609
      Begin VB.ComboBox cboTaxClss 
         Height          =   300
         Left            =   9540
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   750
         Width           =   1245
      End
      Begin VB.ComboBox cboOutClss 
         Height          =   300
         Left            =   9540
         Style           =   2  '드롭다운 목록
         TabIndex        =   32
         Top             =   420
         Width           =   1245
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10950
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   6
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   9540
         TabIndex        =   5
         Top             =   75
         Width           =   1245
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   6360
         TabIndex        =   4
         Top             =   420
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   6360
         TabIndex        =   3
         Top             =   75
         Width           =   1485
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   75
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   465
         Width           =   600
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3435
         TabIndex        =   7
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116719617
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3435
         TabIndex        =   8
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116719617
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   2130
         TabIndex        =   9
         Top             =   75
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
            Caption         =   "출고 일자"
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   11
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   12
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   7890
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   75
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   5160
         TabIndex        =   14
         Top             =   420
         Width           =   1125
         _ExtentX        =   1984
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
            Caption         =   "품     명"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   7890
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   420
         Width           =   300
         _ExtentX        =   529
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
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   8310
         TabIndex        =   17
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
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
            Width           =   1035
         End
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   735
         Left            =   90
         TabIndex        =   19
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1296
         _Version        =   196609
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   8310
         TabIndex        =   30
         Top             =   420
         Width           =   1185
         _ExtentX        =   2090
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
            Caption         =   "출고구분"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   1035
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4755
         TabIndex        =   23
         Top             =   135
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   4755
         TabIndex        =   22
         Top             =   420
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8340
      TabIndex        =   26
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmOutwareDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_ROW = 23
Private Const LIMIT_WIDTH = 1000


Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    ElseIf Index = 1 Or Index = 2 Or Index = 3 Then
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    Else
        If chkSearch(Index).Value = vbChecked Then
            cboOutClss.Enabled = True
        Else
            cboOutClss.Enabled = False
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(2))
    End If
End Sub

Private Sub cmdPrint_Click()
'    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
        Call FillGridPrint
'    End If
End Sub

'''Sub ColResize(ByVal pType As String)
'''    Dim II%, JJ As Integer
'''
'''    If pType = "-" Then
'''        JJ = -1
'''    Else
'''        JJ = 1
'''    End If
'''
'''    With grdData
'''        .Redraw = flexRDBuffered
''''        .ColWidth(0) = .ColWidth(0) + 360 * JJ
'''        .ColWidth(1) = .ColWidth(1) + 50 * JJ
'''        .ColWidth(2) = .ColWidth(2) + 600 * JJ
'''        .ColWidth(3) = .ColWidth(3) + 400 * JJ
'''        .ColWidth(4) = .ColWidth(4) + 200 * JJ
'''        .ColWidth(5) = .ColWidth(5) + 100 * JJ
'''        .ColWidth(6) = .ColWidth(6) + 100 * JJ
'''        .ColWidth(7) = .ColWidth(7) + 150 * JJ
'''        .ColWidth(8) = .ColWidth(8) + 150 * JJ
'''        .ColWidth(9) = .ColWidth(9) + 50 * JJ
'''        .ColWidth(10) = .ColWidth(10) + 10 * JJ
'''
'''        .Redraw = flexRDDirect
'''    End With
'''
'''End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
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
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    Call InitGrid

    For i = 0 To 1
        dtpDate(i) = Now
    Next i
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    
    With cboOutClss
        .AddItem "1. 정상출고":         .ItemData(0) = 1
        .AddItem "2. 제직불량":         .ItemData(1) = 2
        .AddItem "3. 가공불량":         .ItemData(2) = 3
        .AddItem "4. Sample, 시가공":   .ItemData(3) = 4
        .AddItem "5. 정산분":           .ItemData(4) = 5
        .AddItem "6. 전체":             .ItemData(5) = 10
        .AddItem "7. 가공불량":         .ItemData(6) = 11
        .AddItem "8. 염색불량":         .ItemData(7) = 12
        .AddItem "9. 검사불량":         .ItemData(8) = 13
        .AddItem "10. 염색수정":        .ItemData(9) = 14
        .AddItem "11. 가공수정":        .ItemData(10) = 15
        .AddItem "12. 검사수정":        .ItemData(11) = 16
        .AddItem "13. 재포장":          .ItemData(12) = 17
        .AddItem "14. 기타":            .ItemData(13) = 18
        
        .ListIndex = 0
        .Enabled = False
    End With
    
    With cboTaxClss
        .AddItem "9. 전체"
        .AddItem "0. 비사용"
        .AddItem "1. 사용"
        .ListIndex = 0
    End With
    
    
    pnlProgress.Visible = False
End Sub

Private Sub optOrder_Click(Index As Integer)
'    With grdData
        If optOrder(0).Value Then
'            .ColWidth(4) = 1350
'            .ColWidth(3) = 0
            chkSearch(3).Caption = "Order No."
        Else
'            .ColWidth(4) = 0
'            .ColWidth(3) = 1350
            chkSearch(3).Caption = "관리번호"
        End If
'    End With
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
        Call NextFocus
    ElseIf KeyAscii = vbKeyReturn And Index = 3 Then
        Call NextFocus
    End If
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Redraw = flexRDNone
        .Cols = 12
        
        Call SetVSFlexGrid(grdData)
        
        .Rows = 4
        .FixedCols = 0
        .FixedRows = 4
        
        .RowHeightMin = 270
        .RowHeight(3) = 400
        
        .TextMatrix(3, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(3, 1) = "월일":             .ColWidth(1) = 600:                 .ColAlignment(1) = flexAlignCenterTop
        .TextMatrix(3, 2) = "거 래 처":         .ColWidth(2) = 1450:                .ColAlignment(2) = flexAlignLeftTop
        .TextMatrix(3, 3) = "Order No":         .ColWidth(3) = 1200:                .ColAlignment(3) = flexAlignLeftTop
        .TextMatrix(3, 4) = "관리":             .ColWidth(4) = 500:                 .ColAlignment(4) = flexAlignCenterTop
        .TextMatrix(3, 5) = "실출고처":         .ColWidth(5) = 1250:                .ColAlignment(5) = flexAlignLeftTop
        .TextMatrix(3, 6) = "품    명":         .ColWidth(6) = 1650:                .ColAlignment(6) = flexAlignLeftTop
        .TextMatrix(3, 7) = "가공":             .ColWidth(7) = 1000:                .ColAlignment(7) = flexAlignLeftTop
        .TextMatrix(3, 8) = "COLOR 명":         .ColWidth(8) = 1200:                .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "절 수":            .ColWidth(9) = 800:                 .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "출고량":          .ColWidth(10) = 1000:               .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(3, 11) = "":                .ColWidth(11) = 0:                  .ColAlignment(11) = flexAlignCenterCenter
          
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .ColHidden(0) = True
        .ColHidden(11) = True
        
        .ColFormat(8) = "#,###"
        .ColFormat(9) = "#,###"
        .ColFormat(10) = "#,###"
        
        .MergeCells = flexMergeFree
        For i = 0 To 7
            .MergeCol(i) = True
        Next i

        .ScrollBars = flexScrollBarVertical
        .SelectionMode = flexSelectionListBox
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGridData()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, sOrderID$
    Dim sOutQty$
    Dim sCustom$, sOrderNO$, sOutCustom$, sArticle$, sWorkName$, sColor$, nOutSeq%
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOutwareDetail(IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                 IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), cboOutClss.ItemData(cboOutClss.ListIndex), Left(cboTaxClss, 1))
    Set oOutware = Nothing
        
    With grdData
        .Redraw = flexRDDirect
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            
            Select Case rs!UnitClss
                Case 0:                 sOutQty = Format(rs!OutQty, "#,###")
                Case 1:                 sOutQty = Format(rs!OutQty, "#,###") & " M"
                Case Else:              sOutQty = Format(rs!OutQty, "#,###")
            End Select
            
            If rs!Depth = "0" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = MakeDate(DF_MD, rs!ResultDate)
                .TextMatrix(.Rows - 1, 2) = rs!kCustom
                .TextMatrix(.Rows - 1, 3) = rs!OrderNo
                .TextMatrix(.Rows - 1, 4) = Right(MakeOrderID(rs!OrderID, OM_EXPAND), 4)
                .TextMatrix(.Rows - 1, 5) = IIf(rs!OutCustom = "", " ", rs!OutCustom)
                .TextMatrix(.Rows - 1, 6) = rs!Article
                .TextMatrix(.Rows - 1, 7) = rs!WorkName
                .TextMatrix(.Rows - 1, 8) = rs!Color
                .TextMatrix(.Rows - 1, 9) = rs!OutRoll
                .TextMatrix(.Rows - 1, 10) = rs!OutQty
                
            ElseIf rs!Depth = "1" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = MakeDate(DF_MD, rs!ResultDate)
                .TextMatrix(.Rows - 1, 2) = rs!kCustom
                .TextMatrix(.Rows - 1, 3) = rs!OrderNo
                .TextMatrix(.Rows - 1, 4) = Right(MakeOrderID(rs!OrderID, OM_EXPAND), 4)
                .TextMatrix(.Rows - 1, 5) = .TextMatrix(.Rows - 2, 5)
                .TextMatrix(.Rows - 1, 6) = rs!Article
                .TextMatrix(.Rows - 1, 7) = rs!WorkName
                .TextMatrix(.Rows - 1, 8) = "ORDER 계"
                .TextMatrix(.Rows - 1, 9) = rs!OutRoll
                .TextMatrix(.Rows - 1, 10) = sOutQty
                
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE9E9E9
                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = MakeDate(DF_MD, rs!ResultDate)
                .TextMatrix(.Rows - 1, 2) = rs!kCustom
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = ""
                .TextMatrix(.Rows - 1, 10) = ""
                
                .RowHidden(.Rows - 1) = True
            ElseIf rs!Depth = "2" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = MakeDate(DF_MD, rs!ResultDate)
                .TextMatrix(.Rows - 1, 2) = rs!kCustom
                .TextMatrix(.Rows - 1, 3) = "소     계"
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = rs!OutRoll
                .TextMatrix(.Rows - 1, 10) = rs!OutQtyY
                
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE5E5E5
                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            ElseIf rs!Depth = "3" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = MakeDate(DF_MD, rs!ResultDate)
                .TextMatrix(.Rows - 1, 2) = "일 자 계"
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = rs!OutRoll
                .TextMatrix(.Rows - 1, 10) = rs!OutQtyY

                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            ElseIf rs!Depth = "4" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = "기 간 계"
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = rs!OutRoll
                .TextMatrix(.Rows - 1, 10) = rs!OutQtyY

                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0C0C0
                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            ElseIf rs!Depth = "5" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = "월    계"
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = rs!OutRoll
                .TextMatrix(.Rows - 1, 10) = rs!OutQtyY

                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HD5D6D1
                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            End If
            
            .TextMatrix(.Rows - 1, 11) = rs!Depth
            
''            If rs!Depth <> "0" Then
''                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
''            End If
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
    
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
'    Call ChangeScroll
    pnlProgress.Visible = False
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareDetail.FillGridData", Err.Description)
End Sub

Sub FillGridReSET()
    Dim II As Integer
    
    With grdData
            
        For II = .FixedRows To .Rows - 1
            Select Case .TextMatrix(II, 11)
                Case "1"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE9E9E9
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "2"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE5E5E5
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "3"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "4"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0C0C0
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                Case "5"
                    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HD5D6D1
                    .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            End Select
        Next II
    End With
    
End Sub
Sub FillGridPrint()
    Dim i%
    Dim sDate As String, eDate As String, nPageHV As Integer
    
    If chkSearch(0).Value Then
        sDate = Format(dtpDate(0), "YYYY/MM/DD")
        eDate = Format(dtpDate(1), "YYYY/MM/DD")
    Else
        sDate = ""
        eDate = ""
    End If
    
    With grdData
        .Redraw = flexRDNone
        .ExtendLastCol = False
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
        Call SetPrintMode(grdData, 1, True, nPageHV)

        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "가공지 출고 명세서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 2, 1, 2, 5) = "▶ 출고일자 : " & sDate & " ~ " & eDate
    '    .Cell(flexcpText, 1, .Cols - 4, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD")
        
        .ColWidth(2) = 1400
        .ColWidth(6) = 1550
        .ColWidth(7) = 800
        
        For i = .FixedRows To .Rows - 1
            ' 일계, 총계의 금액은 BackColor을 설정 한다.
            If (.TextMatrix(i, 11) = "3" Or .TextMatrix(i, 11) = "4" Or .TextMatrix(i, 11) = "5") Then
                .Cell(flexcpBackColor, i, 9, i, 10) = PRNHeaderColor
            End If
        Next i
        
        
        .PrintGrid "태을염직", True, 1, 0, 1400
        
        Call SetPrintMode(grdData, 1, False)
        
        .ColWidth(2) = 1450
        .ColWidth(6) = 1650
        .ColWidth(7) = 1000
                
        .ExtendLastCol = True
        Call FillGridReSET

        .Redraw = flexRDDirect
    End With
End Sub

'Private Sub ChangeScroll()
'    With grdData
'        .ColWidth(10) = LIMIT_WIDTH - IIf(.Rows > LIMIT_ROW, 240, 0)
'    End With
'End Sub



