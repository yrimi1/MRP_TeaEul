VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutwareResult 
   ClientHeight    =   9255
   ClientLeft      =   3660
   ClientTop       =   2820
   ClientWidth     =   11850
   Icon            =   "frmOutwareResult.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   420
      TabIndex        =   24
      Top             =   4020
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
         TabIndex        =   25
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
         TabIndex        =   26
         Top             =   120
         Width           =   270
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   7335
      Left            =   0
      TabIndex        =   29
      Top             =   960
      Width           =   3585
      _cx             =   6324
      _cy             =   12938
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   28
      Top             =   8400
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   1667
      _Version        =   196609
      Begin VB.TextBox txtExchangeRate 
         Height          =   315
         Left            =   9210
         TabIndex        =   31
         Top             =   540
         Visible         =   0   'False
         Width           =   1425
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   7920
         TabIndex        =   30
         Top             =   540
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "환율"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   2040
         MousePointer    =   99  '사용자 정의
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   510
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1410
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   510
         Width           =   600
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   5760
         TabIndex        =   4
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   5760
         TabIndex        =   3
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   9210
         TabIndex        =   2
         Top             =   180
         Width           =   1425
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10710
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   1
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   2715
         TabIndex        =   7
         Top             =   180
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23658497
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   2715
         TabIndex        =   8
         Top             =   555
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23658497
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   1410
         TabIndex        =   9
         Top             =   180
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
            Left            =   60
            TabIndex        =   10
            Top             =   30
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   4470
         TabIndex        =   11
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   7290
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   4470
         TabIndex        =   14
         Top             =   540
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
         Left            =   7290
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   540
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
         Left            =   7920
         TabIndex        =   17
         Top             =   180
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
            Width           =   1095
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
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   4005
         TabIndex        =   23
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4005
         TabIndex        =   22
         Top             =   270
         Width           =   360
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdColor 
      Height          =   7365
      Left            =   3600
      TabIndex        =   27
      Top             =   960
      Width           =   8205
      _cx             =   14473
      _cy             =   12991
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
Attribute VB_Name = "frmOutwareResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bLoading As Boolean

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else
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

Private Sub cmdSearch_Click()
    Call FillGridOrder
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
    
    pnlProgress.Visible = False
End Sub

Private Sub grdOrder_RowColChange()
    If m_bLoading Then Exit Sub
    
    Call FillGridColor
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdOrder
        If optOrder(0).Value Then
            .ColWidth(3) = 1350
            .ColWidth(2) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(3) = 0
            .ColWidth(2) = 1350
            chkSearch(3).Caption = "관리번호"
        End If
    End With
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
    
    With grdOrder
        .Redraw = flexRDNone
        .Cols = 6
        Call SetVSFlexGrid(grdOrder)

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = ""
        .TextArray(1) = "":               .ColWidth(1) = 0:          .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "관리번호":       .ColWidth(2) = 1350:       .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "Order No.":      .ColWidth(3) = 0:          .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처":         .ColWidth(4) = 1800:       .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "품명":           .ColWidth(5) = 1700:       .ColAlignment(5) = flexAlignLeftCenter
                
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
    End With

    With grdColor
        .Redraw = flexRDNone
        .Cols = 9
        Call SetVSFlexGrid(grdColor)

        .Rows = 4
        .FixedRows = 4
        .FixedCols = 1

        .RowHeightMin = 300
        .RowHeight(3) = 400
        
        .TextMatrix(3, 0) = ""
        .TextMatrix(3, 1) = "색상명":      .ColWidth(1) = 1500:       .ColAlignment(1) = flexAlignLeftTop
        .TextMatrix(3, 2) = "오더량":      .ColWidth(2) = 900:        .ColAlignment(2) = flexAlignRightTop
        .TextMatrix(3, 3) = "출고일자":    .ColWidth(3) = 1000:       .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(3, 4) = "출고절수":    .ColWidth(4) = 900:        .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(3, 5) = "출고수량":    .ColWidth(5) = 900:        .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(3, 6) = "누계절수":    .ColWidth(6) = 900:        .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(3, 7) = "누계수량":    .ColWidth(7) = 900:        .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(3, 8) = "과부족":      .ColWidth(8) = 900:        .ColAlignment(8) = flexAlignRightCenter
        
        .ColFormat(2) = "#,##0"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0"
        .ColFormat(6) = "#,##0"
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
                
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .MergeCells = flexMergeFree
        
        For i = 0 To 2
            .MergeCol(i) = True
        Next i
        
        .ExplorerBar = flexExNone
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGridOrder()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    m_bLoading = True
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOutwareOrder(IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)) _
                                    , IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag _
                                    , IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag _
                                    , IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3))
    Set oOutware = Nothing
        
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            .AddItem CStr(.Rows) & vbTab & False & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                rs!kCustom & vbTab & rs!Article
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            Call FillGridColor
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    m_bLoading = False
    Exit Sub

ErrHandler:
    m_bLoading = True
    pnlProgress.Visible = False
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareResult.FillGridOrder", Err.Description)
End Sub

Private Sub FillGridColor()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, nOrderSeq%, sResultDate$
    Dim nTOutRoll#, nTOutQty#
    
    On Error GoTo ErrHandler
           
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOutwareOrderDetail(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE))
    Set oOutware = Nothing
        
    With grdColor
        .Redraw = flexRDDirect
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            If rs!ResultDate = "Z" Or rs!ResultDate = "ZZ" Then
                .AddItem ""
                .RowHidden(.Rows - 1) = True
            End If
            
            .AddItem CStr(i + 1)
            .TextMatrix(.Rows - 1, 1) = Trim(rs!Color)
            .TextMatrix(.Rows - 1, 2) = Format(rs!ColorQty, "#,###") & IIf(rs!UnitClss = "0", " Y", " M")

            .TextMatrix(.Rows - 1, 3) = MakeDate(DF_LONG, rs!ResultDate)
            .TextMatrix(.Rows - 1, 4) = rs!OutRoll
            .TextMatrix(.Rows - 1, 5) = rs!OutQty
            
            If rs!ResultDate = "Z" Or rs!ResultDate = "ZZ" Then
                .TextMatrix(.Rows - 1, 6) = rs!OutRoll
                .TextMatrix(.Rows - 1, 7) = rs!OutQty
                .TextMatrix(.Rows - 1, 8) = 0
                
                nTOutRoll = 0
                nTOutQty = 0
            Else
                .TextMatrix(.Rows - 1, 6) = nTOutRoll + rs!OutRoll
                .TextMatrix(.Rows - 1, 7) = nTOutQty + rs!OutQty
                .TextMatrix(.Rows - 1, 8) = 0
                
                nTOutRoll = nTOutRoll + rs!OutRoll
                nTOutQty = nTOutQty + rs!OutQty
            End If
            
            If rs!ResultDate = "Z" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
            ElseIf rs!ResultDate = "ZZ" Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0C0C0
            End If
            nOrderSeq = rs!OrderSeq
            sResultDate = rs!ResultDate
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
                
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    Exit Sub

ErrHandler:
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareResult.FillGridColor", Err.Description)
End Sub




