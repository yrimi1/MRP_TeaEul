VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOutwareLot 
   ClientHeight    =   9255
   ClientLeft      =   420
   ClientTop       =   630
   ClientWidth     =   11865
   Icon            =   "frmOutwareLot.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.Frame fraSearch 
      Height          =   2265
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   3530
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   300
         Index           =   1
         Left            =   75
         MousePointer    =   99  '사용자 정의
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   495
         Width           =   510
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   300
         Index           =   2
         Left            =   75
         MousePointer    =   99  '사용자 정의
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   825
         Width           =   510
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   12
         Top             =   1185
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   765
         Left            =   2685
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   165
         Width           =   765
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1380
         TabIndex        =   10
         Top             =   1875
         Width           =   1515
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1380
         TabIndex        =   9
         Top             =   1530
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   630
         TabIndex        =   15
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113639425
         CurrentDate     =   36271
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   630
         TabIndex        =   16
         Top             =   825
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113639425
         CurrentDate     =   36271
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   17
         Top             =   165
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "출고일자"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   18
            Top             =   60
            Value           =   1  '확인
            Width           =   1050
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   19
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거  래  처"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   20
            Top             =   60
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   2
         Left            =   75
         TabIndex        =   21
         Top             =   1875
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   22
            Top             =   60
            Width           =   1125
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   2940
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1185
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         Enabled         =   0   'False
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   22
         Left            =   75
         TabIndex        =   24
         Top             =   1530
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품       명"
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   25
            Top             =   60
            Width           =   1125
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   2940
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1530
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         Enabled         =   0   'False
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   1
         Left            =   1950
         TabIndex        =   28
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   0
         Left            =   1950
         TabIndex        =   27
         Top             =   885
         Width           =   360
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   6210
      Left            =   0
      TabIndex        =   0
      Top             =   2220
      Width           =   3495
      _cx             =   6165
      _cy             =   10954
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
   Begin VB.Frame fraOrder 
      Height          =   795
      Left            =   30
      TabIndex        =   5
      Top             =   8430
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1200
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   495
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin Threed.SSPanel pnlRollNo 
      Height          =   8430
      Left            =   3540
      TabIndex        =   4
      Top             =   30
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14870
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   7950
         Left            =   30
         TabIndex        =   6
         Top             =   45
         Width           =   8235
         _cx             =   14526
         _cy             =   14023
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
      Begin VSFlex7LCtl.VSFlexGrid grdColorTotal 
         Height          =   360
         Left            =   0
         TabIndex        =   7
         Top             =   8040
         Width           =   8265
         _cx             =   14579
         _cy             =   635
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
         FixedRows       =   0
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   3
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8340
      TabIndex        =   29
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
Attribute VB_Name = "frmOutwareLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bloading As Boolean


Private Sub cmdPrint_Click()
    Call MakeExcelGrid(grdColor)
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus

End Sub

Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 11970, 9660

    Call SetOperate(Me)
    Me.Show

    Call InitGrid

    dtpDate(0) = Now
    dtpDate(1) = Now
    
    For i = 1 To 2
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
    txtSearch(1).Enabled = False
    txtSearch(2).Enabled = False
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(Index) Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True

            dtpDate(0).SetFocus
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False

            cmdSearch.SetFocus
        End If
    Else
        If chkSearch(Index) Then
            If Index < 3 Then
                cmdFind(Index).Enabled = True
                cmdFind(Index).Enabled = True
            End If
            txtSearch(Index).Enabled = True
    
            txtSearch(Index).SetFocus
        Else
            If Index < 3 Then
                cmdFind(Index).Enabled = False
                cmdFind(Index).Enabled = False
            End If
            txtSearch(Index).Enabled = False
    
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(1))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, 0, False, txtSearch(2))
        End If
    End If
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

Private Sub grdOrder_RowColChange()
    If m_bloading Then Exit Sub

    Call FillGridColor
End Sub


Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkSearch(3).Caption = "Order No"
    Else
        chkSearch(3).Caption = "관리번호"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%

    With grdOrder
        .Cols = 6
        Call SetVSFlexGrid(grdOrder)

        .Redraw = flexRDNone

        .TextArray(0) = " "
        .TextArray(1) = "Order No":     .ColWidth(1) = 0:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "관리번호":     .ColWidth(2) = 1490:       .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "거래처명":     .ColWidth(3) = 15:      .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "거래처":       .ColWidth(4) = 0
        .TextArray(5) = "단위":         .ColWidth(5) = 0

        .Redraw = flexRDDirect
    End With

    With grdColor
        .Cols = 6
        Call SetVSFlexGrid(grdColor)

        .Redraw = flexRDNone
        .FixedCols = 0
        
        .TextArray(0) = ""
        .TextArray(1) = "":             .ColWidth(1) = 250
        .TextArray(2) = "색상명":       .ColWidth(2) = 3000:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "수주수량":     .ColWidth(3) = 1500:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "출고절수":     .ColWidth(4) = 1500:            .ColAlignment(4) = flexAlignRightCenter:    .ColFormat(4) = GetFormat(g_nPointPos)
        .TextArray(5) = "출고수량":     .ColWidth(5) = 1500:            .ColAlignment(5) = flexAlignRightCenter:    .ColFormat(5) = GetFormat(g_nPointPos)

        .GridLines = flexGridNone
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 450
        .RowHeightMin = 275


        .Redraw = flexRDDirect
    End With
    
    With grdColorTotal
        .Cols = 4
        Call SetVSFlexGrid(grdColorTotal)

        .Redraw = flexRDNone

        .FixedCols = 1
        .FixedRows = 0
        .Rows = 1

        .RowHeight(0) = 300
        .ScrollBars = flexScrollBarNone

        .TextArray(0) = "합          계":   .ColWidth(0) = 3000 + 610:  .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = " ":                .ColWidth(1) = 1500:                .ColAlignment(1) = flexAlignRightCenter:    .ColFormat(1) = "#,###"
        .TextArray(2) = " ":                .ColWidth(2) = 1500:                .ColAlignment(2) = flexAlignRightCenter:    .ColFormat(2) = "#,###"
        .TextArray(3) = " ":                .ColWidth(3) = 1500:                .ColAlignment(3) = flexAlignRightCenter:    .ColFormat(3) = "#,###"
        
         .Redraw = flexRDDirect
    End With

End Sub

Private Sub FillGridOrder()
    Dim oOutware As PlusLib2.COutWare
    Dim rs       As Recordset
    Dim i%, iNowRow%

    On Error GoTo ErrHandler

    m_bloading = True

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon

    Set rs = oOutware.GetOutwareOrder(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1), 1, 0), txtSearch(1).Tag, _
        IIf(chkSearch(2), 1, 0), txtSearch(2).Tag, _
        IIf(chkSearch(3), IIf(optOrder(0), 2, 1), 0), IIf(optOrder(0), txtSearch(3), MakeOrderID(txtSearch(3), OM_REDUCE)))
    Set oOutware = Nothing

    With grdOrder
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!kCustom

            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        .Redraw = flexRDDirect
        .SetFocus
    End With

    m_bloading = False

    Call FillGridColor

    Exit Sub

ErrHandler:
    m_bloading = False

    Set rs = Nothing
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGridColor()
    Dim oOutware As PlusLib2.COutWare
    Dim rs       As Recordset
    Dim i%, nTotal(2) As Long, nTop As Integer
    Dim nOrderSeq%
    
    If grdOrder.Rows = grdOrder.FixedRows Then
        grdColor.Rows = grdColor.FixedRows
        grdColor.HighLight = flexHighlightNever

        Exit Sub
    End If

    On Error GoTo ErrHandler
    m_bloading = True

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon

    Set rs = oOutware.GetOutWareOrderByLot(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 2), OM_REDUCE))
    Set oOutware = Nothing

    nOrderSeq = -1
    With grdColor
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            If rs!OrderSeq <> nOrderSeq Then
                .AddItem CStr(i) & vbTab & "" & vbTab & rs!Color & vbTab & SetCurrency(rs!ColorQty) & vbTab & 0 & vbTab & _
                    0 & vbTab & 0

                Call DoFlexGridGroup(grdColor, .Rows - 1, 0)
                Call GridCollapse(nTop)
                
                nTotal(0) = nTotal(0) + IIf(IsNull(rs!ColorQty), 0, rs!ColorQty)
                nTop = .Rows - 1
            End If
            
            .AddItem "" & vbTab & "" & vbTab & CStr(CheckNull(rs!LotNo)) & vbTab & "" & vbTab & _
                CheckNum(rs!OutRoll) & vbTab & CheckNum(rs!OutQty)
            
            .TextMatrix(nTop, 4) = .TextMatrix(nTop, 4) + CheckNum(rs!OutRoll)
            .TextMatrix(nTop, 5) = .TextMatrix(nTop, 5) + CheckNum(rs!OutQty)
            

            nTotal(1) = nTotal(1) + CheckNum(rs!OutRoll)
            nTotal(2) = nTotal(2) + CheckNum(rs!OutQty)

            nOrderSeq = rs!OrderSeq
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        Call GridCollapse(nTop)

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways

            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If

        .Redraw = flexRDDirect
    End With

    With grdColorTotal
        .TextMatrix(0, 1) = nTotal(0)
        .TextMatrix(0, 2) = nTotal(1)
        .TextMatrix(0, 3) = nTotal(2)
    End With

    m_bloading = False

    Exit Sub

ErrHandler:

    Set rs = Nothing
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(iRow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = &H0&        '&HE0E0E0
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        End Select
    End With
End Sub

Private Sub grdColor_DblClick()
    With grdColor
        If .Row < 1 Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub

Private Sub GridCollapse(Row As Integer)
   
    With grdColor
    
        If Row >= .FixedRows Then
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub

