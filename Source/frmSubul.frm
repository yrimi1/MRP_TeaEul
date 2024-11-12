VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSubul 
   ClientHeight    =   9435
   ClientLeft      =   1665
   ClientTop       =   1530
   ClientWidth     =   15240
   Icon            =   "frmSubul.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15240
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7515
      Left            =   30
      TabIndex        =   20
      Top             =   990
      Width           =   15165
      _cx             =   26749
      _cy             =   13256
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
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1667
      _Version        =   196609
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   465
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   600
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   4
         Top             =   435
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   7170
         TabIndex        =   3
         Top             =   435
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   9120
         TabIndex        =   2
         Top             =   435
         Width           =   1485
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   14310
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
         Left            =   3450
         TabIndex        =   7
         Top             =   75
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   529
         _Version        =   393216
         Format          =   53936129
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   2130
         TabIndex        =   8
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
         Caption         =   "수불 월"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   5160
         TabIndex        =   9
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   10
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   6690
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   435
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
         Left            =   7140
         TabIndex        =   12
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
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
            TabIndex        =   13
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   8700
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   435
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
         Left            =   9120
         TabIndex        =   15
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
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
            TabIndex        =   16
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   735
         Left            =   90
         TabIndex        =   17
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1170
         End
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   11790
      TabIndex        =   21
      Tag             =   "PERM_ADDNEW"
      Top             =   8640
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      저장(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13530
      TabIndex        =   22
      Top             =   8640
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmSubul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bloading As Boolean

Private Sub cmdSave_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    If SaveData Then
        Call FillGridData
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
    
    dtpDate(0) = Now
'    dtpDate(1) = Now
    Call InitGrid

    For i = 1 To 2
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    
    cmdSave.Picture = LoadResPicture("CHECK", vbResIcon)
    
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If chkSearch(Index) Then
        If Index = 1 Or Index = 2 Then
            cmdFind(Index).Enabled = True
        End If
        txtSearch(Index).Enabled = True
        txtSearch(Index).SetFocus
    Else
        If Index = 1 Or Index = 2 Then
            cmdFind(Index).Enabled = False
        End If
        txtSearch(Index).Enabled = False
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With grdData
            If IsNumeric(.TextMatrix(Row, Col)) Then
                If Col = 20 Then
                    .Col = .Col + 1
                    
                    .TextMatrix(Row, 16) = CSng(.TextMatrix(Row, 10)) + CSng(.TextMatrix(Row, 12)) - CSng(.TextMatrix(Row, 14)) + CSng(.TextMatrix(Row, 18)) + CSng(CheckNum(.TextMatrix(Row, 20)))
                    .TextMatrix(Row, 22) = "*"
                ElseIf Col = 21 Then
                    .TextMatrix(Row, 17) = CSng(.TextMatrix(Row, 11)) + CSng(.TextMatrix(Row, 13)) - CSng(.TextMatrix(Row, 15)) + CSng(.TextMatrix(Row, 19)) + CSng(CheckNum(.TextMatrix(Row, 21)))
                    .TextMatrix(Row, 22) = "*"
                    If .Row < .Rows - 1 Then
                        .Col = 20
                        .Row = .Row + 1
                    End If
                End If
            Else
                .TextMatrix(Row, Col) = "0"
            End If
        End With
End Sub

Private Sub grdData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdData
        If Col < 20 Then
            .FocusRect = flexFocusNone
            .SelectionMode = flexSelectionByRow
            Cancel = True
        Else
            .FocusRect = flexFocusHeavy
            .SelectionMode = flexSelectionFree
        End If
    End With
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
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    End If
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        grdData.ColWidth(3) = 0
        grdData.ColWidth(4) = 1350
        chkSearch(3).Caption = "Order No"
    Else
        grdData.ColWidth(3) = 1350
        grdData.ColWidth(4) = 0
        chkSearch(3).Caption = "관리번호"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%

    With grdData
        .Cols = 26
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 2
        .FixedRows = 2
        .FixedCols = 1
'        .FrozenCols = 9
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        .TextArray(0) = " "
        .TextArray(1) = "거래처":       .ColWidth(1) = 1100:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "품명":         .ColWidth(2) = 1600:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호":     .ColWidth(3) = 1350:    .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "OrderNo":      .ColWidth(4) = 0:       .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "가공구분":     .ColWidth(5) = 1000:     .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "원단폭":       .ColWidth(6) = 700:     .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "축율":         .ColWidth(7) = 0:     .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "수주량":       .ColWidth(8) = 800:     .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "단위":         .ColWidth(9) = 500:     .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(10) = "전기이월":    .ColWidth(10) = 700:   .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "전기이월":    .ColWidth(11) = 860:   .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "입고":        .ColWidth(12) = 700:   .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "입고":        .ColWidth(13) = 860:   .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "출고":        .ColWidth(14) = 700:   .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "출고":        .ColWidth(15) = 860:   .ColAlignment(15) = flexAlignRightCenter
        .TextArray(16) = "재고":        .ColWidth(16) = 700:   .ColAlignment(16) = flexAlignRightCenter
        .TextArray(17) = "재고":        .ColWidth(17) = 860:   .ColAlignment(17) = flexAlignRightCenter
        .TextArray(18) = "기조정":      .ColWidth(18) = 0:     .ColAlignment(18) = flexAlignRightCenter
        .TextArray(19) = "기조정":      .ColWidth(19) = 0:     .ColAlignment(19) = flexAlignRightCenter
        .TextArray(20) = "조정":        .ColWidth(20) = 600:   .ColAlignment(20) = flexAlignRightCenter
        .TextArray(21) = "조정":        .ColWidth(21) = 860:   .ColAlignment(21) = flexAlignRightCenter
        .TextArray(22) = "수정구분":    .ColWidth(22) = 0
        .TextArray(23) = "거래처코드":  .ColWidth(23) = 0
        .TextArray(24) = "품명코드":    .ColWidth(24) = 0
        .TextArray(25) = "가공구분":    .ColWidth(25) = 0
        
        .TextArray(.Cols + 0) = " "
        .TextArray(.Cols + 1) = "거래처"
        .TextArray(.Cols + 2) = "품명"
        .TextArray(.Cols + 3) = "관리번호"
        .TextArray(.Cols + 4) = "OrderNo"
        .TextArray(.Cols + 5) = "가공구분"
        .TextArray(.Cols + 6) = "원단폭"
        .TextArray(.Cols + 7) = "축율"
        .TextArray(.Cols + 8) = "수주량"
        .TextArray(.Cols + 9) = "단위"
        .TextArray(.Cols + 10) = "절수"
        .TextArray(.Cols + 11) = "수량"
        .TextArray(.Cols + 12) = "절수"
        .TextArray(.Cols + 13) = "수량"
        .TextArray(.Cols + 14) = "절수"
        .TextArray(.Cols + 15) = "수량"
        .TextArray(.Cols + 16) = "절수"
        .TextArray(.Cols + 17) = "수량"
        .TextArray(.Cols + 18) = "절수"
        .TextArray(.Cols + 19) = "수량"
        .TextArray(.Cols + 20) = "절수"
        .TextArray(.Cols + 21) = "수량"
        .TextArray(.Cols + 22) = "수정구분"
        .TextArray(.Cols + 23) = "거래처코드"
        .TextArray(.Cols + 24) = "품명코드"
        .TextArray(.Cols + 25) = "가공구분"
        
        .ColFormat(8) = "#,##0"
        .ColFormat(10) = "#,##0"
        .ColFormat(11) = "#,##0"
        .ColFormat(12) = "#,##0"
        .ColFormat(13) = "#,##0"
        .ColFormat(14) = "#,##0"
        .ColFormat(15) = "#,##0"
        .ColFormat(16) = "#,##0"
        .ColFormat(17) = "#,##0"
        .ColFormat(18) = "#,##0"
        .ColFormat(19) = "#,##0"
        .ColFormat(20) = "#,##0"
        .ColFormat(21) = "#,##0"
        
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
        For i = 0 To 9
            .MergeCol(i) = True
        Next i
        
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub FillGridData()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset
    Dim i%, sOrderID$, bFlag As Boolean

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon

    Set rs = oSubul.GetSubul(MakeDate(DF_SHORT, dtpDate(0)), _
                        IIf(chkSearch(1), 1, 0), txtSearch(1).Tag, _
                        IIf(chkSearch(2), 1, 0), txtSearch(2).Tag, _
                        IIf(chkSearch(3), IIf(optOrder(0), 2, 1), 0), txtSearch(2))
    Set oSubul = Nothing

    With grdData
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            DoEvents

            .AddItem CStr(.Rows - 1) & vbTab & rs!KCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!OrderNo & vbTab & rs!workname & vbTab & "" & vbTab & rs!ChunkRate & vbTab & rs!OrderQty & vbTab & _
                IIf(rs!UnitClss = "0", "Y", "M") & vbTab & _
                rs!OverRoll & vbTab & rs!OverQty & vbTab & rs!InRoll & vbTab & rs!InQty & vbTab & _
                rs!OutRoll & vbTab & rs!OutQty & vbTab & _
                rs!OverRoll + rs!InRoll - rs!OutRoll + rs!SetRoll & vbTab & rs!OverQty + rs!InQty - rs!OutQty + rs!SetQty & vbTab & _
                rs!SetRoll & vbTab & rs!SetQty & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & rs!CustomID & vbTab & rs!ArticleID & vbTab & rs!workid

            If rs!OrderID <> sOrderID Then
                bFlag = Not bFlag
            End If

            If bFlag Then
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HE0E0E0
            Else
                .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 0
            End If

            sOrderID = rs!OrderID
            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
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
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmSubul.FillGridData", Err.Description)
End Sub


Private Function SaveData() As Boolean
    Dim tItem() As PlusLib2.TSubul
    Dim oSulBul As PlusLib2.CSubul
    Dim i%, iSeq%
    
    SaveData = False

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    ReDim tItem(0)
    With grdData
        For i = grdData.FixedRows To grdData.Rows - 1
            If .TextMatrix(i, 22) = "*" And (CheckNum(.TextMatrix(i, 20)) <> 0 Or CheckNum(.TextMatrix(i, 21)) <> 0) Then
                ReDim Preserve tItem(iSeq)
                tItem(iSeq).subuldate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6) + "99"
                tItem(iSeq).SubulClss = "1"
                tItem(iSeq).IOclss = "9"
                tItem(iSeq).OrderNo = .TextMatrix(i, 4)
                tItem(iSeq).OrderSeq = 0
                tItem(iSeq).CustomID = .TextMatrix(i, 23)
                tItem(iSeq).ArticleID = .TextMatrix(i, 24)
                tItem(iSeq).QtyUnit = 0
                tItem(iSeq).Custom = ""
                tItem(iSeq).workid = .TextMatrix(i, 25)
                tItem(iSeq).Price = 0
                tItem(iSeq).cnt = CSng(CheckNum(.TextMatrix(i, 18))) + CSng(CheckNum(.TextMatrix(i, 20)))
                tItem(iSeq).Qty = CSng(CheckNum(.TextMatrix(i, 19))) + CSng(CheckNum(.TextMatrix(i, 21)))
                tItem(iSeq).OutRealQty = .TextMatrix(i, 21)
                tItem(iSeq).Price = 0
                tItem(iSeq).OrderID = MakeOrderID(.TextMatrix(i, 3), OM_REDUCE)
                
                iSeq = iSeq + 1
            End If
        Next i
    End With

    If iSeq < 1 Then Exit Function
    
    Set oSulBul = New PlusLib2.CSubul
    oSulBul.Connection = g_adoCon
    oSulBul.UserName = g_sUserName
    
    SaveData = oSulBul.UpdateSubulSet(tItem)
    Set oSulBul = Nothing
    
    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    Set oSulBul = Nothing
    Call ErrorBox(Err.Number, "frmSubul.SaveData", Err.Description)
End Function
