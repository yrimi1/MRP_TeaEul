VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcWaiting 
   Caption         =   "공정별 작업 현황"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   3465
   ClientWidth     =   15240
   Icon            =   "frmProcWaiting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15240
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "보류 카드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   9
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "작업중 카드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
   End
   Begin Threed.SSPanel pnlProcess 
      Height          =   495
      Left            =   4890
      TabIndex        =   2
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   873
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkExpand 
         Caption         =   "후공정 리스트 보이기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   11
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox cboProcess 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1830
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   60
         Width           =   2475
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "■ 공정 선택"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   405
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   15372
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "현 공정대기 카드 리스트"
      TabPicture(0)   =   "frmProcWaiting.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pnlWaitTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdProcess(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "현 공정계획 카드 리스트"
      TabPicture(1)   =   "frmProcWaiting.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdProcess(1)"
      Tab(1).Control(1)=   "pnlWaitTab(1)"
      Tab(1).ControlCount=   2
      Begin VSFlex7LCtl.VSFlexGrid grdProcess 
         Height          =   8205
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   450
         Width           =   15045
         _cx             =   26538
         _cy             =   14473
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
         BackColorSel    =   16777215
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
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
         ExtendLastCol   =   -1  'True
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
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   60
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   14389120
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "현 공정대기 카드 리스트"
         BevelOuter      =   0
         FloodColor      =   14389120
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   375
         Index           =   1
         Left            =   -69870
         TabIndex        =   13
         Top             =   60
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   14389120
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "현 공정계획 카드 리스트"
         BevelOuter      =   0
         FloodColor      =   14389120
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdProcess 
         Height          =   8205
         Index           =   1
         Left            =   -74940
         TabIndex        =   14
         Top             =   450
         Width           =   15045
         _cx             =   26538
         _cy             =   14473
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
         BackColorSel    =   16777215
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
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
         ExtendLastCol   =   -1  'True
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
      Height          =   510
      Left            =   13590
      TabIndex        =   10
      Top             =   0
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   900
      _Version        =   196609
      Caption         =   "        닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmProcWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboProcess_Click()
    pnlWaitTab(0).Caption = "[" & Trim(cboProcess.Text) & "]" & " 공정대기 카드 리스트"
    SSTab1.TabCaption(0) = "[" & Trim(cboProcess.Text) & "]" & " 공정대기 카드 리스트"
    pnlWaitTab(1).Caption = "[" & Trim(cboProcess.Text) & "]" & " 공정계획 카드 리스트"
    SSTab1.TabCaption(1) = "[" & Trim(cboProcess.Text) & "]" & " 공정계획 카드 리스트"
    Call FillGridOrder
    grdProcess(SSTab1.Tab).SetFocus
End Sub

Private Sub chkExpand_Click()
Dim idx%
    For idx = 0 To 1
        With grdProcess(idx)
            If chkExpand.Value = 1 Then
                .ColWidth(10) = 8000
                .ScrollBars = flexScrollBarBoth
            Else
                .ColWidth(10) = 0
                .ScrollBars = flexScrollBarVertical
            End If
        End With
    Next idx
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15360, 9840
    
    Call SetOperate(Me)
    Call MakeProcessCombo
    Call InitGrid
    
End Sub

Private Sub MakeProcessCombo()
    Dim oProcess As PlusLib2.cprocess
    Dim rs As ADODB.Recordset
    Dim i%
    
    Set oProcess = New PlusLib2.cprocess
    oProcess.Connection = g_adoCon
    
    Set rs = oProcess.GetProcess()
    Set oProcess = Nothing
    
    
    With cboProcess
        .Clear
        For i = 0 To rs.RecordCount - 1

            .AddItem rs!Process
            .ItemData(.NewIndex) = rs!processid

            rs.MoveNext
        Next i
        cboProcess.ListIndex = -1

    End With

End Sub


Private Sub InitGrid()
    Dim i%, idx%
    
    For idx = 0 To 1
        With grdProcess(idx)
            .Redraw = flexRDNone
            
            .SelectionMode = flexSelectionFree
'            .FocusRect = flexFocusNone
            .ScrollBars = flexScrollBarBoth
            
            .Rows = 4:          .Cols = 28
            .FixedRows = 4:     .FixedCols = 0
            
            .RowHeight(0) = 0
            .RowHeight(1) = 0
            .RowHeight(2) = 0
            .RowHeight(3) = 400
    
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 0
            Next i
    
            .TextMatrix(3, 0) = "순번":             .ColWidth(0) = 600:     .ColAlignment(0) = flexAlignCenterCenter
            If idx = 0 Then
                .TextMatrix(3, 1) = "완료공정":         .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignLeftCenter
            Else
                .TextMatrix(3, 1) = "대기공정":         .ColWidth(1) = 1300:    .ColAlignment(1) = flexAlignLeftCenter
            End If
            .TextMatrix(3, 2) = "카드번호":         .ColWidth(2) = 1600:     .ColAlignment(2) = flexAlignLeftCenter
            .TextMatrix(3, 3) = "거래처":           .ColWidth(3) = 1500:    .ColAlignment(3) = flexAlignLeftCenter
            .TextMatrix(3, 4) = "품명":             .ColWidth(4) = 2900:     .ColAlignment(4) = flexAlignLeftCenter
            .TextMatrix(3, 5) = "색상":             .ColWidth(5) = 1800:     .ColAlignment(5) = flexAlignLeftCenter
            .TextMatrix(3, 6) = "관리번호":         .ColWidth(6) = 1500:    .ColAlignment(6) = flexAlignCenterCenter
            .TextMatrix(3, 7) = "Order No.":        .ColWidth(7) = 1700:    .ColAlignment(7) = flexAlignLeftCenter
            .TextMatrix(3, 8) = "절수":             .ColWidth(8) = 800:     .ColAlignment(8) = flexAlignRightCenter
            .TextMatrix(3, 9) = "수량":             .ColWidth(9) = 1000:    .ColAlignment(9) = flexAlignRightCenter
            .TextMatrix(3, 10) = "후공정 계획":     .ColWidth(10) = 0:      .ColAlignment(10) = flexAlignLeftCenter
            If idx = 0 Then
                .TextMatrix(3, 21) = "완료공정코드":    .ColWidth(21) = 0:      .ColAlignment(21) = flexAlignCenterCenter
            Else
                .TextMatrix(3, 21) = "대기공정코드":    .ColWidth(21) = 0:      .ColAlignment(21) = flexAlignCenterCenter
            End If
            .TextMatrix(3, 22) = "카드번호":        .ColWidth(22) = 0:      .ColAlignment(22) = flexAlignCenterCenter
            .TextMatrix(3, 23) = "분할번호":        .ColWidth(23) = 0:      .ColAlignment(23) = flexAlignCenterCenter
            .TextMatrix(3, 24) = "거래처코드":      .ColWidth(24) = 0:      .ColAlignment(24) = flexAlignCenterCenter
            .TextMatrix(3, 25) = "품명코드":        .ColWidth(25) = 0:      .ColAlignment(25) = flexAlignCenterCenter
            .TextMatrix(3, 26) = "색상코드":        .ColWidth(26) = 0:      .ColAlignment(26) = flexAlignCenterCenter
            .TextMatrix(3, 27) = "관리번호":        .ColWidth(27) = 0:      .ColAlignment(27) = flexAlignCenterCenter
            
            .ColFormat(8) = "##,##0"
            .ColFormat(9) = "##,##0"
            
            .Redraw = flexRDDirect
        End With
    Next idx
        
End Sub




Private Sub FillGridOrder()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    Dim nTotRoll As Long, nTotQty As Long
    Dim i%
    
    On Error GoTo ErrHandler
    
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetCardWaiting(cboProcess.ItemData(cboProcess.ListIndex), 0)
    Set oCard = Nothing
        
    With grdProcess(0)
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            
            .TextMatrix(.Rows - 1, 0) = CStr(i)
            .TextMatrix(.Rows - 1, 1) = rs!Process
            .TextMatrix(.Rows - 1, 2) = IIf(Len(Trim(rs!SplitID)) > 0, MakeCardID(rs!CardID, OM_EXPAND) & "(" & rs!SplitID & ")", MakeCardID(rs!CardID, OM_EXPAND))
            .TextMatrix(.Rows - 1, 3) = rs!KCustom
            .TextMatrix(.Rows - 1, 4) = rs!Article
            .TextMatrix(.Rows - 1, 5) = rs!Color
            .TextMatrix(.Rows - 1, 6) = MakeOrderID(rs!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 7) = rs!OrderNo
            .TextMatrix(.Rows - 1, 8) = rs!Roll
            .TextMatrix(.Rows - 1, 9) = rs!Qty
            .TextMatrix(.Rows - 1, 10) = rs!AfterProc
            
            .TextMatrix(.Rows - 1, 21) = rs!compprocid
            .TextMatrix(.Rows - 1, 22) = rs!CardID
            .TextMatrix(.Rows - 1, 23) = rs!SplitID
            .TextMatrix(.Rows - 1, 24) = rs!CustomID
            .TextMatrix(.Rows - 1, 25) = rs!ArticleID
            .TextMatrix(.Rows - 1, 26) = rs!colorid
            .TextMatrix(.Rows - 1, 27) = rs!OrderID
            
            nTotRoll = nTotRoll + rs!Roll
            nTotQty = nTotQty + rs!Qty
            
            Select Case Trim(rs!UseClss)
                Case "작업"
                    .Cell(flexcpBackColor, .Rows - 1, 2) = vbBlue
                    .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite
                Case "보류"
                    .Cell(flexcpBackColor, .Rows - 1, 2) = vbRed
                    .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite
            End Select
            rs.MoveNext
        Next i
        
        If chkExpand.Value = 1 Then
            .ColWidth(10) = 8000
            .ScrollBars = flexScrollBarBoth
        Else
            .ColWidth(10) = 0
            .ScrollBars = flexScrollBarVertical
        End If
        
        If rs.RecordCount > 0 Then
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 350
            
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 7) = " "
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .TextMatrix(.Rows - 1, 8) = nTotRoll
            .TextMatrix(.Rows - 1, 9) = nTotQty
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
        End If
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
        .Row = 0
'        .SetFocus
    End With
    
    
    nTotRoll = 0
    nTotQty = 0
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    Set rs = oCard.GetCardWaiting(cboProcess.ItemData(cboProcess.ListIndex), 1)
    Set oCard = Nothing
        
    With grdProcess(1)
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            
            .TextMatrix(.Rows - 1, 0) = CStr(i)
            .TextMatrix(.Rows - 1, 1) = rs!Process
            .TextMatrix(.Rows - 1, 2) = IIf(Len(Trim(rs!SplitID)) > 0, MakeCardID(rs!CardID, OM_EXPAND) & "(" & rs!SplitID & ")", MakeCardID(rs!CardID, OM_EXPAND))
            .TextMatrix(.Rows - 1, 3) = rs!KCustom
            .TextMatrix(.Rows - 1, 4) = rs!Article
            .TextMatrix(.Rows - 1, 5) = rs!Color
            .TextMatrix(.Rows - 1, 6) = MakeOrderID(rs!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 7) = rs!OrderNo
            .TextMatrix(.Rows - 1, 8) = rs!Roll
            .TextMatrix(.Rows - 1, 9) = rs!Qty
            .TextMatrix(.Rows - 1, 10) = rs!AfterProc
            
            .TextMatrix(.Rows - 1, 21) = rs!waitprocid
            .TextMatrix(.Rows - 1, 22) = rs!CardID
            .TextMatrix(.Rows - 1, 23) = rs!SplitID
            .TextMatrix(.Rows - 1, 24) = rs!CustomID
            .TextMatrix(.Rows - 1, 25) = rs!ArticleID
            .TextMatrix(.Rows - 1, 26) = rs!colorid
            .TextMatrix(.Rows - 1, 27) = rs!OrderID
            
            nTotRoll = nTotRoll + rs!Roll
            nTotQty = nTotQty + rs!Qty
            
            Select Case Trim(rs!UseClss)
                Case "작업"
                    .Cell(flexcpBackColor, .Rows - 1, 2) = vbBlue
                    .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite
                Case "보류"
                    .Cell(flexcpBackColor, .Rows - 1, 2) = vbRed
                    .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite
            End Select
            rs.MoveNext
        Next i
        
        If chkExpand.Value = 1 Then
            .ColWidth(10) = 8000
            .ScrollBars = flexScrollBarBoth
        Else
            .ColWidth(10) = 0
            .ScrollBars = flexScrollBarVertical
        End If
        
        If rs.RecordCount > 0 Then
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 350
            
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 7) = " "
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .TextMatrix(.Rows - 1, 8) = nTotRoll
            .TextMatrix(.Rows - 1, 9) = nTotQty
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
        End If
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
        .Row = 0
'        .SetFocus
    End With
    
    
    Exit Sub

ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmProcWorking.FillGridOrder", Err.Description)
End Sub

