VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCardHistory 
   Caption         =   "카드별 진행"
   ClientHeight    =   9315
   ClientLeft      =   1545
   ClientTop       =   3465
   ClientWidth     =   15240
   Icon            =   "frmCardHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15240
   Begin VB.Frame fraSearch 
      Height          =   585
      Left            =   0
      TabIndex        =   5
      Top             =   -90
      Width           =   4230
      Begin MSMask.MaskEdBox txtCard 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin Threed.SSPanel pnlName 
         Height          =   435
         Index           =   4
         Left            =   60
         TabIndex        =   6
         Top             =   120
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   767
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "카드 번호"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13545
      TabIndex        =   1
      Top             =   30
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
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
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   -30
      TabIndex        =   2
      Top             =   510
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15478
      _Version        =   393216
      TabHeight       =   600
      TabMaxWidth     =   4410
      TabCaption(0)   =   "카드 이력"
      TabPicture(0)   =   "frmCardHistory.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdCardModiList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdCardHist"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmCardHistory.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmCardHistory.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VSFlex7LCtl.VSFlexGrid grdCardHist 
         Height          =   5205
         Left            =   60
         TabIndex        =   3
         Top             =   420
         Width           =   15150
         _cx             =   26723
         _cy             =   9181
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
         ScrollBars      =   2
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
      Begin VSFlex7LCtl.VSFlexGrid grdCardModiList 
         Height          =   3105
         Left            =   60
         TabIndex        =   4
         Top             =   5640
         Width           =   15150
         _cx             =   26723
         _cy             =   5477
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
         ScrollBars      =   2
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "■  카드번호 8자리만 입력합니다(분할번호 제외)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   4935
      TabIndex        =   7
      Top             =   180
      Width           =   4335
   End
End
Attribute VB_Name = "frmCardHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15300, 9660
    
    ' 차후에 쓰일것임..
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    
    Call SetOperate(Me)
    Call InitGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    Call SetVSFlexGrid(grdCardHist)
    With grdCardHist
        .Redraw = flexRDNone

        .Rows = 4:          .Cols = 17
        .FixedRows = 4:     .FixedCols = 0
        
        .RowHeightMin = 0
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True

        .RowHeight(3) = 350
        .TextMatrix(3, 0) = "카드번호":     .ColWidth(0) = 1500:        .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(3, 1) = "분할":         .ColWidth(1) = 230:         .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(3, 2) = "분할":         .ColWidth(2) = 230:         .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(3, 3) = "분할":         .ColWidth(3) = 230:         .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "분할":         .ColWidth(4) = 230:         .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(3, 5) = "상태":         .ColWidth(5) = 500:         .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(3, 6) = "공정명":       .ColWidth(6) = 1200:         .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "관리번호":     .ColWidth(7) = 1300:        .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(3, 8) = "거래처":       .ColWidth(8) = 1300:        .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "품명":         .ColWidth(9) = 2200:        .ColAlignment(9) = flexAlignLeftCenter
        .TextMatrix(3, 10) = "Order No.":   .ColWidth(10) = 1200:      .ColAlignment(10) = flexAlignLeftCenter
        .TextMatrix(3, 11) = "색상명":      .ColWidth(11) = 1300:       .ColAlignment(11) = flexAlignLeftCenter
        .TextMatrix(3, 12) = "절수":        .ColWidth(12) = 700:        .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(3, 13) = "수량":        .ColWidth(13) = 700:        .ColAlignment(13) = flexAlignRightCenter
        .TextMatrix(3, 14) = "합격":        .ColWidth(14) = 700:        .ColAlignment(14) = flexAlignRightCenter
        .TextMatrix(3, 15) = "불합격":      .ColWidth(15) = 700:        .ColAlignment(15) = flexAlignRightCenter
        .TextMatrix(3, 16) = "출고":        .ColWidth(16) = 700:        .ColAlignment(16) = flexAlignRightCenter
        
        .MergeCells = flexMergeFree
        .MergeRow(3) = True
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
       
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .OutlineBar = flexOutlineBarSimple
'        .ScrollBars = flexScrollBarBoth
        .GridLines = flexGridNone
        
        .WordWrap = True
        .Redraw = flexRDDirect
    End With
    
    Call SetVSFlexGrid(grdCardModiList)
    With grdCardModiList
        .Redraw = flexRDNone

        .Rows = 4:          .Cols = 10
        .FixedRows = 4:     .FixedCols = 0
        
'        .RowHeightMin = 400
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .RowHeight(3) = 400

        .TextMatrix(3, 0) = "카드번호":         .ColWidth(0) = 1500:        .ColAlignment(0) = flexAlignLeftCenter
        .TextMatrix(3, 1) = "분할번호":         .ColWidth(1) = 0:           .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "순위":             .ColWidth(2) = 500:         .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "변경구분":         .ColWidth(3) = 1100:        .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "변경자":           .ColWidth(4) = 700:         .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(3, 5) = "변경일":           .ColWidth(5) = 700:         .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "변경사유":         .ColWidth(6) = 1000:        .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "변경전 계획공정":  .ColWidth(7) = 4800:        .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(3, 8) = "변경후 계획공정":  .ColWidth(8) = 4000:        .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "수정":             .ColWidth(9) = 1000:        .ColAlignment(9) = flexAlignLeftCenter
        
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
       
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
'        .OutlineBar = flexOutlineBarSimple
        .ScrollBars = flexScrollBarBoth
'        .GridLines = flexGridNone
        
        .WordWrap = True
        .Redraw = flexRDDirect
    End With
    
End Sub





Public Sub txtCard_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtCard)) > 0 And Len(Trim(txtCard)) = 8 Then
            Call FillGridCardHist
        Else
            MsgBox "카드번호가 잘못되었습니다", vbInformation, "Key 입력 오류"
            txtCard.SetFocus
        End If
    End If
End Sub

Private Sub FillGridCardHist()
Dim oCard As Pluslib2.CCard
Dim rs As Recordset
Dim iNowRow%, iRecCnt%, i%
Dim sSplitID(4) As String

    On Error GoTo ErrHandler



    Screen.MousePointer = vbHourglass
    
    Set oCard = New Pluslib2.CCard
    oCard.Connection = g_adoCon

    ' 카드 이력
    With grdCardHist
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        Set rs = oCard.GetCardAllList(txtCard)
        
        If rs.RecordCount > 0 Then
            For iRecCnt = 1 To rs.RecordCount
                .Rows = .Rows + 2
            
                .RowHeight(.Rows - 2) = 350
                Select Case Len(Trim(rs!SplitID))
                    Case 0:
                        sSplitID(0) = " ":   sSplitID(1) = " ":   sSplitID(2) = " ":   sSplitID(3) = " "
                        .IsSubtotal(.Rows - 2) = True
                        .RowOutlineLevel(.Rows - 2) = 0
                    Case 1:
                        sSplitID(0) = Trim(rs!SplitID):   sSplitID(1) = Trim(rs!SplitID):   sSplitID(2) = Trim(rs!SplitID):   sSplitID(3) = Trim(rs!SplitID)
                        .IsSubtotal(.Rows - 2) = True
                        .RowOutlineLevel(.Rows - 2) = 1
                    Case 2:
                        sSplitID(0) = " ":   sSplitID(1) = Trim(rs!SplitID):  sSplitID(2) = Trim(rs!SplitID):   sSplitID(3) = Trim(rs!SplitID)
                        .IsSubtotal(.Rows - 2) = True
                        .RowOutlineLevel(.Rows - 2) = 2
                    Case 3:
                        sSplitID(0) = " ":   sSplitID(1) = " ":   sSplitID(2) = Trim(rs!SplitID):   sSplitID(3) = Trim(rs!SplitID)
                        .IsSubtotal(.Rows - 2) = True
                        .RowOutlineLevel(.Rows - 2) = 3
                    Case 4:
                        sSplitID(0) = " ":   sSplitID(1) = " ":   sSplitID(2) = "":   sSplitID(3) = Trim(rs!SplitID)
                        .IsSubtotal(.Rows - 2) = True
                        .RowOutlineLevel(.Rows - 2) = 4
                End Select
            
                .TextMatrix(.Rows - 2, 0) = IIf(iRecCnt = 1, MakeCardID(rs!CardID, OM_EXPAND), " ")
                .TextMatrix(.Rows - 2, 1) = sSplitID(0)
                .TextMatrix(.Rows - 2, 2) = sSplitID(1)
                .TextMatrix(.Rows - 2, 3) = sSplitID(2)
                .TextMatrix(.Rows - 2, 4) = sSplitID(3)
                .TextMatrix(.Rows - 2, 5) = rs!UseClss
                Select Case rs!UseClss
                    Case "작업":
                            .Cell(flexcpBackColor, .Rows - 2, 5) = vbBlue
                            .Cell(flexcpForeColor, .Rows - 2, 5) = vbWhite
                    Case "보류":
                            .Cell(flexcpBackColor, .Rows - 2, 5) = vbRed
                            .Cell(flexcpForeColor, .Rows - 2, 5) = vbWhite
                End Select
                .TextMatrix(.Rows - 2, 6) = rs!Process
                
                .TextMatrix(.Rows - 2, 7) = MakeOrderID(rs!OrderID, OM_EXPAND)
                .TextMatrix(.Rows - 2, 8) = Trim(rs!kCustom)
                .TextMatrix(.Rows - 2, 9) = Trim(rs!Article)
                .TextMatrix(.Rows - 2, 10) = Trim(rs!OrderNo)
                .TextMatrix(.Rows - 2, 11) = Trim(rs!Color)
                .TextMatrix(.Rows - 2, 12) = SetCurrency(rs!Roll)
                .TextMatrix(.Rows - 2, 13) = SetCurrency(rs!Qty)
                .TextMatrix(.Rows - 2, 14) = SetCurrency(rs!okqty)
                .TextMatrix(.Rows - 2, 15) = SetCurrency(rs!noqty)
                .TextMatrix(.Rows - 2, 16) = SetCurrency(rs!OutQty)
                .IsSubtotal(.Rows - 2) = True
            
                .Cell(flexcpText, .Rows - 1, 5, .Rows - 1, .Cols - 1) = rs!cardproc
                .Cell(flexcpBackColor, .Rows - 1, 5, .Rows - 1, .Cols - 1) = &HE0E0E0

                .RowHeight(.Rows - 1) = 600
            
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 30
                .Cell(flexcpBackColor, .Rows - 1, 5, .Rows - 1, .Cols - 1) = vbBlue
                rs.MoveNext
            Next iRecCnt
            
            .Cell(flexcpFontBold, .FixedRows, 0, .Rows - 1, 4) = True
            .Cell(flexcpFontSize, .FixedRows, 0, .Rows - 1, 4) = 10
            .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 4) = &H8000000F
            .MergeCells = flexMergeFree
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
                .MergeCol(i) = True
            Next i
            For i = 0 To .Rows - 1
                .MergeRow(i) = True
            Next i
        Else
            MsgBox LoadResString(203), vbInformation
        End If
        Set rs = Nothing
        
        
        .Redraw = flexRDDirect
    End With
    
    
    
    ' 카드 변경 이력
    With grdCardModiList
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        Set rs = oCard.GetCardModiList(txtCard)
        
        If rs.RecordCount > 0 Then
            For iRecCnt = 1 To rs.RecordCount
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 450
            
'                .TextMatrix(.Rows - 1, 0) = IIf(Trim(rs!SplitID) = "", MakeCardID(rs!CardID, OM_EXPAND), MakeCardID(rs!CardID, OM_EXPAND) & "(" & rs!SplitID & ")")
                .TextMatrix(.Rows - 1, 0) = IIf(rs!ReWorkClss = "*", "■ ", " ") & MakeCardID(rs!CardID, OM_EXPAND, rs!SplitID)
                .TextMatrix(.Rows - 1, 1) = rs!SplitID
                .TextMatrix(.Rows - 1, 2) = IIf(rs!histseq = 0, rs!UseClss, rs!histseq)
                If rs!histseq = 0 Then
                    Select Case rs!UseClss
                        Case "작업":
                                .Cell(flexcpBackColor, .Rows - 1, 2) = vbBlue
                                .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite
                        Case "보류":
                                .Cell(flexcpBackColor, .Rows - 1, 2) = vbRed
                                .Cell(flexcpForeColor, .Rows - 1, 2) = vbWhite
                    End Select
                    .RowHeight(.Rows - 1) = 450
                    
                End If
                .TextMatrix(.Rows - 1, 3) = IIf(rs!histseq = 0, rs!Process, rs!ModiClss)
                .TextMatrix(.Rows - 1, 4) = rs!PersonName
                .TextMatrix(.Rows - 1, 5) = IIf(Format(rs!modidate, "YYYY") = "1900", "", MakeDate(DF_MD, rs!modidate))
                .TextMatrix(.Rows - 1, 6) = rs!modireason
                .TextMatrix(.Rows - 1, 7) = rs!preplanproc
                .TextMatrix(.Rows - 1, 8) = rs!postplanproc
                .TextMatrix(.Rows - 1, 9) = rs!ReWorkReason
            
                rs.MoveNext
            Next iRecCnt
        End If
        Set rs = Nothing
        .Redraw = flexRDDirect
    End With
    
    Set oCard = Nothing

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCard = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmCardHistory.FillGridCardHist", Err.Description)
End Sub
