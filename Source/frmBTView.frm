VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBTView 
   ClientHeight    =   9255
   ClientLeft      =   3240
   ClientTop       =   2790
   ClientWidth     =   11865
   Icon            =   "frmBTView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   8775
      TabIndex        =   31
      Top             =   90
      Width           =   1575
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   45
      TabIndex        =   28
      Top             =   8610
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   196609
      Caption         =   " 색상 세부내역 "
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "보임"
         Value           =   -1
      End
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   1
         Left            =   1035
         TabIndex        =   30
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "숨김"
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   2
      Left            =   5250
      TabIndex        =   7
      Top             =   795
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   5250
      TabIndex        =   6
      Top             =   450
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   0
      Left            =   5250
      TabIndex        =   5
      Top             =   105
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   930
      Left            =   10860
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   2
      ToolTipText     =   "검색"
      Top             =   105
      Width           =   900
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   0
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   465
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   45
      MousePointer    =   99  '사용자 정의
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8580
      TabIndex        =   3
      Top             =   8535
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   10230
      TabIndex        =   4
      Top             =   8535
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   8085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   4
      Left            =   3885
      TabIndex        =   8
      Top             =   450
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품      명"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   1155
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   0
      Left            =   6855
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   9
      Left            =   3885
      TabIndex        =   11
      Top             =   105
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "거 래 처"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   2445
      TabIndex        =   13
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Format          =   65929217
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   2445
      TabIndex        =   14
      Top             =   480
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      Format          =   65929217
      CurrentDate     =   36871
   End
   Begin VSFlex7LCtl.VSFlexGrid grdBt 
      Height          =   7335
      Left            =   30
      TabIndex        =   15
      Top             =   1155
      Width           =   11805
      _cx             =   20823
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
      Begin VSFlex7LCtl.VSFlexGrid grdBtShow 
         Height          =   2355
         Left            =   7275
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   3840
         _cx             =   6773
         _cy             =   4154
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
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   17
      Top             =   0
      Width           =   0
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   810
      TabIndex        =   18
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T 접수일자"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   19
         Top             =   60
         Value           =   1  '확인
         Width           =   1425
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   8
      Left            =   810
      TabIndex        =   20
      Top             =   795
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T 발송일자"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   21
         Top             =   60
         Width           =   1410
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   2
      Left            =   2445
      TabIndex        =   22
      Top             =   810
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   65929217
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   10
      Left            =   3885
      TabIndex        =   23
      Top             =   795
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "실 험 자"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   1140
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   1
      Left            =   6855
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   450
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   6930
      TabIndex        =   26
      Top             =   8535
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   2
      Left            =   6855
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   795
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   7260
      TabIndex        =   32
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "B/T NO"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   33
         Top             =   60
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmBTView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\BtList.rpt"

Private Const LIMIT_ROW1 = 25
Private Const LIMIT_ROW2 = 25
Private Const LIMIT_ROW3 = 5
Private Const LIMIT_ROW4 = 11
Private Const LIMIT_ROW5 = 10
Private Const LIMIT_WIDTH1 = 1380
Private Const LIMIT_WIDTH2 = 1635
Private Const LIMIT_WIDTH3 = 1965
Private Const LIMIT_WIDTH4 = 2085
Private Const LIMIT_WIDTH5 = 1890
Private Const LIMIT_WIDTH6 = 1000

Private m_sFlag         As String
Private m_nSelected     As Integer
Private m_bloading      As Boolean
Private m_bSortForward  As Boolean
Private m_bSaved        As Boolean
Private m_sBtID         As String
Private m_nBtSeq        As Integer


Private Sub cmdExcel_Click()
    If grdBt.Rows = 1 Then
        MsgBox LoadResString(203), vbInformation
        Exit Sub
    End If
    Call MakeExcelGrid(grdBt)
End Sub



Private Sub Form_Load()
    Me.Move 0, 0, 11985, 9660
    Dim i%
    
    Call SetOperate(Me)
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    dtpDate(2) = Now
    
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
   
    For i = 0 To 2
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i

    Call InitGrid
    
    Show

    txtSearch(0).Enabled = False
    txtSearch(1).Enabled = False
    cmdFind(0).Enabled = False
    cmdFind(1).Enabled = False
    cmdFind(2).Enabled = False
    
 '   Call FillGridBt
End Sub

Private Sub chkSearch_Click(Index As Integer)
    
    If Index > 2 Then
        If Index = 3 Then
            If chkSearch(3).Value = vbChecked Then
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            Else
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            End If
        ElseIf Index = 4 Then
            If chkSearch(4).Value = vbChecked Then
                dtpDate(2).Enabled = True
            Else
                dtpDate(2).Enabled = False
            End If
        
        ElseIf Index = 5 Then
            If chkSearch(5).Value = vbChecked Then
                txtSearch(3).Enabled = True
            Else
                txtSearch(3).Enabled = False
            End If
        End If
    Else
        If chkSearch(Index) Then
            If Index = 0 Then
                cmdFind(0).Enabled = True
                txtSearch(0).Enabled = True
                txtSearch(0).SetFocus
            ElseIf Index = 1 Then
                cmdFind(1).Enabled = True
                txtSearch(1).Enabled = True
                txtSearch(1).SetFocus
            ElseIf Index = 2 Then
                cmdFind(2).Enabled = True
                txtSearch(2).Enabled = True
                txtSearch(2).SetFocus
            ElseIf Index = 5 Then
                txtSearch(3).Enabled = True
                txtSearch(3).SetFocus
            End If
        Else
            If Index = 0 Then
                cmdFind(0).Enabled = False
                txtSearch(0).Enabled = False
            ElseIf Index = 1 Then
                cmdFind(1).Enabled = False
                txtSearch(1).Enabled = False
            ElseIf Index = 2 Then
                cmdFind(2).Enabled = False
                txtSearch(2).Enabled = False
            ElseIf Index = 5 Then
                txtSearch(3).Enabled = False
    
            End If
        End If
    End If
End Sub



Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then       ' 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then   ' 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If

    cmdSearch.SetFocus
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub


Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
        Case 1
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
        Case 2
            Call ReturnCode(LG_PERSON, , False, txtSearch(2))
  
    End Select
End Sub

Private Sub cmdSearch_Click()
    Call FillGridBt
End Sub


Private Sub grdBt_DblClick()
    With grdBt
        If .MouseRow < .FixedRows Then Exit Sub
        
        .Row = .MouseRow
        If .IsSubtotal(.Row) Then Exit Sub

    End With
End Sub


Private Sub grdBt_RowColChange()
    If m_bloading Then Exit Sub

    If optView(1).Value = True Then
        grdBtShow.Visible = False
        Exit Sub
    Else
        grdBtShow.Visible = True
    
    End If
    
    With grdBt
        If .IsSubtotal(.Row) = True Then
            grdBtShow.Visible = False
            Exit Sub
        Else
            grdBtShow.Visible = True
        End If
    
        If .Row < .FixedRows Or .Row >= .Rows Or .Rows = .FixedRows Then Exit Sub

        Call ShowBTData

        .SetFocus
    End With
End Sub




Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyReturn Then Exit Sub
    
    If Index = 3 Then
        Call cmdFind_Click(1)
    ElseIf Index = 4 Then
        Call cmdFind_Click(2)
    
    End If
    
End Sub

Private Sub optView_Click(Index As Integer, Value As Integer)
    
    With grdBt
    
        If .Rows = .FixedRows Then Exit Sub
        
        If .IsSubtotal(.Row) = True Then Exit Sub
    
    End With
    
    If Index = 1 Then
        grdBtShow.Visible = False
    Else
        grdBtShow.Visible = True
        Call ShowBTData
    End If
End Sub



Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub



Private Sub chkSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub


Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(0))
        ElseIf Index = 1 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_PERSON, , False, txtSearch(2))
        End If
        
        cmdSearch.SetFocus
    End If
End Sub



Private Sub cmdPrint_Click()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim nChkDate%, sDate$, eDate$
    Dim nChkSendDate%, SendDate$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim nChkPerson%, sPersonID$
    Dim nChkBTNO%, sBTNO$
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    If grdBt.Rows = grdBt.FixedRows Then Exit Sub
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    nChkDate = IIf(chkSearch(3), 1, 0)          ' 접수일
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkSendDate = IIf(chkSearch(4), 1, 0)      ' 발송일
    SendDate = MakeDate(DF_SHORT, dtpDate(2))
    nChkCustom = IIf(chkSearch(0), 1, 0)        ' 거래처
    sCustom = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(1), 1, 0)       ' 품명
    sArticle = txtSearch(1).Tag
    nChkPerson = IIf(chkSearch(2), 1, 0)        ' 작성자
    sPersonID = txtSearch(2).Tag
    nChkBTNO = IIf(chkSearch(5), 1, 0)
    sBTNO = txtSearch(3)
    
    Set rs = oBt.GetBtList(nChkDate, sDate, eDate, nChkSendDate, SendDate, nChkCustom, sCustom, _
            nChkArticle, sArticle, nChkPerson, sPersonID, nChkBTNO, sBTNO)
    
    Set oBt = Nothing
    
    
    ReDim sParam(1)
    sParam(0) = "B/T 접수대장"
    sParam(1) = CompanyName
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdBt
        .Cols = 18
        .Rows = 1
        
        .Redraw = flexRDNone
        
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarVertical
        .ExplorerBar = flexExSortShow
        .ScrollTrack = True
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 550
        .ColWidth(0) = 360
        .RowHeightMin = 450

             .TextArray(1) = "거래처":                       .ColWidth(1) = 2500:    .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "접수번호" & vbCrLf & "차수":   .ColWidth(2) = 1600:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "B/T NO" & vbCrLf & "접수일자": .ColWidth(3) = 2500:    .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "품명" & vbCrLf & "발송일자":   .ColWidth(4) = 3000:    .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "색상수":                       .ColWidth(5) = 1000:     .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "실험자":                       .ColWidth(6) = 900:     .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "접수" & vbCrLf & "등록일":     .ColWidth(7) = 1100:    .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(8) = "접수" & vbCrLf & "작성자":     .ColWidth(8) = 1100:    .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(9) = "발송" & vbCrLf & "등록일":     .ColWidth(9) = 1100:    .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(10) = "발송" & vbCrLf & "작성자":    .ColWidth(10) = 900:    .ColAlignment(9) = flexAlignCenterCenter
        .TextArray(11) = "거래처":                      .ColWidth(11) = 0
        .TextArray(12) = "거래처ID":                    .ColWidth(12) = 0
        .TextArray(13) = "BTID":                        .ColWidth(13) = 0
        .TextArray(14) = "BTNO":                        .ColWidth(14) = 0
        .TextArray(15) = "품명":                        .ColWidth(15) = 0
        .TextArray(16) = "품명ID":                      .ColWidth(16) = 0
        .TextArray(17) = "발송작성자ID":                .ColWidth(17) = 0

        For i = 6 To 10
            .ColHidden(i) = True
            
        Next i
        
        .Redraw = flexRDDirect
    End With

    With grdBtShow
        .Cols = 2
        Call SetVSFlexGrid(grdBtShow)

        .Redraw = False

        .TextArray(1) = "색상명":     .ColWidth(1) = 900:             .ColAlignment(1) = flexAlignLeftCenter

        .Redraw = True
    End With
    

End Sub


Private Sub FillGridBt()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, iNowRow%
    Dim nChkDate%, sDate$, eDate$
    Dim nChkSendDate%, SendDate$
    Dim nChkCustom%, sCustom$
    Dim nChkArticle%, sArticle$
    Dim nChkPerson%, sPersonID$
    Dim nChkBTNO%, sBTNO$
    Dim sPreBTID$, nCnt%, nBeforeTop%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    nChkDate = IIf(chkSearch(3), 1, 0)          ' 접수일
    sDate = MakeDate(DF_SHORT, dtpDate(0))
    eDate = MakeDate(DF_SHORT, dtpDate(1))
    nChkSendDate = IIf(chkSearch(4), 1, 0)      ' 발송일
    SendDate = MakeDate(DF_SHORT, dtpDate(2))
    nChkCustom = IIf(chkSearch(0), 1, 0)        ' 거래처
    sCustom = txtSearch(0).Tag
    nChkArticle = IIf(chkSearch(1), 1, 0)       ' 품명
    sArticle = txtSearch(1).Tag
    nChkPerson = IIf(chkSearch(2), 1, 0)        ' 작성자
    sPersonID = txtSearch(2).Tag
    nChkBTNO = IIf(chkSearch(5), 1, 0)        ' BTID
    sBTNO = txtSearch(3)
    
    Set rs = oBt.GetBtList(nChkDate, sDate, eDate, nChkSendDate, SendDate, nChkCustom, sCustom, _
                nChkArticle, sArticle, nChkPerson, sPersonID, nChkBTNO, sBTNO)
    
    Set oBt = Nothing

    nCnt = 1
    
    With grdBt
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            
            If sPreBTID <> rs!BTID Then
                sPreBTID = rs!BTID
                
                .AddItem CStr(nCnt) & vbTab & rs!KCustom & vbTab & MakeBTID(rs!BTID, OM_EXPAND) & vbTab & rs!BTNO & vbTab & rs!Article
            
                Call DoFlexGridGroup(.Rows - 1, 1)  ' 그리드 서브토탈
'                Call GridCollapse(nBeforeTop)       ' 서브토탈 row를 접힌 상태로 출력
'
                nBeforeTop = .Rows - 1
                nCnt = nCnt + 1
            End If
            
            .AddItem "" & vbTab & " " & vbTab & rs!BTIDSeq & vbTab & MakeDate(DF_LONG, CheckNull(rs!RecpDate)) & vbTab & MakeDate(DF_LONG, CheckNull(rs!SendDate)) & vbTab & _
                        rs!ColorCnt & vbTab & CheckNull(rs!Name) & vbTab & MakeDate(DF_LONG, CheckNull(rs!RecpDTime)) & vbTab & CheckNull(rs!RecpName) & vbTab & _
                        MakeDate(DF_LONG, CheckNull(rs!SendDTime)) & vbTab & CheckNull(rs!SendName) & vbTab & CheckNull(rs!KCustom) & vbTab & CheckNull(rs!CustomID) & vbTab & _
                        rs!BTID & vbTab & rs!BTNO & vbTab & CheckNull(rs!Article) & vbTab & CheckNull(rs!ArticleID) & vbTab & CheckNull(rs!SendPerID)
                        
            
            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0
            End If

            rs.MoveNext
        Next i

        rs.Close
        Set rs = Nothing

        .Redraw = flexRDDirect
        
        If .Rows > .FixedRows Then
            cmdPrint.Enabled = True
            .HighLight = flexHighlightAlways

            If m_bSaved = True Then
                Call FindNewRow(m_sBtID, m_nBtSeq)
            End If

        Else
            cmdPrint.Enabled = False
            .HighLight = flexHighlightNever

            grdBtShow.Visible = False
            MsgBox LoadResString(203), vbInformation
        End If

        .SetFocus
    End With

    Screen.MousePointer = vbDefault

    m_bSaved = False
    m_bloading = False

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oBt = Nothing

    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub DoFlexGridGroup(Row As Integer, Level As Integer)
    With grdBt
        ' Set the row as a group
        .IsSubtotal(Row) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(Row) = 1

        Select Case Level
            Case 0
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = &HE0E0E0
                .Cell(flexcpFontBold, Row, 0, Row, .Cols - 1) = True
            Case 1, 2
                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = &HE0E0E0
        End Select
        
        
    End With
End Sub

Private Sub GridCollapse(Row As Integer)
    
    With grdBt
    
        If Row >= .FixedRows Then
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub


Private Function MakeBTID(sBTID As String, nType As EORDERMAKE) As String
     If nType = OM_EXPAND Then
        MakeBTID = Left(sBTID, 2) & "-" & Mid(sBTID, 3, 2) & "-" & Mid(sBTID, 5, 4)
    Else
        MakeBTID = Replace(sBTID, "-", "")
    End If
    

End Function

Private Sub FindNewRow(sBTID As String, nSeq As Integer)
    Dim i%
    
    With grdBt
        For i = .FixedRows To .Rows - 1
            If .IsSubtotal(i) = False Then
                If (.TextMatrix(i, 13) = sBTID) And (.TextMatrix(i, 2) = nSeq) Then
                    .Row = i
                    .TopRow = i
                    Exit Sub
                End If
            End If
        Next i
    
    End With

End Sub




Private Sub ShowBTData()
    Dim oBt As PlusLib2.CBt
    Dim rs As Recordset
    Dim i%, sBTID$, nReworkSeq%, nBTSeq%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oBt = New PlusLib2.CBt
    oBt.Connection = g_adoCon

    With grdBt
        sBTID = .TextMatrix(.Row, 13)
        nBTSeq = .TextMatrix(.Row, 2)
    End With

    Set rs = oBt.GetBtSub(sBTID, nBTSeq)
    With grdBtShow
        .Redraw = False

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!Color

            rs.MoveNext
        Next i

        
        If .Rows > .FixedRows Then
            
            If .Rows < LIMIT_ROW5 Then
                .Height = (.RowHeight(.FixedRows) + 40) * .Rows + 350
                .ScrollBars = flexScrollBarNone
            Else
                .Height = 2700
                .ScrollBars = flexScrollBarVertical
            End If
        Else
            .Height = .RowHeight(0) + 110
        End If
        
        .Redraw = True
        .SetFocus
    End With
    
    With grdBt
        If .Rows = .FixedRows Then Exit Sub

        If .Row < (.TopRow + 7) Then
            grdBtShow.Top = 4400
        Else
            grdBtShow.Top = 900
        End If
    End With

    rs.Close

    Set rs = Nothing
    Set oBt = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oBt = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub





